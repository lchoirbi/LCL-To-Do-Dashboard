from __future__ import annotations

import hashlib
import os
import re
import time
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any

import dateparser
import msal
import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup
from dateparser.search import search_dates


APP_DIR = Path(__file__).parent
DATA_DIR = APP_DIR / "data"
OVERRIDES = DATA_DIR / "task_overrides.csv"
GRAPH = "https://graph.microsoft.com/v1.0"

STATUSES = ["to do", "in-progress", "complete"]
URGENCIES = ["High", "Medium", "Low"]
BANNERS = [
    "Unassigned",
    "Loblaws",
    "No Frills",
    "Real Canadian Superstore",
    "Zehrs",
    "Fortinos",
    "Independent",
    "Valu-mart",
    "Provigo",
    "Maxi",
    "Atlantic Superstore",
    "Dominion",
    "Wholesale Club",
    "Shoppers Drug Mart",
    "T&T",
]
BANNER_WORDS = {
    "Loblaws": ["loblaws", "loblaw", "lcl"],
    "No Frills": ["no frills", "nofrills"],
    "Real Canadian Superstore": ["real canadian superstore", "rcss", "superstore"],
    "Zehrs": ["zehrs"],
    "Fortinos": ["fortinos"],
    "Independent": ["your independent grocer", "independent"],
    "Valu-mart": ["valu-mart", "valu mart", "valumart"],
    "Provigo": ["provigo"],
    "Maxi": ["maxi"],
    "Atlantic Superstore": ["atlantic superstore"],
    "Dominion": ["dominion"],
    "Wholesale Club": ["wholesale club"],
    "Shoppers Drug Mart": ["shoppers drug mart", "shoppers", "sdm"],
    "T&T": ["t&t", "t and t", "tnt"],
}
BASE_SCOPES = ["User.Read", "Mail.Read", "Calendars.Read", "Chat.Read"]
CHANNEL_SCOPES = ["Team.ReadBasic.All", "Channel.ReadBasic.All", "ChannelMessage.Read.All"]

ACTION_RE = re.compile(
    r"\b(action required|action item|follow[- ]?up|following up|circle back|outstanding|pending|"
    r"blocked|blocker|need(?:ed|s)?|please|can you|could you|would you|will you|should|must|"
    r"send|share|provide|review|approve|confirm|align|update|submit|reconcile|resolve|"
    r"check|chase|close out|due|deadline|eta)\b",
    re.I,
)
URGENT_RE = re.compile(r"\b(urgent|asap|immediate|high priority|critical|eod|cob|today|tomorrow|escalat)\b", re.I)
DONE_RE = re.compile(r"^\s*(done|complete|completed|closed|resolved|sent|submitted)\b", re.I)
PROGRESS_RE = re.compile(r"\b(in progress|working on|underway|started)\b", re.I)
OWNER_RE = re.compile(r"\b(?:owner|owned by|assigned to|assignee)\s*[:\-]\s*([A-Z][A-Za-z.'-]+(?:\s+[A-Z][A-Za-z.'-]+){0,2})")


class GraphError(RuntimeError):
    pass


def secret(name: str, default: str | None = None) -> str | None:
    try:
        return str(st.secrets[name])
    except Exception:
        return os.environ.get(name, default)


def csv(value: str | None, lower: bool = False) -> list[str]:
    parts = [part.strip() for part in (value or "").split(",") if part.strip()]
    return [part.lower().lstrip("@") for part in parts] if lower else parts


def clean(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def html(value: str | None) -> str:
    return BeautifulSoup(value or "", "html.parser").get_text(" ", strip=True)


def dt(value: str | None) -> datetime | None:
    if not value:
        return None
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        parsed = dateparser.parse(value)
        return parsed.replace(tzinfo=timezone.utc) if parsed and parsed.tzinfo is None else parsed


def graph_dt(value: datetime) -> str:
    return value.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def deep(mapping: dict[str, Any] | None, *keys: str) -> Any:
    cur: Any = mapping or {}
    for key in keys:
        if not isinstance(cur, dict):
            return None
        cur = cur.get(key)
    return cur


def is_token_valid(result: dict[str, Any] | None) -> bool:
    return bool(result and result.get("access_token") and int(result.get("expires_on", 0)) > time.time() + 60)


def microsoft_token(client_id: str | None, tenant_id: str, scopes: list[str]) -> str | None:
    if is_token_valid(st.session_state.get("token")):
        return st.session_state["token"]["access_token"]
    if not client_id:
        st.sidebar.error("Add AZURE_CLIENT_ID to Streamlit secrets to enable Microsoft Graph refresh.")
        return None

    app = msal.PublicClientApplication(client_id, authority=f"https://login.microsoftonline.com/{tenant_id}")
    flow = st.session_state.get("device_flow")
    if st.sidebar.button("Sign in to Microsoft", use_container_width=True):
        flow = app.initiate_device_flow(scopes=scopes)
        st.session_state["device_flow"] = flow
    if flow and "user_code" in flow:
        st.sidebar.link_button("Open Microsoft sign-in", flow["verification_uri"], use_container_width=True)
        st.sidebar.code(flow["user_code"])
        if st.sidebar.button("I completed sign-in", use_container_width=True):
            result = app.acquire_token_by_device_flow(flow)
            if "access_token" in result:
                st.session_state["token"] = result
                st.session_state.pop("device_flow", None)
                st.rerun()
            st.sidebar.error(result.get("error_description", "Microsoft sign-in failed."))
    return None


class GraphClient:
    def __init__(self, token: str) -> None:
        self.headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
            "Prefer": 'outlook.body-content-type="text"',
        }

    def get(self, path: str, params: dict[str, Any] | None = None) -> dict[str, Any]:
        url = path if path.startswith("https://") else f"{GRAPH}{path}"
        res = requests.get(url, headers=self.headers, params=params, timeout=45)
        if res.status_code >= 400:
            raise GraphError(f"{res.status_code}: {res.text[:500]}")
        return res.json()

    def list(self, path: str, params: dict[str, Any] | None = None, limit: int = 100) -> list[dict[str, Any]]:
        out: list[dict[str, Any]] = []
        url = path if path.startswith("https://") else f"{GRAPH}{path}"
        first = True
        while url and len(out) < limit:
            res = requests.get(url, headers=self.headers, params=params if first else None, timeout=45)
            first = False
            if res.status_code >= 400:
                raise GraphError(f"{res.status_code}: {res.text[:500]}")
            payload = res.json()
            out.extend(payload.get("value", []))
            url = payload.get("@odata.nextLink")
        return out[:limit]


def rec(kind: str, item: dict[str, Any], title: str, when: datetime | None, sender: dict[str, str], body: str) -> dict[str, Any]:
    return {
        "source_type": kind,
        "source_id": item.get("id", ""),
        "source_title": clean(title),
        "source_date": when,
        "from_name": sender.get("name") or sender.get("displayName") or "",
        "from_email": sender.get("address") or sender.get("userIdentityType") or "",
        "importance": item.get("importance", "normal"),
        "web_link": item.get("webLink") or item.get("webUrl") or "",
        "text": clean(f"{title}. {body}"),
    }


def fetch_inbox(client: GraphClient, start: datetime, limit: int) -> list[dict[str, Any]]:
    rows = client.list(
        "/me/mailFolders/inbox/messages",
        {
            "$top": min(limit, 50),
            "$filter": f"receivedDateTime ge {graph_dt(start)}",
            "$orderby": "receivedDateTime desc",
            "$select": "id,subject,from,receivedDateTime,bodyPreview,body,importance,webLink",
        },
        limit,
    )
    return [
        rec(
            "Outlook Inbox",
            row,
            row.get("subject", "(no subject)"),
            dt(row.get("receivedDateTime")),
            deep(row, "from", "emailAddress") or {},
            html(deep(row, "body", "content") or row.get("bodyPreview", "")),
        )
        for row in rows
    ]


def fetch_calendar(client: GraphClient, start: datetime, end: datetime, limit: int) -> list[dict[str, Any]]:
    rows = client.list(
        "/me/calendarView",
        {
            "startDateTime": graph_dt(start),
            "endDateTime": graph_dt(end),
            "$top": min(limit, 50),
            "$orderby": "start/dateTime desc",
            "$select": "id,subject,organizer,attendees,start,bodyPreview,body,importance,webLink,isCancelled",
        },
        limit,
    )
    records = []
    for row in rows:
        if row.get("isCancelled"):
            continue
        attendees = ", ".join(
            clean(deep(attendee, "emailAddress", "name"))
            for attendee in row.get("attendees", []) or []
            if deep(attendee, "emailAddress", "name")
        )
        body = f"{html(deep(row, 'body', 'content') or row.get('bodyPreview', ''))}. Attendees: {attendees}"
        records.append(
            rec(
                "Outlook Calendar",
                row,
                row.get("subject", "(no subject)"),
                dt(deep(row, "start", "dateTime")),
                deep(row, "organizer", "emailAddress") or {},
                body,
            )
        )
    return records


def mentions_me(message: dict[str, Any], user_id: str | None) -> bool:
    return bool(user_id and any(deep(m, "mentioned", "user", "id") == user_id for m in message.get("mentions", []) or []))


def fetch_chats(client: GraphClient, start: datetime, user_id: str | None, chats: int, per_chat: int) -> list[dict[str, Any]]:
    records = []
    for chat in client.list("/me/chats", {"$top": min(chats, 50)}, chats):
        try:
            messages = client.list(f"/chats/{chat['id']}/messages", {"$top": min(per_chat, 50)}, per_chat)
        except GraphError:
            continue
        topic = chat.get("topic") or chat.get("chatType") or "Teams chat"
        for msg in messages:
            when = dt(msg.get("createdDateTime"))
            body = html(deep(msg, "body", "content"))
            if (when and when < start) or not body:
                continue
            if not mentions_me(msg, user_id) and not ACTION_RE.search(body):
                continue
            records.append(rec("Teams Chat", msg, topic, when, deep(msg, "from", "user") or {}, body) | {"mentions_me": mentions_me(msg, user_id)})
    return records


def fetch_channels(client: GraphClient, start: datetime, user_id: str | None, teams: int, channels: int, per_channel: int) -> list[dict[str, Any]]:
    records = []
    for team in client.list("/me/joinedTeams", {"$top": min(teams, 50)}, teams):
        try:
            team_channels = client.list(f"/teams/{team['id']}/channels", {"$top": min(channels, 50)}, channels)
        except GraphError:
            continue
        for channel in team_channels:
            try:
                messages = client.list(
                    f"/teams/{team['id']}/channels/{channel['id']}/messages",
                    {"$top": min(per_channel, 50)},
                    per_channel,
                )
            except GraphError:
                continue
            title = f"{team.get('displayName', 'Team')} / {channel.get('displayName', 'Channel')}"
            for msg in messages:
                when = dt(msg.get("createdDateTime"))
                body = html(deep(msg, "body", "content"))
                if (when and when < start) or not body or not mentions_me(msg, user_id):
                    continue
                records.append(rec("Teams Channel", msg, title, when, deep(msg, "from", "user") or {}, body) | {"mentions_me": True})
    return records


def sender_is_external(email: str, domains: list[str]) -> bool:
    domain = email.split("@")[-1].lower() if "@" in email else ""
    return bool(domain and not any(domain == internal or domain.endswith(f".{internal}") for internal in domains))


def due_date(text: str, base: datetime | None) -> str:
    base = base or datetime.now(timezone.utc)
    if re.search(r"\b(eod|cob)\b", text, re.I):
        return base.date().isoformat()
    try:
        found = search_dates(text, settings={"RELATIVE_BASE": base.replace(tzinfo=None), "PREFER_DATES_FROM": "future", "DATE_ORDER": "MDY"})
    except Exception:
        found = None
    if not found:
        return ""
    floor, ceiling = base.date() - timedelta(days=7), base.date() + timedelta(days=370)
    for _, parsed in found:
        if floor <= parsed.date() <= ceiling:
            return parsed.date().isoformat()
    return ""


def banner_for(text: str) -> str:
    lower = text.lower()
    for banner, words in BANNER_WORDS.items():
        if any(word in lower for word in words):
            return banner
    return "Unassigned"


def owner_for(text: str, record: dict[str, Any], current_user: str) -> str:
    match = OWNER_RE.search(text)
    if match:
        return match.group(1).strip()
    if re.search(r"\b(i will|i'll|my team will|we will)\b", text, re.I) and record.get("from_name"):
        return record["from_name"]
    return current_user or "Me"


def task_from(record: dict[str, Any], sentence: str, current_user: str, internal_domains: list[str]) -> dict[str, Any]:
    context = clean(f"{record.get('source_title', '')}. {sentence}. {record.get('text', '')}")
    due = due_date(sentence, record.get("source_date"))
    importance = str(record.get("importance", "")).lower()
    urgency = "High" if importance == "high" or URGENT_RE.search(context) else "Low"
    if due:
        parsed_due = date.fromisoformat(due)
        urgency = "High" if parsed_due <= date.today() + timedelta(days=3) else "Medium" if parsed_due <= date.today() + timedelta(days=14) else urgency
    elif urgency == "Low" and re.search(r"\b(this week|next week|soon|follow[- ]?up|outstanding|pending)\b", context, re.I):
        urgency = "Medium"

    desc = clean(sentence)
    desc = f"{desc[:257]}..." if len(desc) > 260 else desc
    raw_id = "|".join([record.get("source_type", ""), record.get("source_id", ""), str(record.get("source_date", "")), desc])
    external = sender_is_external(record.get("from_email", ""), internal_domains)
    status = "complete" if DONE_RE.search(sentence) else "in-progress" if PROGRESS_RE.search(sentence) else "to do"
    return {
        "task_id": hashlib.sha1(raw_id.encode("utf-8")).hexdigest()[:16],
        "task_description": desc,
        "urgency": urgency,
        "owner": owner_for(sentence, record, current_user),
        "status": status,
        "due_date": due,
        "banner": banner_for(context),
        "external_prompt": "Yes" if external else "No",
        "source": record.get("source_type", ""),
        "source_title": record.get("source_title", ""),
        "source_date": record.get("source_date"),
        "requester": record.get("from_name") or record.get("from_email", ""),
        "link": record.get("web_link", ""),
    }


def extract_tasks(records: list[dict[str, Any]], current_user: str, internal_domains: list[str]) -> list[dict[str, Any]]:
    tasks, seen = [], set()
    for record in records:
        text = clean(f"{record.get('source_title', '')}. {record.get('text', '')}")
        sentences = [clean(s) for s in re.split(r"(?<=[.!?])\s+|[\r\n]+|\s+-\s+", text) if clean(s)]
        candidates = [sentence for sentence in sentences if ACTION_RE.search(sentence)] or (sentences[:1] if record.get("mentions_me") else [])
        for sentence in candidates[:3]:
            if len(sentence) < 8:
                continue
            task = task_from(record, sentence, current_user, internal_domains)
            if task["task_id"] not in seen:
                seen.add(task["task_id"])
                tasks.append(task)
    return tasks


def load_overrides() -> dict[str, dict[str, Any]]:
    if not OVERRIDES.exists():
        return {}
    try:
        data = pd.read_csv(OVERRIDES).fillna("")
    except Exception:
        return {}
    return {str(row["task_id"]): row.to_dict() for _, row in data.iterrows()}


def apply_overrides(tasks: list[dict[str, Any]]) -> list[dict[str, Any]]:
    overrides = load_overrides()
    for task in tasks:
        for column in ["owner", "status", "due_date", "banner", "urgency"]:
            value = overrides.get(task["task_id"], {}).get(column, "")
            if pd.notna(value) and str(value).strip():
                task[column] = value
    return tasks


def save_overrides(df: pd.DataFrame) -> None:
    DATA_DIR.mkdir(exist_ok=True)
    out = df[["task_id", "owner", "status", "due_date", "banner", "urgency"]].copy()
    out["due_date"] = out["due_date"].apply(lambda v: v.isoformat() if isinstance(v, (date, datetime)) else ("" if pd.isna(v) else v))
    out.to_csv(OVERRIDES, index=False)


def sidebar() -> dict[str, Any]:
    st.sidebar.header("Refresh")
    options = {
        "days": st.sidebar.slider("Lookback days", 30, 180, 90, step=15),
        "inbox": st.sidebar.toggle("Outlook inbox", True),
        "calendar": st.sidebar.toggle("Outlook calendar", True),
        "chats": st.sidebar.toggle("Teams chats", True),
        "channels": st.sidebar.toggle("Teams channel mentions", True),
    }
    with st.sidebar.expander("Source limits"):
        options |= {
            "max_inbox": int(st.number_input("Inbox messages", 25, 500, 200, 25)),
            "max_calendar": int(st.number_input("Calendar events", 25, 500, 200, 25)),
            "max_chats": int(st.number_input("Chats", 5, 100, 30, 5)),
            "max_chat_messages": int(st.number_input("Messages per chat", 10, 200, 75, 5)),
            "max_teams": int(st.number_input("Teams", 1, 50, 10, 1)),
            "max_channels": int(st.number_input("Channels per team", 1, 50, 10, 1)),
            "max_channel_messages": int(st.number_input("Messages per channel", 10, 200, 50, 5)),
        }
    return options


def sort_tasks(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["due_date"] = pd.to_datetime(df["due_date"], errors="coerce").dt.date
    df["source_date"] = pd.to_datetime(df["source_date"], errors="coerce")
    df["_status"] = df["status"].map({"to do": 0, "in-progress": 1, "complete": 2}).fillna(3)
    df["_urgency"] = df["urgency"].map({"High": 0, "Medium": 1, "Low": 2}).fillna(3)
    df["_due"] = pd.to_datetime(df["due_date"], errors="coerce").fillna(pd.Timestamp.max)
    return df.sort_values(["_status", "_urgency", "_due", "source_date"], ascending=[True, True, True, False]).drop(columns=["_status", "_urgency", "_due"])


def empty_df() -> pd.DataFrame:
    return pd.DataFrame(columns=["task_id", "task_description", "urgency", "owner", "status", "due_date", "banner", "external_prompt", "source", "source_title", "source_date", "requester", "link"])


def main() -> None:
    st.set_page_config(page_title="LCL To-Do Dashboard", layout="wide")
    st.title("LCL To-Do Dashboard")

    options = sidebar()
    internal_domains = csv(secret("INTERNAL_EMAIL_DOMAINS", "rbi.com"), lower=True)
    configured_owners = csv(secret("OWNER_OPTIONS", "Me"))
    scopes = BASE_SCOPES + (CHANNEL_SCOPES if options["channels"] else [])
    token = microsoft_token(secret("AZURE_CLIENT_ID"), secret("AZURE_TENANT_ID", "organizations") or "organizations", scopes)
    client = GraphClient(token) if token else None

    if st.sidebar.button("Refresh tasks", type="primary", use_container_width=True, disabled=client is None) and client:
        start = datetime.now(timezone.utc) - timedelta(days=options["days"])
        end = datetime.now(timezone.utc)
        warnings, records = [], []
        with st.status("Pulling Microsoft sources...", expanded=True) as status:
            profile = client.get("/me", {"$select": "id,displayName,userPrincipalName,mail"})
            current_user = profile.get("displayName") or profile.get("mail") or "Me"
            for label, enabled, func in [
                ("Outlook inbox", options["inbox"], lambda: fetch_inbox(client, start, options["max_inbox"])),
                ("Outlook calendar", options["calendar"], lambda: fetch_calendar(client, start, end, options["max_calendar"])),
                ("Teams chats", options["chats"], lambda: fetch_chats(client, start, profile.get("id"), options["max_chats"], options["max_chat_messages"])),
                ("Teams channels", options["channels"], lambda: fetch_channels(client, start, profile.get("id"), options["max_teams"], options["max_channels"], options["max_channel_messages"])),
            ]:
                if not enabled:
                    continue
                try:
                    records.extend(func())
                except GraphError as exc:
                    warnings.append(f"{label}: {exc}")
            st.session_state["tasks"] = apply_overrides(extract_tasks(records, current_user, internal_domains))
            st.session_state["warnings"] = warnings
            st.session_state["last_refresh"] = datetime.now().strftime("%Y-%m-%d %H:%M")
            status.update(label=f"Refresh complete: {len(records)} source items", state="complete")

    uploaded = st.sidebar.file_uploader("Import task CSV", type=["csv"])
    if uploaded:
        st.session_state["tasks"] = pd.read_csv(uploaded).fillna("").to_dict("records")
        st.session_state["last_refresh"] = "CSV import"
    show_completed = st.sidebar.toggle("Show completed", value=False)

    df = pd.DataFrame(st.session_state.get("tasks", [])) if st.session_state.get("tasks") else empty_df()
    if df.empty:
        st.info("Connect Microsoft and refresh, or import a task CSV.")
        return

    df = sort_tasks(df)
    if not show_completed:
        df = df[df["status"] != "complete"]
    if st.session_state.get("last_refresh"):
        st.caption(f"Last refresh: {st.session_state['last_refresh']}")
    for warning in st.session_state.get("warnings", []):
        st.warning(warning)

    due = pd.to_datetime(df["due_date"], errors="coerce")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Open", int((df["status"] != "complete").sum()))
    col2.metric("High", int((df["urgency"] == "High").sum()))
    col3.metric("Due 7 days", int(((due >= pd.Timestamp(date.today())) & (due <= pd.Timestamp(date.today() + timedelta(days=7)))).sum()))
    col4.metric("External", int((df["external_prompt"] == "Yes").sum()))

    owners = sorted({"Me", *configured_owners, *[clean(o) for o in df["owner"].dropna().astype(str) if clean(o)], *[clean(r) for r in df["requester"].dropna().astype(str) if clean(r)]})
    cols = ["task_description", "urgency", "owner", "status", "due_date", "banner", "external_prompt", "source", "source_title", "requester", "link", "task_id"]
    edited = st.data_editor(
        df[cols],
        hide_index=True,
        use_container_width=True,
        num_rows="fixed",
        disabled=["task_description", "external_prompt", "source", "source_title", "requester", "link", "task_id"],
        column_config={
            "task_description": st.column_config.TextColumn("Task description", width="large"),
            "urgency": st.column_config.SelectboxColumn("Urgency", options=URGENCIES, required=True),
            "owner": st.column_config.SelectboxColumn("Owner", options=owners, required=True),
            "status": st.column_config.SelectboxColumn("Status", options=STATUSES, required=True),
            "due_date": st.column_config.DateColumn("Due date", format="YYYY-MM-DD"),
            "banner": st.column_config.SelectboxColumn("LCL banner", options=BANNERS, required=True),
            "link": st.column_config.LinkColumn("Source link"),
        },
    )

    left, right = st.columns([1, 4])
    if left.button("Save edits", type="primary", use_container_width=True):
        save_overrides(edited)
        st.success("Saved owner, status, due date, urgency, and banner overrides.")
    right.download_button("Download CSV", edited.to_csv(index=False).encode("utf-8"), "lcl_todo_dashboard.csv", "text/csv")


if __name__ == "__main__":
    main()
