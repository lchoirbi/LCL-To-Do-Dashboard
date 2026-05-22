# LCL To-Do Dashboard

Streamlit dashboard for organizing outstanding Loblaw Business tasks from Outlook inbox, Outlook calendar, and Teams activity.

[![Open in Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.io/deploy?repository=https://github.com/lchoirbi/LCL-To-Do-Dashboard&mainModule=app.py)

## What It Shows

- Task description extracted from recent Outlook and Teams text.
- Urgency, sorted with high-priority and near-due work first.
- Owner, status, due date, and LCL banner as editable fields.
- External prompt flag based on sender domain.
- Source links back to Outlook or Teams when Microsoft Graph provides them.

## Microsoft Graph Setup

Create an Azure App Registration configured as a public client, then add delegated Microsoft Graph permissions:

- `User.Read`
- `Mail.Read`
- `Calendars.Read`
- `Chat.Read`

For Teams channel mention scanning, also add:

- `Team.ReadBasic.All`
- `Channel.ReadBasic.All`
- `ChannelMessage.Read.All`

Some Teams channel permissions may require tenant admin consent. The dashboard will still load any source surfaces that consent allows and will show warnings for blocked surfaces.

Set these Streamlit secrets:

```toml
AZURE_CLIENT_ID = "your-app-registration-client-id"
AZURE_TENANT_ID = "organizations"
INTERNAL_EMAIL_DOMAINS = "rbi.com"
OWNER_OPTIONS = "Me, Name One, Name Two"
```

## Local Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

The app uses Microsoft device login. Manual owner, status, due date, urgency, and banner edits are stored locally in `data/task_overrides.csv` and merged back into future refreshes by task ID.
