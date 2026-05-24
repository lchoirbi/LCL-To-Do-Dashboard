# Microsoft Outlook and Teams Connection Setup

The dashboard connects to Outlook and Teams through Microsoft Graph. Streamlit Cloud needs a Microsoft Entra app registration client ID before it can ask you to sign in.

## 1. Create an app registration

In Microsoft Entra admin center, create an app registration for this dashboard.

Suggested settings:

- Account type: organizational directory only, or multitenant organizational directories if your tenant requires it.
- Platform/authentication type: public client/native app.
- Client secret: not needed for this dashboard.

Copy the Application (client) ID.

## 2. Add delegated Microsoft Graph permissions

Add these delegated permissions:

- `User.Read`
- `Mail.Read`
- `Calendars.Read`
- `Chat.Read`

For Teams channel mentions, also add:

- `Team.ReadBasic.All`
- `Channel.ReadBasic.All`
- `ChannelMessage.Read.All`

Teams channel permissions often require tenant admin consent. If admin consent is not available, turn off `Teams channel mentions` in the dashboard sidebar and use Outlook plus Teams chats first.

## 3. Add Streamlit secrets

In Streamlit Cloud, open the app, choose `Manage app`, then add secrets:

```toml
AZURE_CLIENT_ID = "your-app-registration-client-id"
AZURE_TENANT_ID = "organizations"
INTERNAL_EMAIL_DOMAINS = "rbi.com"
OWNER_OPTIONS = "Me, Name One, Name Two"
```

Save the secrets and reboot the Streamlit app.

## 4. Sign in from the dashboard

After the app restarts, use `Sign in to Microsoft` in the dashboard sidebar. Microsoft will show a device code login. After you complete the sign-in, return to the dashboard and click `I completed sign-in`, then `Refresh tasks`.
