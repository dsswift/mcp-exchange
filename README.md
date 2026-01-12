# MCP Exchange Online Server

MCP (Model Context Protocol) server for Microsoft Exchange Online integration via Microsoft Graph API. Provides tools for querying emails, managing messages, and checking calendar availability.

## Features

### Email Tools
- **list_mail_folders** - List all mail folders (Inbox, Archive, etc.)
- **search_emails** - Search emails with filters (sender, subject, date range, read/unread, attachments)
- **get_email** - Get full email details including body
- **archive_email** - Move email to Archive folder
- **delete_email** - Move email to Deleted Items
- **create_draft** - Create a draft email

### Calendar Tools
- **list_calendars** - List all calendars
- **list_events** - List calendar events in a date range
- **get_event** - Get full event details
- **get_free_busy** - Check free/busy schedule for users (for scheduling meetings)

## Prerequisites

- Python 3.10+
- Microsoft 365 account with Exchange Online
- Azure AD app registration with appropriate permissions

## Azure AD App Registration Setup

1. Go to [Azure Portal](https://portal.azure.com) → Microsoft Entra ID → App registrations → **New registration**

2. Configure the app:
   - **Name**: `MCP Exchange` (or your preference)
   - **Supported account types**: "Accounts in this organizational directory only"
   - **Redirect URI**: Leave blank
   - Click **Register**

3. Copy these values for configuration:
   - **Application (client) ID** → `EXCHANGE_CLIENT_ID`
   - **Directory (tenant) ID** → `EXCHANGE_TENANT_ID`

4. Add API permissions (API permissions → Add a permission → Microsoft Graph → Delegated):
   - `User.Read` - Sign in and read user profile
   - `Mail.ReadWrite` - Read, move, and delete emails
   - `Mail.Send` - Create draft emails
   - `Calendars.Read` - Read calendar events
   - `Calendars.Read.Shared` - Read free/busy for other users

5. Enable public client flow (Authentication → Advanced settings):
   - Set **"Allow public client flows"** to **Yes**
   - Click **Save**

6. Grant admin consent (if required by your organization):
   - Click **"Grant admin consent for [organization]"**
   - Or request consent from your IT administrator

## Installation

```bash
# Clone the repository
git clone https://github.com/dsswift/mcp-exchange.git
cd mcp-exchange

# Install with uv (recommended)
uv sync

# Or with pip (creates venv manually)
python -m venv .venv
source .venv/bin/activate
pip install -e .
```

## Configuration

Create a `.env` file (copy from `.env.sample`):

```bash
cp .env.sample .env
```

Required environment variables:

| Variable | Required | Default | Description |
|----------|----------|---------|-------------|
| `EXCHANGE_CLIENT_ID` | Yes | - | Azure AD app client ID |
| `EXCHANGE_TENANT_ID` | No | `common` | Azure AD tenant ID |
| `EXCHANGE_TOKEN_CACHE` | No | `~/.mcp-exchange/token_cache.json` | Token cache location |
| `EXCHANGE_TIMEOUT` | No | `30` | HTTP timeout in seconds |

## Usage

### Running Standalone

```bash
# Run the server
uv run mcp-exchange
```

On first run, you'll be prompted to authenticate:

```
============================================================
AUTHENTICATION REQUIRED
============================================================

To sign in, visit: https://microsoft.com/devicelogin
Enter this code: XXXXXXXX

Waiting for authentication...
```

### Registering with Claude Code

```bash
# Using uvx with Git URL (recommended)
claude mcp add exchange --scope user \
  -e EXCHANGE_CLIENT_ID=your-client-id \
  -e EXCHANGE_TENANT_ID=your-tenant-id \
  -- uvx --from git+https://github.com/dsswift/mcp-exchange.git mcp-exchange

# Or with local installation
claude mcp add exchange --scope user \
  -e EXCHANGE_CLIENT_ID=your-client-id \
  -e EXCHANGE_TENANT_ID=your-tenant-id \
  -- uv run --directory /path/to/mcp-exchange mcp-exchange
```

### MCP Client Configuration

For other MCP clients, add to your configuration:

```json
{
  "mcpServers": {
    "exchange": {
      "command": "uvx",
      "args": ["--from", "git+https://github.com/dsswift/mcp-exchange.git", "mcp-exchange"],
      "env": {
        "EXCHANGE_CLIENT_ID": "your-client-id",
        "EXCHANGE_TENANT_ID": "your-tenant-id"
      }
    }
  }
}
```

## Tool Examples

### Search Emails

```python
# Recent unread emails
search_emails(is_read=False)

# Emails from a specific sender
search_emails(sender="boss@company.com")

# Emails with attachments from this month
search_emails(from_date="2024-01-01", has_attachments=True)
```

### Check Free/Busy

```python
# Check when your boss is available today
get_free_busy(
    emails="chris@company.com",
    start_time="2024-01-15T09:00:00",
    end_time="2024-01-15T17:00:00",
    timezone="America/New_York"
)

# Check multiple people for a meeting
get_free_busy(
    emails="chris@company.com,alex@company.com,jordan@company.com",
    start_time="2024-01-15T08:00:00",
    end_time="2024-01-15T18:00:00"
)
```

### Create a Draft Email

```python
create_draft(
    subject="Meeting Follow-up",
    body="Hi team,\n\nThanks for the productive meeting today...",
    to_recipients="team@company.com,manager@company.com"
)
```

## Development

```bash
# Install dev dependencies
uv pip install -e ".[dev]"

# Run tests
pytest

# Run linter
ruff check src tests

# Run type checker
mypy src
```

## Authentication Notes

- **Device code flow**: The server uses device code flow, which works well for CLI tools and doesn't require a redirect URI
- **Token caching**: Tokens are cached locally, so you only need to authenticate once (until the refresh token expires)
- **Scopes**: The server requests only the permissions needed for the tools it provides
- **No app-only auth**: This server uses delegated permissions (user authentication), not app-only authentication

## Troubleshooting

### "Authentication failed" errors

1. Verify your `EXCHANGE_CLIENT_ID` is correct
2. Ensure "Allow public client flows" is enabled in Azure AD
3. Check that admin consent has been granted for the required permissions

### "Permission denied" errors

1. Verify the API permissions are configured correctly
2. Request admin consent from your IT administrator
3. Check that your account has access to the mailbox/calendars you're querying

### Token cache issues

```bash
# Clear the token cache to force re-authentication
rm ~/.mcp-exchange/token_cache.json
```

## License

MIT
