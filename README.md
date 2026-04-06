# Copilot Cowork MCP

An MCP server that connects [GitHub Copilot CLI](https://docs.github.com/copilot/concepts/agents/about-copilot-cli) to **Microsoft 365 Copilot** (Cowork).

Ask questions about your emails, calendar, Teams messages, meetings, documents, contacts, and org data — all from the terminal. Send Teams messages, emails, and even images as attachments.

## Requirements

- Python 3.10+
- [GitHub Copilot CLI](https://docs.github.com/copilot/concepts/agents/about-copilot-cli)
- [Microsoft 365 Copilot license](https://learn.microsoft.com/en-us/microsoft-365/copilot/cowork/get-started) with [Copilot Cowork](https://m365.cloud.microsoft/chat) enabled

## Quick Start

1. **Clone and install dependencies**

   ```bash
   git clone https://github.com/msartem/copilot_cowork_mcp.git
   cd copilot_cowork_mcp
   pip install -r requirements.txt
   ```

2. **Launch Copilot CLI from the repo directory**

   ```bash
   copilot
   ```

3. **Sign in** (first time only)

   Copilot will automatically call `cowork_sign_in` when needed. An Edge
   browser window opens — sign in with your Microsoft 365 account. The
   auth code is captured automatically via Playwright (no manual copy-paste).

4. **Try it out**

   ```
   Summarize my unread emails from today
   What meetings do I have tomorrow?
   Send a Teams message to myself saying hello from CLI
   Send an email to John saying the report is ready
   Send the image /path/to/chart.png to my team on Teams
   Email the screenshot /path/to/screenshot.png to Sarah
   ```

After first sign-in, auth is silent via cached refresh tokens (~90 day
lifetime, auto-renewing). No browser window on subsequent uses.

## How It Works

1. You ask Copilot CLI a question about your M365 data
2. Copilot routes it to the `cowork_send_message` tool
3. The MCP server forwards it to Microsoft 365 Copilot
4. The response streams back to your terminal

For actions (send message, send email, etc.), `cowork_send_message` returns the
action details for review, then `cowork_action_approve` executes it. This gives
you a chance to confirm before anything is sent.

```
You → Copilot CLI → cowork MCP → M365 Copilot → your M365 data
```

## Tools

| Tool | Description |
|------|-------------|
| `cowork_sign_in` | Sign in to Microsoft 365 (opens Edge, one-time) |
| `cowork_send_message` | Send a message to M365 Copilot (multi-turn) |
| `cowork_send_image` | Send an image with optional text (Teams, email attachments) |
| `cowork_action_approve` | Approve and execute a pending action (send email, Teams message, etc.) |
| `cowork_new_session` | Start a fresh conversation |
| `cowork_session_info` | Show session state (user, runtime, turn count) |

## Authentication

Authentication uses [Playwright](https://playwright.dev/python/) to
automate browser sign-in via Microsoft Edge. This works identically on
macOS, Linux, and Windows — no platform-specific code.

### First sign-in

1. Copilot calls `cowork_sign_in` (or you trigger it manually)
2. An Edge window opens to the Microsoft sign-in page
3. Sign in with your M365 account (SSO may auto-complete this)
4. The auth code is captured automatically — the browser closes
5. Tokens are cached locally

### Subsequent runs

Cached refresh tokens are used — no browser window, no interaction needed.
Tokens auto-refresh in the background every 45 minutes.

### Browser selection

Playwright uses Microsoft Edge by default. To use a different Chromium-based browser:

```bash
export COWORK_BROWSER=chrome    # or: msedge (default), chromium
```

> **Note**: `chromium` requires `playwright install chromium`. Edge and Chrome
> use your installed browser directly — no extra install needed.

### Manual auth

```bash
python auth.py          # Interactive sign-in
python auth.py logout   # Clear cached tokens
```

### Token storage

| Platform | Cache location |
|----------|---------------|
| macOS / Linux | `~/.copilot-cowork-mcp/token_cache.json` |
| Windows | `%LOCALAPPDATA%\copilot-cowork-mcp\token_cache.json` |

### Environment Variables

| Variable | Description |
|----------|-------------|
| `COWORK_TOKEN` | Manual JWT override (skips browser auth) |
| `COWORK_TENANT` | Azure AD tenant ID (default: `common`) |
| `COWORK_BROWSER` | Browser for sign-in: `msedge`, `chrome`, `chromium` |
| `COWORK_SHOW_THINKING` | Show Cowork's reasoning in action responses (default: `false`) |

## Technical Details

See [TECHNICAL.md](TECHNICAL.md) for details on authentication, communication protocol, SSE events, and action approval flow.

## Disclaimer

This is an independent community project — **not affiliated with or supported by Microsoft**.

It works by communicating with the same undocumented web APIs that power the
[M365 Copilot web app](https://m365.cloud.microsoft/chat). These APIs may
change or break without notice. Use at your own risk.
