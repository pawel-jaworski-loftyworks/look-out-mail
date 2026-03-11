# Outlook Mail Tray

A macOS menu bar app that shows your recent Outlook emails, and flashes red when mail is unread.

## Features

- Mail icon in menu bar with unread count badge
- Shows 10 most recent emails
- 🔴 Red dot indicator for unread emails
- Auto-refreshes every 60 seconds
- Click "Get new token..." to refresh authentication

## Installation

1. Create and activate virtual environment:

```bash
python3 -m venv venv
source venv/bin/activate
```

2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Install Playwright browser (required for login):

```bash
playwright install chromium
```

## Usage

### Step 1: Get Authentication Token

Run the login script to authenticate with Outlook:

```bash
./venv/bin/python outlook_login.py
```

This opens a browser window where you can log into your Outlook account.

**IMPORTANT: You MUST close the browser window after logging in for the token to be saved and the tray app to work.**

The script will:
- Open a Chromium browser
- Let you log into Outlook
- Automatically capture the authentication token
- Save your session (you won't need to log in again next time)

### Step 2: Run the Tray App

```bash
./venv/bin/python outlook_tray.py
```

The mail icon will appear in your menu bar.

### Running in Background

To run the tray app in the background:

```bash
./venv/bin/python outlook_tray.py &
```

## Files

| File                    | Description                                     |
|-------------------------|-------------------------------------------------|
| `outlook_login.py`      | Opens browser for Outlook login, captures token |
| `outlook_tray.py`       | Menu bar tray application                       |
| `fetch_outlook_mail.py` | CLI script to fetch emails (standalone)         |
| `token.json`            | Saved authentication token (auto-generated)     |
| `browser_data/`         | Persistent browser profile (auto-generated)     |
| `gotmail.svg`           | Custom tray icon                                |

## Token Expiration

When your token expires (usually after several hours), the tray app will show an error. To refresh:

1. Click the tray icon
2. Select "Get new token..."
3. **Close the browser window after it loads your inbox**
4. The token will be automatically refreshed

## Troubleshooting

### "No token - run outlook_login.py"

Run the login script first to authenticate.

### "Token expired"

Your authentication token has expired. Click "Get new token..." in the tray menu or run `outlook_login.py` again.

### Icon not appearing

Make sure you closed the browser window after logging in. The tray app waits for the browser to close before starting.
