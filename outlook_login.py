#!/usr/bin/env python3
"""
Open browser for Outlook login and capture Bearer token.
Uses persistent browser profile - login once, reuse session.
Saves token to token.json for use by fetch_outlook_mail.py
"""

import json
from pathlib import Path
from playwright.sync_api import sync_playwright

TOKEN_FILE = Path(__file__).parent / "token.json"
BROWSER_DATA_DIR = Path(__file__).parent / "browser_data"
OUTLOOK_URL = "https://outlook.cloud.microsoft/mail/"


def save_token(token: str, anchor_mailbox: str = "", session_id: str = ""):
    """Save token and metadata to file."""
    data = {
        "bearer_token": token,
        "anchor_mailbox": anchor_mailbox,
        "session_id": session_id,
    }
    TOKEN_FILE.write_text(json.dumps(data, indent=2))
    print(f"\nToken saved to {TOKEN_FILE}")


def extract_token_from_headers(headers: dict) -> str | None:
    """Extract Bearer token from request headers."""
    auth = headers.get("authorization", "")
    if auth.startswith("Bearer "):
        return auth[7:]  # Remove "Bearer " prefix
    return None


def main():
    print("Opening browser for Outlook...")
    print(f"Browser profile saved at: {BROWSER_DATA_DIR}")
    print("(After first login, you won't need to log in again)\n")

    captured_token = None
    anchor_mailbox = ""
    session_id = ""
    token_captured_time = None

    with sync_playwright() as p:
        # Use persistent context - saves cookies/session between runs
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(BROWSER_DATA_DIR),
            headless=False,
            channel="chromium",
        )
        page = context.pages[0] if context.pages else context.new_page()

        def handle_request(request):
            nonlocal captured_token, anchor_mailbox, session_id, token_captured_time

            # Look for OWA API requests that have Bearer token
            if "outlook" in request.url and "service.svc" in request.url:
                headers = request.headers
                token = extract_token_from_headers(headers)

                if token and token != captured_token:
                    captured_token = token
                    anchor_mailbox = headers.get("x-anchormailbox", "")
                    session_id = headers.get("x-owa-sessionid", "")
                    token_captured_time = True

                    # Show token preview
                    print(f"Captured token: {token[:3]}...{token[-3:]}")
                    print(f"Anchor mailbox: {anchor_mailbox}")
                    print("\nToken captured! Closing browser...")

        # Listen for requests
        page.on("request", handle_request)

        # Navigate to Outlook
        page.goto(OUTLOOK_URL)

        print("Waiting for token capture...")
        print("(Browser will auto-close once token is captured)\n")

        # Wait for token capture or browser close
        import time
        while context.pages:
            if token_captured_time:
                # Give a moment for any final requests, then close
                time.sleep(1)
                break
            try:
                context.pages[0].wait_for_event("close", timeout=500)
            except:
                pass

        context.close()

    if captured_token:
        save_token(captured_token, anchor_mailbox, session_id)
        print("\nSuccess! Token captured")
    else:
        print("\nNo token captured. Make sure to navigate to your inbox.")


if __name__ == "__main__":
    main()
