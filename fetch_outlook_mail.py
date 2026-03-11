#!/usr/bin/env python3
"""
Fetch newest emails from Outlook using OWA API.
Uses Bearer token from browser session (via token.json or fallback).
"""

import json
import sys
from pathlib import Path
import requests

TOKEN_FILE = Path(__file__).parent / "token.json"
BASE_URL = "https://outlook.cloud.microsoft/owa/service.svc"


def load_token() -> tuple[str, str, str]:
    """Load token from token.json file."""
    if not TOKEN_FILE.exists():
        print(f"ERROR: {TOKEN_FILE} not found.")
        print("Run outlook_login.py first to get a token.")
        sys.exit(1)

    data = json.loads(TOKEN_FILE.read_text())
    bearer_token = data.get("bearer_token", "")
    anchor_mailbox = data.get("anchor_mailbox", "")
    session_id = data.get("session_id", "")

    if not bearer_token:
        print("ERROR: No bearer_token in token.json")
        sys.exit(1)

    return bearer_token, anchor_mailbox, session_id


# Load credentials
BEARER_TOKEN, ANCHOR_MAILBOX, SESSION_ID = load_token()


def get_headers(action: str) -> dict:
    """Build headers for OWA API request."""
    return {
        "accept": "*/*",
        "action": action,
        "authorization": f"Bearer {BEARER_TOKEN}",
        "content-type": "application/json; charset=utf-8",
        "x-anchormailbox": ANCHOR_MAILBOX,
        "x-owa-sessionid": SESSION_ID,
        "x-req-source": "Mail",
        "prefer": 'IdType="ImmutableId"',
    }


def find_conversations(count: int = 10) -> dict:
    """Fetch latest conversations from inbox."""
    url = f"{BASE_URL}?action=FindConversation&app=Mail"

    payload = {
        "__type": "FindConversationJsonRequest:#Exchange",
        "Header": {
            "__type": "JsonRequestHeaders:#Exchange",
            "RequestServerVersion": "Exchange2016",
            "TimeZoneContext": {
                "__type": "TimeZoneContext:#Exchange",
                "TimeZoneDefinition": {
                    "__type": "TimeZoneDefinitionType:#Exchange",
                    "Id": "Central European Standard Time"
                }
            }
        },
        "Body": {
            "__type": "FindConversationRequest:#Exchange",
            "ParentFolderId": {
                "__type": "TargetFolderId:#Exchange",
                "BaseFolderId": {
                    "__type": "DistinguishedFolderId:#Exchange",
                    "Id": "inbox"
                }
            },
            "ConversationShape": {
                "__type": "ConversationResponseShape:#Exchange",
                "BaseShape": "Default"
            },
            "Paging": {
                "__type": "IndexedPageView:#Exchange",
                "BasePoint": "Beginning",
                "Offset": 0,
                "MaxEntriesReturned": count
            },
            "SortOrder": [{
                "__type": "SortResults:#Exchange",
                "Order": "Descending",
                "Path": {
                    "__type": "PropertyUri:#Exchange",
                    "FieldURI": "ConversationLastDeliveryTime"
                }
            }]
        }
    }

    response = requests.post(url, headers=get_headers("FindConversation"), json=payload)
    response.raise_for_status()
    return response.json()


def print_conversations(data: dict) -> None:
    """Pretty print conversation list."""
    body = data.get("Body", {})
    conversations = body.get("Conversations", [])

    if not conversations:
        print("No conversations found.")
        if body.get("MessageText"):
            print(f"Error: {body.get('MessageText')}")
        return

    print(f"\n{'='*70}")
    print(f" Found {len(conversations)} conversations")
    print(f"{'='*70}\n")

    for i, conv in enumerate(conversations, 1):
        topic = conv.get("ConversationTopic", "(No subject)")
        preview = conv.get("Preview", "")[:100] if conv.get("Preview") else ""
        last_delivery = conv.get("LastDeliveryTime", "")
        unread_count = conv.get("UnreadCount", 0)
        message_count = conv.get("MessageCount", 0)

        # Get sender from UniqueSenders or GlobalUniqueSenders
        senders = conv.get("UniqueSenders", []) or conv.get("GlobalUniqueSenders", [])
        sender = senders[0] if senders else "Unknown"

        # Format unread indicator
        unread_marker = " *NEW*" if unread_count > 0 else ""

        print(f"{i}. {topic}{unread_marker}")
        print(f"   From: {sender}")
        print(f"   Date: {last_delivery}")
        print(f"   Messages: {message_count} (Unread: {unread_count})")
        if preview:
            print(f"   Preview: {preview}...")
        print()


def main():
    print("Fetching latest emails from Outlook...")

    try:
        data = find_conversations(count=10)
        print_conversations(data)
    except requests.exceptions.HTTPError as e:
        if e.response.status_code == 401:
            print("ERROR: Token expired. Run: ./venv/bin/python outlook_login.py")
        else:
            print(f"HTTP Error: {e}")
            print(f"Response: {e.response.text}")
    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    main()
