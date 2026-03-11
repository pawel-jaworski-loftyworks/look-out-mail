#!/usr/bin/env python3
"""
Outlook mail tray icon for macOS.
Shows recent emails in the menu bar.
"""

import sys
import json
from pathlib import Path
from datetime import datetime

from PySide6.QtWidgets import QApplication, QSystemTrayIcon, QMenu
from PySide6.QtGui import QIcon, QPixmap, QPainter, QColor, QFont, QAction
from PySide6.QtCore import QTimer, Qt

import requests

# Configuration
TOKEN_FILE = Path(__file__).parent / "token.json"
BASE_URL = "https://outlook.cloud.microsoft/owa/service.svc"
REFRESH_INTERVAL_MS = 60_000  # 1 minute


def load_token() -> tuple[str, str, str]:
    """Load token from token.json file."""
    if not TOKEN_FILE.exists():
        return "", "", ""
    data = json.loads(TOKEN_FILE.read_text())
    return (
        data.get("bearer_token", ""),
        data.get("anchor_mailbox", ""),
        data.get("session_id", ""),
    )


def get_headers(token: str, anchor_mailbox: str, session_id: str, action: str) -> dict:
    """Build headers for OWA API request."""
    return {
        "accept": "*/*",
        "action": action,
        "authorization": f"Bearer {token}",
        "content-type": "application/json; charset=utf-8",
        "x-anchormailbox": anchor_mailbox,
        "x-owa-sessionid": session_id,
        "x-req-source": "Mail",
        "prefer": 'IdType="ImmutableId"',
    }


def fetch_conversations(token: str, anchor_mailbox: str, session_id: str, count: int = 10) -> list[dict]:
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

    headers = get_headers(token, anchor_mailbox, session_id, "FindConversation")
    response = requests.post(url, headers=headers, json=payload, timeout=30)
    response.raise_for_status()

    data = response.json()
    return data.get("Body", {}).get("Conversations", []) or []


def create_mail_icon(unread_count: int = 0) -> QIcon:
    """Create a mail icon from SVG with optional unread badge."""
    from PySide6.QtSvg import QSvgRenderer

    size = 22  # macOS menu bar icon size
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.transparent)

    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing)

    # Load and render SVG
    svg_path = Path(__file__).parent / "gotmail.svg"
    if svg_path.exists():
        renderer = QSvgRenderer(str(svg_path))
        renderer.render(painter)
    else:
        # Fallback: simple envelope if SVG not found
        painter.setPen(QColor(80, 80, 80))
        painter.setBrush(QColor(255, 255, 255))
        painter.drawRect(2, 5, 18, 12)
        painter.drawLine(2, 5, 11, 12)
        painter.drawLine(20, 5, 11, 12)

    # Draw unread badge if needed
    if unread_count > 0:
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor(255, 59, 48))  # iOS red
        painter.drawEllipse(12, 0, 10, 10)

        painter.setPen(QColor(255, 255, 255))
        painter.setFont(QFont("Arial", 7, QFont.Bold))
        badge_text = str(unread_count) if unread_count < 10 else "9+"
        painter.drawText(12, 0, 10, 10, Qt.AlignCenter, badge_text)

    painter.end()
    return QIcon(pixmap)


class OutlookTray(QSystemTrayIcon):
    def __init__(self):
        super().__init__()

        self.conversations = []
        self.unread_count = 0
        self.last_error = None

        # Create menu
        self.menu = QMenu()
        self.setContextMenu(self.menu)

        # Set initial icon
        self.setIcon(create_mail_icon(0))
        self.setToolTip("Outlook Mail")

        # Setup refresh timer
        self.timer = QTimer()
        self.timer.timeout.connect(self.refresh_mail)
        self.timer.start(REFRESH_INTERVAL_MS)

        # Initial load
        self.refresh_mail()

        self.show()

    def refresh_mail(self):
        """Fetch mail and update menu."""
        token, anchor_mailbox, session_id = load_token()

        if not token:
            self.last_error = "No token - run outlook_login.py"
            self.update_menu()
            return

        try:
            self.conversations = fetch_conversations(token, anchor_mailbox, session_id, count=10)
            self.unread_count = sum(c.get("UnreadCount", 0) for c in self.conversations)
            self.last_error = None
        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 401:
                self.last_error = "Token expired - run outlook_login.py"
            else:
                self.last_error = f"HTTP {e.response.status_code}"
            self.conversations = []
            self.unread_count = 0
        except Exception as e:
            self.last_error = str(e)[:50]
            self.conversations = []
            self.unread_count = 0

        self.setIcon(create_mail_icon(self.unread_count))
        self.update_menu()

    def update_menu(self):
        """Rebuild the menu with current conversations."""
        self.menu.clear()

        # Header
        if self.unread_count > 0:
            header = self.menu.addAction(f"Inbox ({self.unread_count} unread)")
        else:
            header = self.menu.addAction("Inbox")
        header.setEnabled(False)
        self.menu.addSeparator()

        # Error message if any
        if self.last_error:
            error_action = self.menu.addAction(f"Error: {self.last_error}")
            error_action.setEnabled(False)
            self.menu.addSeparator()

        # Conversations
        if self.conversations:
            for conv in self.conversations[:10]:
                topic = conv.get("ConversationTopic", "(No subject)")
                unread = conv.get("UnreadCount", 0)
                senders = conv.get("UniqueSenders", []) or conv.get("GlobalUniqueSenders", [])
                sender = senders[0] if senders else ""

                # Truncate long subjects
                if len(topic) > 45:
                    topic = topic[:42] + "..."

                # Format menu item: red dot for unread, gray text for read
                if unread > 0:
                    label = f"🔴 {topic}"
                else:
                    label = f"      {topic}"

                action = self.menu.addAction(label)
                action.setToolTip(f"From: {sender}")

                # Gray out read messages
                if unread == 0:
                    action.setEnabled(False)
        else:
            no_mail = self.menu.addAction("No messages")
            no_mail.setEnabled(False)

        self.menu.addSeparator()

        # Refresh action
        refresh_action = self.menu.addAction("Refresh")
        refresh_action.triggered.connect(self.refresh_mail)

        # Login action
        login_action = self.menu.addAction("Get new token...")
        login_action.triggered.connect(self.open_login)

        self.menu.addSeparator()

        # Quit action
        quit_action = self.menu.addAction("Quit")
        quit_action.triggered.connect(QApplication.quit)

    def open_login(self):
        """Open the login script."""
        import subprocess
        script_path = Path(__file__).parent / "outlook_login.py"
        python_path = Path(__file__).parent / "venv" / "bin" / "python"
        subprocess.Popen([str(python_path), str(script_path)])


def main():
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)  # Keep running when no windows

    # Check if system tray is available
    if not QSystemTrayIcon.isSystemTrayAvailable():
        print("System tray not available")
        sys.exit(1)

    tray = OutlookTray()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
