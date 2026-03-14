# Mimirio Microsoft Sync

Frappe v15 app to synchronize ToDo items with Microsoft To Do and expose iCalendar feeds for Events.

## Features
- **Microsoft To Do Sync**: Automatically syncs Frappe ToDo records to a dedicated "ERP Tasks" list in Microsoft To Do.
- **iCalendar Feed**: Secure per-user iCalendar feed for Frappe Event documents.
- **Entra ID Integration**: Leverages existing Microsoft SSO to capture and manage refresh tokens securely.

## Setup
1. Install the app on your Frappe bench.
2. Ensure Microsoft SSO is configured in "Social Login Key".
3. Users should log in via Microsoft SSO once to authorize and capture the refresh token.
4. Access your calendar feed at `/api/method/mimirio_sync.api.get_calendar_feed?token=YOUR_TOKEN`.
