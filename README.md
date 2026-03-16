# Mimirio Microsoft Sync

Frappe v15 app that provides bidirectional synchronization between ERPNext/Frappe and Microsoft 365:

- **ToDo ↔ Microsoft To Do** — syncs to a dedicated "ERP Tasks" list
- **Contacts ↔ Outlook Contacts**
- **Events ↔ Outlook Calendar**
- **iCalendar Feed** — per-user `.ics` endpoint for subscribing in any calendar client

Refresh tokens are captured automatically during Microsoft SSO login via Entra ID (formerly Azure AD).

## Prerequisites

- Frappe v15 with ERPNext
- Microsoft SSO configured as a **Social Login Key** ("Office 365" or "Microsoft")
- The Social Login Key must request these scopes: `offline_access Tasks.ReadWrite Calendars.Read Contacts.ReadWrite`
- A running Redis instance (used by Frappe; also used here for transient token caching)
- Background workers enabled (sync jobs are enqueued, not run inline)

## Installation

### Standard Bench

```bash
# From your bench directory:
bench get-app https://github.com/chief-nerd/frappe-microsoft-sync.git
bench --site YOUR_SITE install-app mimirio_sync
bench --site YOUR_SITE migrate
```

### Docker (frappe_docker)

Add the app to your `apps.json`:

```json
{
  "url": "https://github.com/chief-nerd/frappe-microsoft-sync.git",
  "branch": "main"
}
```

Rebuild the image and update:

```bash
docker compose build
docker compose up -d
docker compose exec backend bench --site YOUR_SITE install-app mimirio_sync
docker compose exec backend bench --site YOUR_SITE migrate
```

## Updating

### Standard Bench

```bash
bench update --apps mimirio_sync
bench --site YOUR_SITE migrate
```

### Docker

Update the branch/tag in `apps.json`, then:

```bash
docker compose build --no-cache
docker compose up -d
docker compose exec backend bench --site YOUR_SITE migrate
```

## Configuration

1. **Enable Microsoft SSO** — Ensure a Social Login Key for "Office 365" (or "Microsoft") exists with the required scopes. The included patch (`v0_1/setup_microsoft_sso`) adds missing scopes automatically on migrate.

2. **User login** — Each user must sign in via Microsoft SSO at least once. This captures the OAuth refresh token needed for background sync.

3. **Per-user settings** — Open the **MS Sync Settings** doctype for a user (auto-created on first sync) to toggle:
   - `Enabled` — master switch
   - `Sync To-Dos` / `Sync Contacts` / `Sync Events` — per-feature toggles
   - `Pull Sync Interval (Minutes)` — how often the scheduler pulls changes from Microsoft (default 60, set 0 for push-only)

4. **iCalendar feed** — Generate a token for a user, then subscribe from any calendar app:
   ```
   https://YOUR_SITE/api/method/mimirio_sync.api.get_calendar_feed?token=TOKEN
   ```

## How It Works

| Direction           | Trigger                             | Mechanism                                                              |
| ------------------- | ----------------------------------- | ---------------------------------------------------------------------- |
| Frappe → Microsoft  | Doc event (`on_update`, `on_trash`) | Enqueued background job via `frappe.enqueue`                           |
| Microsoft → Frappe  | Scheduler (`all` interval)          | `pull_from_microsoft` checks each enabled user                         |
| OAuth token capture | SSO login                           | Monkey-patches `get_info_via_oauth` and `login_oauth_user` at app load |

## License

MIT
