from . import __version__ as app_version
from mimirio_sync import patch_oauth

try:
    patch_oauth()
except Exception:
    pass

app_name = "mimirio_sync"
app_title = "Mimirio Microsoft Sync"
app_publisher = "Mimirio"
app_description = (
    "Syncs Frappe ToDo records with Microsoft To Do and exposes iCalendar feeds."
)
app_email = "hello@mimirio.com"
app_license = "mit"

# ToDo hooks for Microsoft To Do synchronization
doc_events = {
    "ToDo": {
        "on_update": "mimirio_sync.sync.sync_todo_to_microsoft",
        "on_trash": "mimirio_sync.sync.delete_todo_from_microsoft",
    }
}

fixtures = [
    {"dt": "Custom Field", "filters": [["fieldname", "=", "microsoft_todo_id"]]}
]

# Whitelisted API methods
# The instruction said /api/method/mimirio_sync.api.get_calendar_feed
# This is automatically whitelisted if we use @frappe.whitelist in mimirio_sync/api.py
