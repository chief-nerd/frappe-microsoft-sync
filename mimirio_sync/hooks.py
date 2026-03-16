from . import __version__ as app_version

app_name = "mimirio_sync"
app_title = "Mimirio Microsoft Sync"
app_publisher = "Mimirio"
app_description = (
    "Syncs Frappe ToDo records with Microsoft To Do and exposes iCalendar feeds."
)
app_email = "hello@mimirio.com"
app_license = "mit"

before_request = [
    "mimirio_sync.patch_oauth",
]

# Contact hooks for Microsoft Contact synchronization
doc_events = {
    "ToDo": {
        "on_update": "mimirio_sync.sync.sync_todo_to_microsoft",
        "on_trash": "mimirio_sync.sync.delete_todo_from_microsoft",
    },
    "Contact": {
        "on_update": "mimirio_sync.sync.sync_contact_to_microsoft",
        "on_trash": "mimirio_sync.sync.delete_contact_from_microsoft",
    },
    "Event": {
        "on_update": "mimirio_sync.sync.sync_event_to_microsoft",
        "on_trash": "mimirio_sync.sync.delete_event_from_microsoft",
    }
}

fixtures = [
    {"dt": "Custom Field", "filters": [["fieldname", "in", ["microsoft_todo_id", "microsoft_contact_id", "microsoft_event_id"]]]}
]

scheduler_events = {
    "all": [
        "mimirio_sync.sync.pull_from_microsoft"
    ]
}

# Whitelisted API methods
# The instruction said /api/method/mimirio_sync.api.get_calendar_feed
# This is automatically whitelisted if we use @frappe.whitelist in mimirio_sync/api.py
