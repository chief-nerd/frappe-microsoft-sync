from . import __version__ as app_version

app_name = "frappe_microsoft_sync"
app_title = "Frappe Microsoft Sync"
app_publisher = "Mimirio"
app_description = (
    "Syncs Frappe ToDo records with Microsoft To Do and exposes iCalendar feeds."
)
app_email = "hello@mimirio.com"
app_license = "mit"

before_request = [
    "frappe_microsoft_sync.patch_oauth",
]

# Contact hooks for Microsoft Contact synchronization
doc_events = {
    "ToDo": {
        "on_update": "frappe_microsoft_sync.sync.sync_todo_to_microsoft",
        "on_trash": "frappe_microsoft_sync.sync.delete_todo_from_microsoft",
    },
    "Contact": {
        "on_update": "frappe_microsoft_sync.sync.sync_contact_to_microsoft",
        "on_trash": "frappe_microsoft_sync.sync.delete_contact_from_microsoft",
    },
    "Event": {
        "on_update": "frappe_microsoft_sync.sync.sync_event_to_microsoft",
        "on_trash": "frappe_microsoft_sync.sync.delete_event_from_microsoft",
    }
}

fixtures = [
    {"dt": "Custom Field", "filters": [["fieldname", "in", ["microsoft_todo_id", "microsoft_contact_id", "microsoft_event_id"]]]}
]

scheduler_events = {
    "all": [
        "frappe_microsoft_sync.sync.pull_from_microsoft"
    ]
}

# Whitelisted API methods are defined with @frappe.whitelist in frappe_microsoft_sync/api.py
