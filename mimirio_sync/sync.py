import frappe
from mimirio_sync.microsoft_graph import MicrosoftGraphClient


def sync_todo_to_microsoft(doc, method=None):
    if frappe.flags.in_microsoft_sync:
        return

    if not doc.owner:
        return

    settings = frappe.db.get_value("MS Sync Settings", {"user": doc.owner}, ["enabled", "sync_todos"], as_dict=True)
    if not settings or not settings.enabled or not settings.sync_todos:
        return

    frappe.enqueue(
        "mimirio_sync.sync.enqueue_sync_todo",
        user=doc.owner,
        todo_name=doc.name,
        now=frappe.flags.in_test,
    )


def delete_todo_from_microsoft(doc, method=None):
    if frappe.flags.in_microsoft_sync:
        return

    if not doc.owner or not doc.microsoft_todo_id:
        return

    settings = frappe.db.get_value("MS Sync Settings", {"user": doc.owner}, ["enabled", "sync_todos"], as_dict=True)
    if not settings or not settings.enabled or not settings.sync_todos:
        return

    # Enqueue the delete task
    frappe.enqueue(
        "mimirio_sync.sync.enqueue_delete_todo",
        user=doc.owner,
        microsoft_todo_id=doc.microsoft_todo_id,
        now=frappe.flags.in_test,
    )


def enqueue_sync_todo(user, todo_name):
    client = MicrosoftGraphClient(user)
    try:
        todo_doc = frappe.get_doc("ToDo", todo_name)
    except frappe.DoesNotExistError:
        return
    client.sync_todo(todo_doc)


def enqueue_delete_todo(user, microsoft_todo_id):
    client = MicrosoftGraphClient(user)
    client.delete_todo(microsoft_todo_id)


def sync_contact_to_microsoft(doc, method=None):
    if frappe.flags.in_microsoft_sync:
        return

    if not doc.owner:
        return

    settings = frappe.db.get_value("MS Sync Settings", {"user": doc.owner}, ["enabled", "sync_contacts"], as_dict=True)
    if not settings or not settings.enabled or not settings.sync_contacts:
        return

    frappe.enqueue(
        "mimirio_sync.sync.enqueue_sync_contact",
        user=doc.owner,
        contact_name=doc.name,
        now=frappe.flags.in_test,
    )


def delete_contact_from_microsoft(doc, method=None):
    if frappe.flags.in_microsoft_sync:
        return

    if not doc.owner or not doc.microsoft_contact_id:
        return

    settings = frappe.db.get_value("MS Sync Settings", {"user": doc.owner}, ["enabled", "sync_contacts"], as_dict=True)
    if not settings or not settings.enabled or not settings.sync_contacts:
        return

    frappe.enqueue(
        "mimirio_sync.sync.enqueue_delete_contact",
        user=doc.owner,
        microsoft_contact_id=doc.microsoft_contact_id,
        now=frappe.flags.in_test,
    )


def enqueue_sync_contact(user, contact_name):
    client = MicrosoftGraphClient(user)
    try:
        contact_doc = frappe.get_doc("Contact", contact_name)
    except frappe.DoesNotExistError:
        return
    client.sync_contact(contact_doc)


def enqueue_delete_contact(user, microsoft_contact_id):
    client = MicrosoftGraphClient(user)
    client.delete_contact(microsoft_contact_id)


def sync_event_to_microsoft(doc, method=None):
    if frappe.flags.in_microsoft_sync:
        return

    if not doc.owner:
        return

    settings = frappe.db.get_value("MS Sync Settings", {"user": doc.owner}, ["enabled", "sync_events"], as_dict=True)
    if not settings or not settings.enabled or not settings.sync_events:
        return

    frappe.enqueue(
        "mimirio_sync.sync.enqueue_sync_event",
        user=doc.owner,
        event_name=doc.name,
        now=frappe.flags.in_test,
    )


def delete_event_from_microsoft(doc, method=None):
    if frappe.flags.in_microsoft_sync:
        return

    if not doc.owner or not doc.microsoft_event_id:
        return

    settings = frappe.db.get_value("MS Sync Settings", {"user": doc.owner}, ["enabled", "sync_events"], as_dict=True)
    if not settings or not settings.enabled or not settings.sync_events:
        return

    frappe.enqueue(
        "mimirio_sync.sync.enqueue_delete_event",
        user=doc.owner,
        microsoft_event_id=doc.microsoft_event_id,
        now=frappe.flags.in_test,
    )


def enqueue_sync_event(user, event_name):
    client = MicrosoftGraphClient(user)
    try:
        event_doc = frappe.get_doc("Event", event_name)
    except frappe.DoesNotExistError:
        return
    client.sync_event(event_doc)


def enqueue_delete_event(user, microsoft_event_id):
    client = MicrosoftGraphClient(user)
    client.delete_event(microsoft_event_id)


def pull_from_microsoft():
    """Scheduled task to pull changes from Microsoft."""
    from frappe.utils import now_datetime, add_minutes

    settings_list = frappe.get_all("MS Sync Settings", filters={"enabled": 1}, fields=["user", "pull_sync_interval", "last_pull_datetime"])
    for s in settings_list:
        if not s.pull_sync_interval:
            continue

        # Check if it's time to pull
        if not s.last_pull_datetime or add_minutes(s.last_pull_datetime, s.pull_sync_interval) <= now_datetime():
            frappe.enqueue(
                "mimirio_sync.sync.pull_for_user",
                user=s.user,
                now=frappe.flags.in_test
            )


def pull_for_user(user):
    frappe.flags.in_microsoft_sync = True
    try:
        client = MicrosoftGraphClient(user)
        # Verify sync type enabled
        if client.settings.enabled and client.settings.sync_events:
            client.pull_calendar_events()
        
        # Update last pull time
        from frappe.utils import now_datetime
        client.settings.db_set("last_pull_datetime", now_datetime())
    finally:
        frappe.flags.in_microsoft_sync = False
