import frappe
from mimirio_sync.microsoft_graph import MicrosoftGraphClient

def sync_todo_to_microsoft(doc, method=None):
    if not doc.owner:
        return

    if not frappe.db.exists("MS Sync Settings", {"user": doc.owner}):
        return

    frappe.enqueue(
        "mimirio_sync.sync.enqueue_sync_todo",
        user=doc.owner,
        todo_name=doc.name,
        now=frappe.flags.in_test,
    )

def delete_todo_from_microsoft(doc, method=None):
    if not doc.owner or not doc.microsoft_todo_id:
        return
        
    # Enqueue the delete task
    frappe.enqueue(
        "mimirio_sync.sync.enqueue_delete_todo",
        user=doc.owner,
        microsoft_todo_id=doc.microsoft_todo_id,
        now=frappe.flags.in_test
    )

def enqueue_sync_todo(user, todo_name):
    client = MicrosoftGraphClient(user)
    todo_doc = frappe.get_doc("ToDo", todo_name)
    client.sync_todo(todo_doc)

def enqueue_delete_todo(user, microsoft_todo_id):
    client = MicrosoftGraphClient(user)
    client.delete_todo(microsoft_todo_id)
