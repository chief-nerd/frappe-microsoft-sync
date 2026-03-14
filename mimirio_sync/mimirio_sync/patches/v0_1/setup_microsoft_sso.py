import frappe
import json

def execute():
    # Ensure "Office 365" Social Login Key has the correct scopes
    if frappe.db.exists("Social Login Key", "Office 365"):
        doc = frappe.get_doc("Social Login Key", "Office 365")
        
        auth_url_data = {}
        if doc.auth_url_data:
            auth_url_data = json.loads(doc.auth_url_data)
            
        scopes = auth_url_data.get("scope", "")
        required_scopes = ["offline_access", "Tasks.ReadWrite", "Calendars.Read"]
        
        for scope in required_scopes:
            if scope not in scopes:
                scopes += f" {scope}"
                
        auth_url_data["scope"] = scopes.strip()
        doc.auth_url_data = json.dumps(auth_url_data)
        doc.save(ignore_permissions=True)
