import requests
import frappe
from frappe import _
from frappe.utils.password import get_decrypted_password, set_encrypted_password


class MicrosoftGraphClient:
    def __init__(self, user):
        self.user = user
        self.settings = self.get_sync_settings()
        self.access_token = self.get_access_token()

    def get_sync_settings(self):
        settings = frappe.get_all(
            "MS Sync Settings", filters={"user": self.user}, fields=["*"]
        )
        if not settings:
            doc = frappe.get_doc({"doctype": "MS Sync Settings", "user": self.user})
            doc.insert(ignore_permissions=True)
            return doc
        return frappe.get_doc("MS Sync Settings", settings[0].name)

    def get_access_token(self):
        # Retrieve stored refresh token
        refresh_token = get_decrypted_password(
            "User", self.user, "microsoft_refresh_token", raise_exception=False
        )
        if not refresh_token:
            # Try to get from MS Sync Settings if it's there
            refresh_token = get_decrypted_password(
                "MS Sync Settings",
                self.settings.name,
                "refresh_token",
                raise_exception=False,
            )

        if not refresh_token:
            frappe.log_error(
                title="MS Sync Error",
                message=f"No Microsoft refresh token found for user {self.user}",
            )
            return None

        # Refresh the token
        provider_name = (
            "Office 365"
            if frappe.db.exists("Social Login Key", "Office 365")
            else "Microsoft"
        )
        social_login_key = frappe.get_doc("Social Login Key", provider_name)

        data = {
            "client_id": social_login_key.client_id,
            "client_secret": get_decrypted_password(
                "Social Login Key", provider_name, "client_secret"
            ),
            "refresh_token": refresh_token,
            "grant_type": "refresh_token",
            "scope": "offline_access Tasks.ReadWrite Calendars.Read",
        }

        response = requests.post(
            "https://login.microsoftonline.com/common/oauth2/v2.0/token", data=data
        )
        if response.status_code != 200:
            frappe.log_error(
                title="MS Sync Error",
                message=f"Failed to refresh Microsoft token for {self.user}: {response.text}",
            )
            return None

        token_data = response.json()
        new_access_token = token_data.get("access_token")
        new_refresh_token = token_data.get("refresh_token")

        if new_refresh_token:
            set_encrypted_password(
                "MS Sync Settings",
                self.settings.name,
                new_refresh_token,
                "refresh_token",
            )

        return new_access_token

    def request(self, method, endpoint, data=None, params=None):
        if not self.access_token:
            return None

        url = f"https://graph.microsoft.com/v1.0{endpoint}"
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json",
        }

        response = requests.request(
            method, url, headers=headers, json=data, params=params
        )

        if response.status_code == 401:
            # Token might have expired just now, try one refresh
            self.access_token = self.get_access_token()
            if self.access_token:
                headers["Authorization"] = f"Bearer {self.access_token}"
                response = requests.request(
                    method, url, headers=headers, json=data, params=params
                )

        return response

    def get_or_create_todo_list(self):
        if self.settings.microsoft_list_id:
            # Check if it still exists
            res = self.request(
                "GET", f"/me/todo/lists/{self.settings.microsoft_list_id}"
            )
            if res and res.status_code == 200:
                return self.settings.microsoft_list_id

        # Search for "ERP Tasks"
        res = self.request("GET", "/me/todo/lists")
        if res and res.status_code == 200:
            lists = res.json().get("value", [])
            for lst in lists:
                if lst.get("displayName") == "ERP Tasks":
                    self.settings.microsoft_list_id = lst.get("id")
                    self.settings.save(ignore_permissions=True)
                    return self.settings.microsoft_list_id

        # Create it
        res = self.request("POST", "/me/todo/lists", data={"displayName": "ERP Tasks"})
        if res and res.status_code in [200, 201]:
            list_id = res.json().get("id")
            self.settings.microsoft_list_id = list_id
            self.settings.save(ignore_permissions=True)
            return list_id

        return None

    def sync_todo(self, todo_doc):
        list_id = self.get_or_create_todo_list()
        if not list_id:
            return

        task_data = {
            "title": todo_doc.description or "No Description",
            "status": "completed" if todo_doc.status == "Closed" else "notStarted",
        }

        if todo_doc.reference_type and todo_doc.reference_name:
            task_data["body"] = {
                "content": f"Source: Mimirio ERP ({todo_doc.reference_type}: {todo_doc.reference_name})",
                "contentType": "text",
            }

        if todo_doc.date:
            task_data["dueDateTime"] = {
                "dateTime": f"{todo_doc.date}T12:00:00",
                "timeZone": "UTC",
            }

        if todo_doc.microsoft_todo_id:
            # Update
            res = self.request(
                "PATCH",
                f"/me/todo/lists/{list_id}/tasks/{todo_doc.microsoft_todo_id}",
                data=task_data,
            )
            if res and res.status_code == 404:
                # Task deleted in MS, recreate it
                res = self.request(
                    "POST", f"/me/todo/lists/{list_id}/tasks", data=task_data
                )
        else:
            # Create
            res = self.request(
                "POST", f"/me/todo/lists/{list_id}/tasks", data=task_data
            )

        if res and res.status_code in [200, 201]:
            ms_todo_id = res.json().get("id")
            if ms_todo_id != todo_doc.microsoft_todo_id:
                frappe.db.set_value(
                    "ToDo",
                    todo_doc.name,
                    "microsoft_todo_id",
                    ms_todo_id,
                    update_modified=False,
                )

    def delete_todo(self, microsoft_todo_id):
        list_id = self.get_or_create_todo_list()
        if not list_id or not microsoft_todo_id:
            return

        self.request("DELETE", f"/me/todo/lists/{list_id}/tasks/{microsoft_todo_id}")
