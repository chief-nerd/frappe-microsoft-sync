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

    def get_access_token(self, force_refresh=False):
        cache_key = f"ms_access_token_{self.user}"
        if not force_refresh:
            cached_token = frappe.cache().get_value(cache_key)
            if cached_token:
                return cached_token

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

        if new_access_token:
            frappe.cache().set_value(cache_key, new_access_token, expires_in_sec=3000)

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
            self.access_token = self.get_access_token(force_refresh=True)
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

    def sync_contact(self, contact_doc):
        contact_data = {
            "givenName": contact_doc.first_name or "",
            "surname": contact_doc.last_name or "",
            "jobTitle": contact_doc.designation or "",
            "companyName": contact_doc.company_name or "",
            "emailAddresses": [
                {"address": contact_doc.email_id, "name": contact_doc.full_name}
            ] if contact_doc.email_id else [],
            "businessPhones": [contact_doc.phone] if contact_doc.phone else [],
            "mobilePhone": contact_doc.mobile_no or "",
        }

        if contact_doc.microsoft_contact_id:
            # Update
            res = self.request(
                "PATCH",
                f"/me/contacts/{contact_doc.microsoft_contact_id}",
                data=contact_data,
            )
            if res and res.status_code == 404:
                # Contact deleted in MS, recreate it
                res = self.request("POST", "/me/contacts", data=contact_data)
        else:
            # Create
            res = self.request("POST", "/me/contacts", data=contact_data)

        if res and res.status_code in [200, 201]:
            ms_contact_id = res.json().get("id")
            if ms_contact_id != contact_doc.microsoft_contact_id:
                frappe.db.set_value(
                    "Contact",
                    contact_doc.name,
                    "microsoft_contact_id",
                    ms_contact_id,
                    update_modified=False,
                )

    def delete_contact(self, microsoft_contact_id):
        if not microsoft_contact_id:
            return
        self.request("DELETE", f"/me/contacts/{microsoft_contact_id}")

    def sync_event(self, event_doc):
        # We need a timezone for Microsoft. 
        # In a real app, you might want this to be a per-user setting.
        timezone = frappe.utils.get_time_zone() or "UTC"

        # Utility to convert to Microsoft's datetime format
        def format_dt(dt):
            if not dt: return None
            # We assume the dt is in the server's local time
            from frappe.utils import get_datetime
            dt_obj = get_datetime(dt)
            return {
                "dateTime": dt_obj.strftime("%Y-%m-%dT%H:%M:%S"),
                "timeZone": timezone
            }

        event_data = {
            "subject": event_doc.subject or "No Subject",
            "body": {
                "contentType": "html",
                "content": event_doc.description or ""
            },
            "start": format_dt(event_doc.starts_on),
            "end": format_dt(event_doc.ends_on),
            "location": {"displayName": event_doc.location} if event_doc.location else None,
            "isAllDay": bool(event_doc.all_day)
        }

        # Handle participants
        if event_doc.event_participants:
            attendees = []
            for p in event_doc.event_participants:
                email = frappe.db.get_value(p.reference_doctype, p.reference_docname, "email_id")
                if email:
                    attendees.append({
                        "emailAddress": {
                            "address": email,
                            "name": p.reference_docname
                        },
                        "type": "required"
                    })
            event_data["attendees"] = attendees

        if event_doc.microsoft_event_id:
            res = self.request("PATCH", f"/me/events/{event_doc.microsoft_event_id}", data=event_data)
            if res and res.status_code == 404:
                res = self.request("POST", "/me/events", data=event_data)
        else:
            res = self.request("POST", "/me/events", data=event_data)

        if res and res.status_code in [200, 201]:
            ms_event_id = res.json().get("id")
            if ms_event_id != event_doc.microsoft_event_id:
                frappe.db.set_value("Event", event_doc.name, "microsoft_event_id", ms_event_id, update_modified=False)

    def delete_event(self, microsoft_event_id):
        if not microsoft_event_id:
            return
        self.request("DELETE", f"/me/events/{microsoft_event_id}")

    def pull_calendar_events(self):
        # Sync from the last hour (or from setting)
        from frappe.utils import add_hours, get_datetime_str, now_datetime
        
        # Pull events that have been modified recently
        # We could also use delta links for a more robust sync, but this is simpler for a prototype
        res = self.request("GET", "/me/events", params={
            "$filter": f"lastModifiedDateTime ge {(now_datetime() - add_hours(now_datetime(), -2)).strftime('%Y-%m-%dT%H:%M:%SZ')}"
        })

        if res and res.status_code == 200:
            events = res.json().get("value", [])
            for ms_event in events:
                self.process_incoming_event(ms_event)

    def process_incoming_event(self, ms_event):
        ms_id = ms_event.get("id")
        
        # Avoid circular sync
        if frappe.db.exists("Event", {"microsoft_event_id": ms_id}):
            event_doc = frappe.get_doc("Event", {"microsoft_event_id": ms_id})
            # Check if updated recently in MS (compare timestamps if desired)
            # For simplicity, we skip if we already have it linked unless you want full two-way merging
            return

        # Create new event in Frappe
        from frappe.utils import get_datetime
        
        # Microsoft returns datetimes in UTC or the user's timezone
        # We need to convert them to server local time for Frappe naive datetimes
        def parse_ms_dt(ms_dt_obj):
            if not ms_dt_obj: return None
            # ms_dt_obj = {"dateTime": "2023-10-27T12:00:00.0000000", "timeZone": "UTC"}
            import pytz
            from frappe.utils import get_time_zone
            
            dt_str = ms_dt_obj.get("dateTime").split(".")[0] # Strip microseconds
            tz_name = ms_dt_obj.get("timeZone")
            
            source_tz = pytz.timezone(tz_name)
            local_tz = pytz.timezone(get_time_zone() or "UTC")
            
            dt = source_tz.localize(get_datetime(dt_str))
            return dt.astimezone(local_tz).replace(tzinfo=None)

        new_event = frappe.get_doc({
            "doctype": "Event",
            "subject": ms_event.get("subject"),
            "description": ms_event.get("body", {}).get("content"),
            "starts_on": parse_ms_dt(ms_event.get("start")),
            "ends_on": parse_ms_dt(ms_event.get("end")),
            "all_day": ms_event.get("isAllDay"),
            "location": ms_event.get("location", {}).get("displayName"),
            "microsoft_event_id": ms_id,
            "owner": self.user,
            "event_type": "Private"
        })
        new_event.insert(ignore_permissions=True)

