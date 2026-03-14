import frappe
from frappe import _
import secrets
from werkzeug.wrappers import Response


def _ics_escape(text):
    """Escape text per RFC 5545."""
    text = text.replace("\\", "\\\\")
    text = text.replace(";", "\\;")
    text = text.replace(",", "\\,")
    text = text.replace("\n", "\\n")
    text = text.replace("\r", "")
    return text


@frappe.whitelist(allow_guest=True)
def get_calendar_feed(token):
    if not token:
        frappe.throw(_("Token is required"), frappe.PermissionError)

    # Token lookup must bypass permissions — request runs as Guest
    user = frappe.db.get_value(
        "MS Sync Settings",
        {"calendar_sync_token": token},
        "user",
    )
    if not user:
        frappe.throw(_("Invalid token"), frappe.PermissionError)

    from frappe.utils import add_months, nowdate

    start_filter = add_months(nowdate(), -1)
    end_filter = add_months(nowdate(), 12)

    events = frappe.db.sql(
        """
        SELECT DISTINCT e.name, e.subject, e.description,
               e.starts_on, e.ends_on, e.all_day, e.location
        FROM `tabEvent` e
        LEFT JOIN `tabEvent Participants` p ON p.parent = e.name
        WHERE (e.owner = %s OR p.reference_docname = %s)
          AND e.status != 'Cancelled'
          AND e.starts_on >= %s
          AND e.starts_on <= %s
        """,
        (user, user, start_filter, end_filter),
        as_dict=True,
    )

    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//Mimirio//ERP//EN",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
    ]

    for event in events:
        lines.append("BEGIN:VEVENT")
        lines.append(f"UID:{event.name}@mimirio.com")
        lines.append(f"SUMMARY:{_ics_escape(event.subject or 'No Subject')}")
        if event.description:
            lines.append(f"DESCRIPTION:{_ics_escape(event.description)}")

        if event.starts_on:
            if event.all_day:
                lines.append(f"DTSTART;VALUE=DATE:{event.starts_on.strftime('%Y%m%d')}")
            else:
                lines.append(f"DTSTART:{event.starts_on.strftime('%Y%m%dT%H%M%SZ')}")

        if event.ends_on:
            if event.all_day:
                lines.append(f"DTEND;VALUE=DATE:{event.ends_on.strftime('%Y%m%d')}")
            else:
                lines.append(f"DTEND:{event.ends_on.strftime('%Y%m%dT%H%M%SZ')}")

        if event.location:
            lines.append(f"LOCATION:{_ics_escape(event.location)}")

        lines.append("END:VEVENT")

    lines.append("END:VCALENDAR")

    ics_body = "\r\n".join(lines)
    frappe.response.filename = "calendar.ics"
    frappe.response.filecontent = ics_body
    frappe.response.type = "download"
    frappe.response.content_type = "text/calendar; charset=utf-8"


def generate_calendar_token(user):
    """Generate and persist a calendar feed token for a user."""
    token = secrets.token_urlsafe(32)
    from mimirio_sync.microsoft_graph import MicrosoftGraphClient

    client = MicrosoftGraphClient(user)
    client.settings.calendar_sync_token = token
    client.settings.save(ignore_permissions=True)
    return token
