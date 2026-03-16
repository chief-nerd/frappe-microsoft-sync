import frappe
from frappe import _
import secrets
from werkzeug.wrappers import Response


def _to_utc_str(dt):
    from frappe.utils import get_time_zone
    import pytz
    local_tz = pytz.timezone(get_time_zone() or "UTC")
    if hasattr(dt, "tzinfo") and dt.tzinfo:
        utc_dt = dt.astimezone(pytz.utc)
    else:
        utc_dt = local_tz.localize(dt).astimezone(pytz.utc)
    return utc_dt.strftime('%Y%m%dT%H%M%SZ')


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
                lines.append(f"DTSTART:{_to_utc_str(event.starts_on)}")

        if event.ends_on:
            if event.all_day:
                lines.append(f"DTEND;VALUE=DATE:{event.ends_on.strftime('%Y%m%d')}")
            else:
                lines.append(f"DTEND:{_to_utc_str(event.ends_on)}")

        if event.location:
            lines.append(f"LOCATION:{_ics_escape(event.location)}")

        lines.append("END:VEVENT")

    lines.append("END:VCALENDAR")

    ics_body = "\r\n".join(lines)
    frappe.response.filename = "calendar.ics"
    frappe.response.filecontent = ics_body
    frappe.response.type = "download"
    frappe.response.content_type = "text/calendar; charset=utf-8"


@frappe.whitelist(allow_guest=True)
def get_contact_feed(token):
    if not token:
        frappe.throw(_("Token is required"), frappe.PermissionError)

    user = frappe.db.get_value(
        "MS Sync Settings",
        {"calendar_sync_token": token},
        "user",
    )
    if not user:
        frappe.throw(_("Invalid token"), frappe.PermissionError)

    # Fetch contacts owned by the user or linked to relevant documents
    # For a "public" list perspective, we'll fetch all contacts for now
    # but strictly this could be filtered.
    contacts = frappe.get_all(
        "Contact",
        fields=[
            "first_name",
            "last_name",
            "email_id",
            "phone",
            "mobile_no",
            "company_name",
            "designation",
        ],
    )

    vcf_cards = []
    for contact in contacts:
        card = ["BEGIN:VCARD", "VERSION:3.0"]
        full_name = f"{contact.first_name or ''} {contact.last_name or ''}".strip()
        card.append(f"FN:{full_name}")
        card.append(f"N:{contact.last_name or ''};{contact.first_name or ''};;;")

        if contact.email_id:
            card.append(f"EMAIL;TYPE=INTERNET:{contact.email_id}")
        if contact.phone:
            card.append(f"TEL;TYPE=WORK,VOICE:{contact.phone}")
        if contact.mobile_no:
            card.append(f"TEL;TYPE=CELL,VOICE:{contact.mobile_no}")
        if contact.company_name:
            card.append(f"ORG:{contact.company_name}")
        if contact.designation:
            card.append(f"TITLE:{contact.designation}")

        card.append("END:VCARD")
        vcf_cards.append("\n".join(card))

    vcf_body = "\n\n".join(vcf_cards)
    frappe.response.filename = "contacts.vcf"
    frappe.response.filecontent = vcf_body
    frappe.response.type = "download"
    frappe.response.content_type = "text/vcard; charset=utf-8"


def generate_calendar_token(user):
    """Generate and persist a calendar feed token for a user."""
    token = secrets.token_urlsafe(32)
    from mimirio_sync.microsoft_graph import MicrosoftGraphClient

    client = MicrosoftGraphClient(user)
    client.settings.calendar_sync_token = token
    client.settings.save(ignore_permissions=True)
    return token
