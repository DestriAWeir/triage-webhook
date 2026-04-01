"""
Triage Email → Azure DevOps Work Item Webhook

Receives Microsoft Graph mail notifications for triage@ennrgy.com,
creates/updates ADO work items, and sends confirmation emails.
"""

import os
import re
import sys
import json
import logging
import time
from datetime import datetime, timedelta, timezone

import requests
import msal
from flask import Flask, request, jsonify

# Force unbuffered output so Render shows logs immediately
os.environ["PYTHONUNBUFFERED"] = "1"

# ---------------------------------------------------------------------------
# Configuration (set via environment variables on Render)
# ---------------------------------------------------------------------------
TENANT_ID       = os.environ.get("TENANT_ID", "")
CLIENT_ID       = os.environ.get("CLIENT_ID", "")
CLIENT_SECRET   = os.environ.get("CLIENT_SECRET", "")
TRIAGE_MAILBOX  = os.environ.get("TRIAGE_MAILBOX", "triage@ennrgy.com")

ADO_ORG         = os.environ.get("ADO_ORG", "ennrgyai")
ADO_PROJECT     = os.environ.get("ADO_PROJECT", "Risk360")
ADO_PAT         = os.environ.get("ADO_PAT", "")           # Personal Access Token
ADO_WORK_ITEM_TYPE = os.environ.get("ADO_WORK_ITEM_TYPE", "Bug")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
ADO_BASE   = f"https://dev.azure.com/{ADO_ORG}/{ADO_PROJECT}/_apis"

HTTP_TIMEOUT = 30   # seconds for all outbound HTTP calls

# ---------------------------------------------------------------------------
# Logging — force flush after every message
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    stream=sys.stderr,
)
log = logging.getLogger(__name__)

# Ensure all handlers flush immediately
for handler in logging.root.handlers:
    if hasattr(handler, "stream"):
        handler.stream = sys.stderr

# ---------------------------------------------------------------------------
# Flask app
# ---------------------------------------------------------------------------
app = Flask(__name__)

# ---------------------------------------------------------------------------
# Microsoft Graph auth (client credentials / daemon)
# ---------------------------------------------------------------------------
_graph_token_cache = {"token": None, "expires_at": 0}


def get_graph_token():
    """Get an access token for Microsoft Graph using client credentials."""
    now = time.time()
    if _graph_token_cache["token"] and now < _graph_token_cache["expires_at"] - 60:
        return _graph_token_cache["token"]

    log.info("Acquiring new Graph token via MSAL...")
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    msal_app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=authority, client_credential=CLIENT_SECRET
    )
    result = msal_app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" not in result:
        log.error("Failed to acquire Graph token: %s",
                  result.get("error_description", result))
        raise RuntimeError("Could not acquire Graph token")

    log.info("Graph token acquired successfully")
    _graph_token_cache["token"] = result["access_token"]
    _graph_token_cache["expires_at"] = now + result.get("expires_in", 3600)
    return result["access_token"]


def graph_headers():
    return {
        "Authorization": f"Bearer {get_graph_token()}",
        "Content-Type": "application/json",
    }


# ---------------------------------------------------------------------------
# Deduplication cache — Graph often sends duplicate notifications
# ---------------------------------------------------------------------------
_processed_messages = {}          # message_id -> timestamp
DEDUP_TTL_SECONDS = 300           # ignore duplicates within 5 minutes


def _is_duplicate(message_id):
    """Return True if we've already processed this message_id recently."""
    now = time.time()
    # Clean up expired entries
    expired = [k for k, v in _processed_messages.items()
               if now - v > DEDUP_TTL_SECONDS]
    for k in expired:
        del _processed_messages[k]

    if message_id in _processed_messages:
        return True
    _processed_messages[message_id] = now
    return False


# ---------------------------------------------------------------------------
# Azure DevOps helpers
# ---------------------------------------------------------------------------
def ado_headers():
    import base64
    encoded = base64.b64encode(f":{ADO_PAT}".encode()).decode()
    return {
        "Authorization": f"Basic {encoded}",
        "Content-Type": "application/json-patch+json",
    }


def ado_query_by_conversation_id(conversation_id):
    """Search for existing work item by Custom.TriageConversationID."""
    wiql = {
        "query": (
            f"SELECT [System.Id] FROM WorkItems "
            f"WHERE [Custom.TriageConversationID] = '{conversation_id}' "
            f"AND [System.TeamProject] = '{ADO_PROJECT}' "
            f"ORDER BY [System.CreatedDate] DESC"
        )
    }
    url = f"{ADO_BASE}/wit/wiql?api-version=7.1"
    resp = requests.post(url, json=wiql, headers=ado_headers(), timeout=HTTP_TIMEOUT)
    if resp.status_code == 200:
        items = resp.json().get("workItems", [])
        if items:
            return items[0]["id"]
    return None


def ado_query_by_subject(cleaned_subject):
    """Fallback: search by cleaned subject within last 30 days."""
    cutoff = (datetime.now(timezone.utc) - timedelta(days=30)).strftime("%Y-%m-%d")
    wiql = {
        "query": (
            f"SELECT [System.Id] FROM WorkItems "
            f"WHERE [Custom.TriageSubject] = '{cleaned_subject}' "
            f"AND [System.CreatedDate] >= '{cutoff}' "
            f"AND [System.TeamProject] = '{ADO_PROJECT}' "
            f"ORDER BY [System.CreatedDate] DESC"
        )
    }
    url = f"{ADO_BASE}/wit/wiql?api-version=7.1"
    resp = requests.post(url, json=wiql, headers=ado_headers(), timeout=HTTP_TIMEOUT)
    if resp.status_code == 200:
        items = resp.json().get("workItems", [])
        if items:
            return items[0]["id"]
    return None


def download_email_eml(message_id):
    """Download the raw MIME content (.eml) of an email from Graph API."""
    url = f"{GRAPH_BASE}/users/{TRIAGE_MAILBOX}/messages/{message_id}/$value"
    hdrs = {"Authorization": f"Bearer {get_graph_token()}"}
    resp = requests.get(url, headers=hdrs, timeout=HTTP_TIMEOUT)
    if resp.status_code == 200:
        log.info("Downloaded .eml for message %s (%d bytes)",
                 message_id, len(resp.content))
        return resp.content
    else:
        log.error("Failed to download .eml for %s: %s %s",
                  message_id, resp.status_code, resp.text)
        return None


def ado_upload_attachment(file_name, file_bytes):
    """Upload a file to ADO attachment storage. Returns the attachment URL."""
    import base64
    encoded_pat = base64.b64encode(f":{ADO_PAT}".encode()).decode()
    url = (f"{ADO_BASE}/wit/attachments"
           f"?fileName={file_name}&api-version=7.1")
    hdrs = {
        "Authorization": f"Basic {encoded_pat}",
        "Content-Type": "application/octet-stream",
    }
    resp = requests.post(url, data=file_bytes, headers=hdrs, timeout=HTTP_TIMEOUT)
    if resp.status_code in (200, 201):
        attachment_url = resp.json().get("url", "")
        log.info("Uploaded attachment '%s' → %s", file_name, attachment_url)
        return attachment_url
    else:
        log.error("Failed to upload attachment '%s': %s %s",
                  file_name, resp.status_code, resp.text)
        return None


def ado_attach_file_to_work_item(work_item_id, attachment_url, comment=""):
    """Link an uploaded attachment to an ADO work item."""
    patches = [
        {
            "op": "add",
            "path": "/relations/-",
            "value": {
                "rel": "AttachedFile",
                "url": attachment_url,
                "attributes": {
                    "comment": comment
                }
            }
        }
    ]
    url = f"{ADO_BASE}/wit/workitems/{work_item_id}?api-version=7.1"
    resp = requests.patch(url, json=patches, headers=ado_headers(), timeout=HTTP_TIMEOUT)
    if resp.status_code == 200:
        log.info("Attached file to work item #%s", work_item_id)
        return True
    else:
        log.error("Failed to attach file to #%s: %s %s",
                  work_item_id, resp.status_code, resp.text)
        return False


def attach_email_to_work_item(work_item_id, message_id, sender_email,
                               subject, received_dt):
    """Download an email as .eml and attach it to an ADO work item."""
    eml_bytes = download_email_eml(message_id)
    if not eml_bytes:
        return False

    # Create a descriptive filename
    safe_subject = re.sub(r'[^\w\s-]', '', subject or 'email')[:50].strip()
    safe_subject = re.sub(r'\s+', '_', safe_subject)
    file_name = f"{safe_subject}.eml"

    attachment_url = ado_upload_attachment(file_name, eml_bytes)
    if not attachment_url:
        return False

    comment = (f"Email from {sender_email} — "
               f"{received_dt or 'unknown time'}")
    return ado_attach_file_to_work_item(work_item_id, attachment_url, comment)


# Image content types that should be embedded inline
IMAGE_CONTENT_TYPES = {
    "image/png", "image/jpeg", "image/jpg", "image/gif",
    "image/bmp", "image/webp", "image/tiff",
}


def fetch_email_attachments(message_id):
    """Fetch attachment metadata for an email from Graph API."""
    url = (f"{GRAPH_BASE}/users/{TRIAGE_MAILBOX}/messages/{message_id}"
           f"/attachments")
    resp = requests.get(url, headers=graph_headers(), timeout=HTTP_TIMEOUT)
    if resp.status_code == 200:
        attachments = resp.json().get("value", [])
        log.info("Found %d attachment(s) on message %s",
                 len(attachments), message_id)
        return attachments
    else:
        log.error("Failed to fetch attachments for %s: %s %s",
                  message_id, resp.status_code, resp.text)
        return []


def process_email_attachments_for_inline(message_id):
    """
    Fetch email attachments, upload images to ADO, and return:
      - inline_html: HTML <img> tags for images to embed in description
      - file_attachments: list of (ado_url, filename) for non-image files
      - cid_map: dict mapping content_id -> ado_url for replacing cid:
                 references in the email body HTML
    """
    import base64 as b64mod
    attachments = fetch_email_attachments(message_id)
    inline_html_parts = []
    file_attachments = []   # (ado_url, filename) for non-image files
    cid_map = {}            # content_id -> ado_url

    for att in attachments:
        att_name = att.get("name", "attachment")
        content_type = (att.get("contentType") or "").lower()
        content_bytes_b64 = att.get("contentBytes", "")
        is_inline = att.get("isInline", False)
        content_id = att.get("contentId", "")

        if not content_bytes_b64:
            log.info("Skipping attachment '%s' (no contentBytes — may be "
                     "a reference attachment)", att_name)
            continue

        try:
            file_bytes = b64mod.b64decode(content_bytes_b64)
        except Exception:
            log.error("Failed to decode base64 for attachment '%s'", att_name)
            continue

        log.info("Processing attachment '%s' (%s, %d bytes, inline=%s, "
                 "contentId=%s)", att_name, content_type, len(file_bytes),
                 is_inline, content_id)

        # Upload to ADO attachment storage
        ado_url = ado_upload_attachment(att_name, file_bytes)
        if not ado_url:
            continue

        if content_type in IMAGE_CONTENT_TYPES:
            # Track cid: -> ADO URL mapping for replacing in body HTML
            if content_id:
                cid_map[content_id] = ado_url
            # Build an <img> tag for inline embedding (used as fallback
            # for images not referenced via cid: in the body)
            inline_html_parts.append(
                f'<div style="margin:8px 0;">'
                f'<img src="{ado_url}" alt="{att_name}" '
                f'style="max-width:100%;height:auto;" />'
                f'<br/><em style="color:#666;font-size:12px;">'
                f'{att_name}</em></div>'
            )
        else:
            # Non-image: will be attached as a file to the work item
            file_attachments.append((ado_url, att_name))

    inline_html = "\n".join(inline_html_parts)
    return inline_html, file_attachments, cid_map


def build_clean_description(body_html_raw, cid_map, inline_images_html):
    """Build a clean, simple HTML description that ADO can reliably render.

    Instead of trying to sanitize complex Outlook email HTML (which ADO's
    UI renderer silently discards even when the API accepts it), we:
      1. Extract plain text from the email body
      2. Rebuild simple HTML using only basic tags ADO supports
      3. Insert inline images using the ADO attachment URLs from cid_map
    """
    if not body_html_raw:
        return inline_images_html or ""

    # --- Step 1: extract the <body> content if wrapped in full HTML doc ---
    m = re.search(r'<body[^>]*>(.*)</body>', body_html_raw,
                  flags=re.DOTALL | re.IGNORECASE)
    body_fragment = m.group(1) if m else body_html_raw

    # --- Step 2: replace cid: refs with ADO URLs BEFORE stripping HTML ---
    #     This lets us preserve <img> tags that reference uploaded images.
    if cid_map:
        for content_id, ado_url in cid_map.items():
            body_fragment = re.sub(
                rf'<img[^>]*src="cid:{re.escape(content_id)}[^"]*"[^>]*/?>',
                f'<img src="{ado_url}" style="max-width:100%;" />',
                body_fragment,
                flags=re.IGNORECASE,
            )

    # --- Step 3: extract text and images in document order ---
    #     Walk through the HTML, pulling out text runs and <img> tags.
    parts = []
    pos = 0
    # Find every <img> tag (already converted to ADO URLs above)
    for img_m in re.finditer(r'<img\s[^>]*src="(https://dev\.azure\.com/[^"]+)"[^>]*/?>',
                             body_fragment, flags=re.IGNORECASE):
        # Grab text between previous position and this image
        text_chunk = body_fragment[pos:img_m.start()]
        text_only = re.sub(r'<[^>]+>', ' ', text_chunk)
        text_only = re.sub(r'[ \t]+', ' ', text_only).strip()
        text_only = text_only.replace('&nbsp;', ' ')
        if text_only:
            # Convert line breaks to <br> and wrap in <p>
            paragraphs = [p.strip() for p in re.split(r'\n\s*\n', text_only) if p.strip()]
            for para in paragraphs:
                parts.append(f"<p>{para}</p>")
        # Add the image with simple, clean markup
        parts.append(
            f'<p><img src="{img_m.group(1)}" '
            f'style="max-width:100%;height:auto;" /></p>'
        )
        pos = img_m.end()

    # Grab any trailing text after the last image
    text_chunk = body_fragment[pos:]
    text_only = re.sub(r'<[^>]+>', ' ', text_chunk)
    text_only = re.sub(r'[ \t]+', ' ', text_only).strip()
    text_only = text_only.replace('&nbsp;', ' ')
    if text_only:
        paragraphs = [p.strip() for p in re.split(r'\n\s*\n', text_only) if p.strip()]
        for para in paragraphs:
            parts.append(f"<p>{para}</p>")

    # --- Step 4: if no ADO-URL images were found in the body, append
    #     the gallery of all uploaded images as a fallback ---
    if inline_images_html and not any('dev.azure.com' in p for p in parts):
        parts.append(inline_images_html)

    result = "\n".join(parts) if parts else "(No text content in email)"
    return result


def replace_cid_references(body_html, cid_map):
    """Replace cid: image references in email HTML with ADO attachment URLs.

    Email clients embed inline images as <img src="cid:content_id">.
    This swaps those with the actual ADO attachment URLs so images
    render correctly in ADO work item descriptions.
    """
    if not cid_map or not body_html:
        return body_html

    for content_id, ado_url in cid_map.items():
        # Match both cid:content_id and cid:content_id@... forms
        body_html = re.sub(
            rf'src="cid:{re.escape(content_id)}[^"]*"',
            f'src="{ado_url}"',
            body_html,
            flags=re.IGNORECASE,
        )
    return body_html


def ado_create_work_item(title, body_html, conversation_id, cleaned_subject,
                          source, sender_email):
    """Create a new ADO work item."""
    patches = [
        {"op": "add", "path": "/fields/System.Title",
         "value": title},
                {"op": "add", "path": "/fields/Microsoft.VSTS.TCM.ReproSteps",
                          "value": body_html},
        {"op": "add", "path": "/fields/System.Description",
         "value": body_html},
        {"op": "add", "path": "/fields/Custom.TriageConversationID",
         "value": conversation_id or ""},
        {"op": "add", "path": "/fields/Custom.TriageSubject",
         "value": cleaned_subject},
        {"op": "add", "path": "/fields/Custom.AddTriageSource",
         "value": source},
        {"op": "add", "path": "/fields/Custom.TriageSenderEmail",
         "value": sender_email},
    ]
    log.info("Creating ADO work item — Description length: %d chars",
             len(body_html))
    # Log first 300 chars of description for debugging
    log.info("Description preview: %.300s", body_html[:300] if body_html else "(empty)")

    url = f"{ADO_BASE}/wit/workitems/${ADO_WORK_ITEM_TYPE}?api-version=7.1"
    resp = requests.post(url, json=patches, headers=ado_headers(), timeout=HTTP_TIMEOUT)
    if resp.status_code in (200, 201):
        wi = resp.json()
        # Check if Description was actually set in the response
        desc_val = wi.get("fields", {}).get("System.Description", "")
        log.info("Created work item #%s: %s  (Description in response: %d chars)",
                 wi["id"], title, len(desc_val or ""))
        return wi
    else:
        log.error("Failed to create work item: %s %s",
                  resp.status_code, resp.text)
        return None


def ado_add_comment(work_item_id, comment_html):
    """Append a comment to an existing ADO work item."""
    url = (f"{ADO_BASE}/wit/workItems/{work_item_id}"
           f"/comments?api-version=7.1-preview.4")
    payload = {"text": comment_html}
    hdrs = ado_headers()
    hdrs["Content-Type"] = "application/json"
    resp = requests.post(url, json=payload, headers=hdrs, timeout=HTTP_TIMEOUT)
    if resp.status_code in (200, 201):
        log.info("Added comment to work item #%s", work_item_id)
        return True
    else:
        log.error("Failed to add comment to #%s: %s %s",
                  work_item_id, resp.status_code, resp.text)
        return False


# ---------------------------------------------------------------------------
# Email helpers
# ---------------------------------------------------------------------------
def clean_subject(subject):
    """Strip RE:/FW:/Fwd: prefixes from subject."""
    return re.sub(r"^(RE:|Re:|FW:|Fw:|Fwd:)\s*", "", subject or "").strip()


def detect_source(subject, body):
    """Detect whether the email came from HubSpot, Teams, or regular Email."""
    if subject and "[HubSpot]" in subject:
        return "HubSpot"
    if body and "shared this from Microsoft Teams" in body:
        return "Teams"
    return "Email"


def send_confirmation_email(to_email, work_item_id, work_item_title,
                             is_update=False):
    """Send a confirmation email from triage@ennrgy.com."""
    action = "updated" if is_update else "created"
    wi_url = (f"https://dev.azure.com/{ADO_ORG}/{ADO_PROJECT}"
              f"/_workitems/edit/{work_item_id}")

    subject = f"[Triage] Work item #{work_item_id} {action}: {work_item_title}"
    body = (
        f"<p>Your triage request has been {action}.</p>"
        f"<p><strong>Work Item:</strong> #{work_item_id} — {work_item_title}</p>"
        f'<p><a href="{wi_url}">View in Azure DevOps</a></p>'
        f'<br><p style="color:#888;font-size:12px;">'
        f"This is an automated message from the Ennrgy triage system.</p>"
    )
    message = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": body},
            "toRecipients": [{"emailAddress": {"address": to_email}}],
        },
        "saveToSentItems": "false",
    }
    url = f"{GRAPH_BASE}/users/{TRIAGE_MAILBOX}/sendMail"
    resp = requests.post(url, json=message, headers=graph_headers(),
                         timeout=HTTP_TIMEOUT)
    if resp.status_code == 202:
        log.info("Confirmation email sent to %s for WI #%s",
                 to_email, work_item_id)
    else:
        log.error("Failed to send confirmation: %s %s",
                  resp.status_code, resp.text)


# ---------------------------------------------------------------------------
# Core processing: fetch new mail → create/update ADO item → confirm
# ---------------------------------------------------------------------------
def process_message(message_id):
    """Fetch a single mail message and process it into ADO."""
    log.info("Fetching message %s from Graph...", message_id)
    url = f"{GRAPH_BASE}/users/{TRIAGE_MAILBOX}/messages/{message_id}"
    params = {
        "$select": "id,subject,body,from,conversationId,receivedDateTime,hasAttachments"
    }
    resp = requests.get(url, params=params, headers=graph_headers(),
                        timeout=HTTP_TIMEOUT)
    if resp.status_code != 200:
        log.error("Could not fetch message %s: %s", message_id, resp.status_code)
        return

    msg = resp.json()
    subject        = msg.get("subject", "(no subject)")
    body_html_raw  = msg.get("body", {}).get("content", "")
    body_text      = re.sub(r"<[^>]+>", " ", body_html_raw)  # strip HTML for detection
    body_text      = re.sub(r"\s+", " ", body_text).strip()
    conversation_id = msg.get("conversationId", "")
    sender_email   = (msg.get("from", {}).get("emailAddress", {})
                      .get("address", ""))
    received_dt    = msg.get("receivedDateTime", "")

    # --- Skip our own confirmation emails to prevent infinite loops ---
    if sender_email.lower() == TRIAGE_MAILBOX.lower():
        log.info("Skipping self-sent message from %s: '%s'",
                 sender_email, subject)
        return

    if subject.startswith("[Triage]"):
        log.info("Skipping triage confirmation email: '%s'", subject)
        return

    # --- Deduplicate: skip if we already processed this message_id ---
    if _is_duplicate(message_id):
        log.info("Skipping duplicate notification for message %s", message_id)
        return

    cleaned_subj = clean_subject(subject)
    source       = detect_source(subject, body_text)
    log.info("Processing: '%s' from %s [source=%s]", subject, sender_email, source)

    # Thread matching: try ConversationId first, then cleaned subject
    existing_wi_id = ado_query_by_conversation_id(conversation_id)
    if not existing_wi_id:
        existing_wi_id = ado_query_by_subject(cleaned_subj)

    # --- Extract email attachments: inline images + file attachments ---
    has_attachments = msg.get("hasAttachments", False)
    if has_attachments:
        inline_images_html, file_attachments, cid_map = \
            process_email_attachments_for_inline(message_id)
    else:
        inline_images_html, file_attachments, cid_map = "", [], {}

    # Build a clean HTML description from the email content.
    # Outlook email HTML is too complex for ADO's renderer, so we
    # extract the text and images and rebuild simple HTML.
    clean_description = build_clean_description(
        body_html_raw, cid_map, inline_images_html)
    log.info("Clean description built: %d chars (raw email: %d chars, "
             "cid entries: %d)", len(clean_description),
             len(body_html_raw), len(cid_map))

    if existing_wi_id:
        # Update existing work item with a new comment
        comment = (
            f"<p><strong>New message from {sender_email}</strong> "
            f"({source}, {received_dt or 'unknown time'})</p>"
            f"<hr>{clean_description}"
        )
        ado_add_comment(existing_wi_id, comment)
        # Attach the actual email (.eml) to the work item
        attach_email_to_work_item(existing_wi_id, message_id, sender_email,
                                  cleaned_subj, received_dt)
        # Attach any non-image files
        for ado_url, fname in file_attachments:
            ado_attach_file_to_work_item(existing_wi_id, ado_url,
                                         f"Email attachment: {fname}")
        send_confirmation_email(sender_email, existing_wi_id, cleaned_subj,
                                is_update=True)
    else:
        # Create new work item with clean description
        description_html = clean_description
        title = (f"[{source}] {cleaned_subj}" if source != "Email"
                 else cleaned_subj)
        wi = ado_create_work_item(
            title=title,
            body_html=description_html,
            conversation_id=conversation_id,
            cleaned_subject=cleaned_subj,
            source=source,
            sender_email=sender_email,
        )
        if wi:
            # Attach the actual email (.eml) to the new work item
            attach_email_to_work_item(wi["id"], message_id, sender_email,
                                      cleaned_subj, received_dt)
            # Attach any non-image files
            for ado_url, fname in file_attachments:
                ado_attach_file_to_work_item(wi["id"], ado_url,
                                             f"Email attachment: {fname}")
            send_confirmation_email(sender_email, wi["id"], title,
                                    is_update=False)


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------
@app.route("/")
def index():
    return jsonify({"status": "ok", "service": "ennrgy-triage-webhook"})


@app.route("/webhook", methods=["POST"])
def webhook():
    """
    Microsoft Graph webhook endpoint.
    Handles:
      1. Subscription validation (responds with validationToken)
      2. Change notifications (fetches new mail and processes)
    """
    # --- Validation handshake ---
    validation_token = request.args.get("validationToken")
    if validation_token:
        log.info("Subscription validation request received")
        return validation_token, 200, {"Content-Type": "text/plain"}

    # --- Change notification ---
    try:
        payload = request.get_json(force=True)
    except Exception:
        return "Bad request", 400

    notifications = payload.get("value", [])
    log.info("Received %d notification(s)", len(notifications))

    for notif in notifications:
        resource    = notif.get("resource", "")
        change_type = notif.get("changeType", "")
        log.info("Notification: changeType=%s resource=%s",
                 change_type, resource)

        if change_type == "created" and "/messages/" in resource.lower():
            # resource looks like: "Users/{userId}/Messages/{messageId}"
            parts = re.split(r"/messages/", resource, flags=re.IGNORECASE)
            if len(parts) == 2:
                message_id = parts[1]
                log.info("Processing message_id=%s", message_id)
                try:
                    process_message(message_id)
                except Exception as e:
                    log.exception("Error processing message %s: %s",
                                  message_id, e)
        else:
            log.info("Skipping notification (changeType=%s, not a message "
                     "create)", change_type)

    # Always return 202 quickly to acknowledge receipt
    return "", 202


@app.route("/subscribe", methods=["POST"])
def subscribe():
    """
    Create (or renew) a Graph subscription for the triage mailbox.
    Call this endpoint once after deploying, then again before expiry.
    Graph subscriptions for mail last up to 4230 minutes (~2.9 days).
    """
    expiration = datetime.now(timezone.utc) + timedelta(minutes=4230)
    notification_url = (request.json.get("notificationUrl")
                        if request.is_json else None)
    if not notification_url:
        return jsonify({"error": "Provide notificationUrl in JSON body"}), 400

    body = {
        "changeType": "created",
        "notificationUrl": notification_url,
        "resource": f"users/{TRIAGE_MAILBOX}/messages",
        "expirationDateTime": expiration.strftime(
            "%Y-%m-%dT%H:%M:%S.0000000Z"),
        "clientState": "ennrgy-triage-secret",
    }
    resp = requests.post(
        f"{GRAPH_BASE}/subscriptions",
        json=body,
        headers=graph_headers(),
        timeout=HTTP_TIMEOUT,
    )
    if resp.status_code in (200, 201):
        sub = resp.json()
        log.info("Subscription created: id=%s, expires=%s",
                 sub["id"], sub["expirationDateTime"])
        return jsonify(sub), 201
    else:
        log.error("Subscription failed: %s %s", resp.status_code, resp.text)
        return jsonify({"error": resp.text}), resp.status_code


@app.route("/renew", methods=["POST"])
def renew():
    """
    Renew an existing subscription.
    Body: {"subscriptionId": "...", "notificationUrl": "..."}
    """
    data = request.get_json(force=True) if request.is_json else {}
    sub_id = data.get("subscriptionId")
    if not sub_id:
        return jsonify({"error": "Provide subscriptionId"}), 400

    expiration = datetime.now(timezone.utc) + timedelta(minutes=4230)
    body = {
        "expirationDateTime": expiration.strftime(
            "%Y-%m-%dT%H:%M:%S.0000000Z")
    }
    resp = requests.patch(
        f"{GRAPH_BASE}/subscriptions/{sub_id}",
        json=body,
        headers=graph_headers(),
        timeout=HTTP_TIMEOUT,
    )
    if resp.status_code == 200:
        log.info("Subscription renewed: %s", sub_id)
        return jsonify(resp.json()), 200
    else:
        log.error("Renew failed: %s %s", resp.status_code, resp.text)
        return jsonify({"error": resp.text}), resp.status_code


@app.route("/health")
def health():
    """Health check for Render."""
    return jsonify({
        "status": "healthy",
        "timestamp": datetime.now(timezone.utc).isoformat(),
    })


# ---------------------------------------------------------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port, debug=False)
