"""
Triage Email â Azure DevOps Work Item Webhook

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
ADO_PAT         = os.environ.get("ADO_PAT", "")          # Personal Access Token
ADO_WORK_ITEM_TYPE = os.environ.get("ADO_WORK_ITEM_TYPE", "Issue")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
ADO_BASE   = f"https://dev.azure.com/{ADO_ORG}/{ADO_PROJECT}/_apis"

HTTP_TIMEOUT = 30   # seconds for all outbound HTTP calls

# ---------------------------------------------------------------------------
# Logging â force flush after every message
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    stream=sys.stderr,
)
log = logging.getLogger(__name__)

# Ensure all handlers flush immediately
for handler in logging.root.handlers:
    if hasattr(dandler, "stream"):
        handler.stream = sys.stderr

# ---------------------------------------------------------------------------
# Flask app
# ---------------------------------------------------------------------------
app = Flask(__name__)

# ---------------------------------------------------------------------------
# Microsoft Graph auth (client credentials / daemon)
# ----------------------------------------------------------------------------
_graph_token_cache = {"token": None, "expires_at": 0}

dn().get("workItems", [])
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
    resp = requests.post(url, data=file_bytes, headers=hdrs,
                         timeout=HTTP_TIMEOUT)
    if resp.status_code in (200, 201):
        attachment_url = resp.json().get("url", "")
        log.info("Uploaded attachment '%s' â %s", file_name, attachment_url)
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
    resp = requests.patch(url, json=patches, headers=ado_headers(),
                          timeout=HTTP_TIMEOUT)
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

    comment = (f"Email from {sender_email} â "
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
    re
