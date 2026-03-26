"""
Helper script to create or renew the Microsoft Graph webhook subscription.
Run this once after deploying to Render, then set up a cron to renew every 2 days.

Usage:
  python renew_subscription.py create https://your-app.onrender.com/webhook
  python renew_subscription.py renew <subscription_id>
"""

import sys
import os
import requests

RENDER_APP_URL = os.environ.get("RENDER_APP_URL", "")


def create(notification_url):
    """Create a new subscription."""
    url = f"{RENDER_APP_URL}/subscribe"
    resp = requests.post(url, json={"notificationUrl": notification_url})
    print(f"Status: {resp.status_code}")
    print(resp.json())
    if resp.status_code == 201:
        sub = resp.json()
        print(f"\n*** SAVE THIS SUBSCRIPTION ID: {sub['id']} ***")
        print(f"Expires: {sub['expirationDateTime']}")


def renew(subscription_id):
    """Renew an existing subscription."""
    url = f"{RENDER_APP_URL}/renew"
    resp = requests.post(url, json={"subscriptionId": subscription_id})
    print(f"Status: {resp.status_code}")
    print(resp.json())


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    action = sys.argv[1]
    if action == "create":
        if len(sys.argv) < 3:
            print("Provide notificationUrl: python renew_subscription.py create https://your-app.onrender.com/webhook")
            sys.exit(1)
        if not RENDER_APP_URL:
            RENDER_APP_URL = sys.argv[2].rsplit("/webhook", 1)[0]
        create(sys.argv[2])
    elif action == "renew":
        if len(sys.argv) < 3:
            print("Provide subscriptionId")
            sys.exit(1)
        renew(sys.argv[2])
    else:
        print(f"Unknown action: {action}")
