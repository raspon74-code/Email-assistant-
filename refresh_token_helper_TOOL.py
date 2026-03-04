"""
refresh_token_helper_TOOL.py — Standalone OAS Token Validator & Saver

Run this script each morning (or whenever the OAS session expires) to
refresh the token used by the Email Assistant:

    python refresh_token_helper_TOOL.py

Steps:
  1. Open the OAS Shipment Planner in your browser and log in.
  2. Open DevTools → Application → Cookies and copy the value of OAS-TOKEN.
  3. Paste it when prompted by this script.
  4. The script validates the token against the OAS API and saves it to
     oas_token.txt if valid.
"""

import os
import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# =========================================================
# OAS API CONFIGURATION  (must match email_assistant_v21_FINAL.py)
# =========================================================
OAS_API_URL     = "https://nl.oas.shell.com/OASV6/coreapi/api/oard1000/GetExplorerShipmentSummary"
OAS_SITE        = "PERNIS"
OAS_STATUS_FROM = 5
OAS_STATUS_TO   = 10

TOKEN_FILE = "oas_token.txt"


def validate_and_save_token(token: str) -> bool:
    """POST to the OAS API with the supplied token.

    Returns True and saves the token to *TOKEN_FILE* when the API accepts it.
    Returns False and prints a clear error message otherwise.
    """
    payload = {
        "Agent": None,
        "Company": None,
        "DateFrom": None,
        "DateTo": None,
        "Inspector": None,
        "Port": None,
        "ShipmentId": None,
        "ShipmentRef": None,
        "Site": OAS_SITE,
        "StatusFrom": OAS_STATUS_FROM,
        "StatusTo": OAS_STATUS_TO,
        "Timestamp": None,
        "TransportName": None,
    }

    try:
        response = requests.post(
            OAS_API_URL,
            json=payload,
            headers={
                "Content-Type": "application/json",
                "Origin": "https://nl.oas.shell.com",
                "User-Agent": "Mozilla/5.0",
            },
            cookies={"OAS-TOKEN": token},
            proxies={"http": None, "https": None},  # bypass corporate proxy
            timeout=30,
            verify=False,
        )
    except requests.exceptions.RequestException as exc:
        print(f"\n❌ Network error while contacting OAS API: {exc}")
        return False

    if response.status_code == 200:
        try:
            data = response.json()
        except ValueError:
            data = None

        # A 200 with an error body (e.g. {"error": ...}) is treated as invalid
        if isinstance(data, dict) and data.get("error"):
            print(f"\n❌ OAS API returned an error body: {data.get('error')}")
            return False

        # Token is valid — save it
        with open(TOKEN_FILE, "w", encoding="utf-8") as fh:
            fh.write(token)
        print(f"\n✅ Token is valid!  Saved to '{TOKEN_FILE}'.")
        if data:
            print(f"   API returned {len(data)} shipment record(s).")
        return True

    # Non-200 responses
    if response.status_code in (401, 403):
        print(
            f"\n❌ Token rejected by OAS API (HTTP {response.status_code}).\n"
            "   Please log in to the OAS portal again and copy a fresh OAS-TOKEN."
        )
    elif response.status_code == 500:
        body_preview = response.text[:300] if response.text else "(empty)"
        print(
            f"\n❌ OAS API returned HTTP 500.\n"
            f"   Response body: {body_preview}\n"
            "   The token may be invalid or the API may be temporarily unavailable."
        )
    else:
        print(
            f"\n❌ Unexpected response from OAS API (HTTP {response.status_code}).\n"
            f"   Response: {response.text[:300]}"
        )
    return False


if __name__ == "__main__":
    print("=" * 60)
    print("  OAS Token Validator & Saver — Email Assistant v21.0 FINAL")
    print("=" * 60)
    print(
        "\nHow to get your OAS-TOKEN:\n"
        "  1. Open https://nl.oas.shell.com in your browser and log in.\n"
        "  2. Press F12 → Application tab → Cookies → nl.oas.shell.com\n"
        "  3. Copy the VALUE of the 'OAS-TOKEN' cookie.\n"
    )

    token_input = input("Paste your OAS-TOKEN value here: ").strip()

    if not token_input:
        print("\n❌ No token entered.  Exiting.")
    else:
        validate_and_save_token(token_input)
