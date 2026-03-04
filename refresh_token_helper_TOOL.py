"""
refresh_token_helper_TOOL.py — Standalone OAS Token Refresh Tool
-----------------------------------------------------------------
Run this script each morning before starting the Email Assistant.

Usage:
    python refresh_token_helper_TOOL.py

1. Open the OAS portal in your browser and log in.
2. Open DevTools (F12) → Application → Cookies → find OAS-TOKEN.
3. Copy the token value and paste it here when prompted.

The script validates the token against the OAS API and saves it to
oas_token.txt for use by email_assistant_v21_FINAL.py.
"""

import os
import sys
import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# =========================================================
# OAS CONFIGURATION (must match email_assistant_v21_FINAL.py)
# =========================================================

OAS_API_URL     = "https://nl.oas.shell.com/OASV6/coreapi/api/oard1000/GetExplorerShipmentSummary"
OAS_SITE        = "PERNIS"
OAS_STATUS_FROM = 5
OAS_STATUS_TO   = 10
TOKEN_FILE      = "oas_token.txt"

# Validation payload — minimal date window just to test the token
_VALIDATE_PAYLOAD = {
    "Agent": None, "Company": None,
    "DateFrom": None, "DateTo": None,
    "Inspector": None, "Port": None,
    "ShipmentId": None, "ShipmentRef": None,
    "Site": OAS_SITE,
    "StatusFrom": OAS_STATUS_FROM, "StatusTo": OAS_STATUS_TO,
    "Timestamp": None, "TransportName": None
}


def _clear_proxy():
    """Remove proxy environment variables so OAS calls bypass the corporate proxy."""
    for var in ("HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy"):
        os.environ.pop(var, None)


def validate_token(token: str) -> tuple[bool, str]:
    """
    POST to OAS API with the given token.
    Returns (True, "") on success or (False, error_message) on failure.
    """
    _clear_proxy()
    try:
        response = requests.post(
            OAS_API_URL,
            json=_VALIDATE_PAYLOAD,
            headers={
                "Content-Type": "application/json",
                "Origin": "https://nl.oas.shell.com",
                "User-Agent": "Mozilla/5.0",
            },
            cookies={"OAS-TOKEN": token},
            proxies={"http": None, "https": None},
            timeout=30,
            verify=False,
        )

        if response.status_code in (401, 403):
            return False, f"HTTP {response.status_code} — token is invalid or expired."

        response.raise_for_status()
        return True, ""

    except requests.exceptions.ConnectionError as e:
        return False, f"Connection error: {e}"
    except requests.exceptions.Timeout:
        return False, "Request timed out (30 s). Check network connectivity."
    except requests.exceptions.HTTPError as e:
        return False, f"HTTP error: {e}"
    except Exception as e:
        return False, f"Unexpected error: {e}"


def save_token(token: str):
    """Write the validated token to oas_token.txt."""
    with open(TOKEN_FILE, "w") as f:
        f.write(token)


def main():
    print("=" * 60)
    print("  OAS Token Refresh Tool")
    print("=" * 60)
    print()
    print("Steps:")
    print("  1. Open https://nl.oas.shell.com in your browser and log in.")
    print("  2. Press F12 → Application tab → Cookies → nl.oas.shell.com")
    print("  3. Find 'OAS-TOKEN', copy its Value.")
    print()

    while True:
        try:
            token = input("Paste OAS-TOKEN value (or 'q' to quit): ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nAborted.")
            sys.exit(0)

        if token.lower() == "q":
            print("Aborted.")
            sys.exit(0)

        if not token:
            print("❌ No token entered. Please try again.\n")
            continue

        print(f"\n🔍 Validating token against OAS API …")
        ok, err = validate_token(token)

        if ok:
            save_token(token)
            print(f"✅ Token is valid! Saved to '{TOKEN_FILE}'.")
            print("   You can now start the Email Assistant.")
            sys.exit(0)
        else:
            print(f"❌ Validation failed: {err}")
            print("   Please check the token and try again.\n")


if __name__ == "__main__":
    main()
