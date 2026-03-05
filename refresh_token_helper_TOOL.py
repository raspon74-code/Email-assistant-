"""
OAS Token & Cookie Refresh Helper Tool
=======================================

Run this script to update your OAS authentication credentials when the API
returns HTTP 401 or HTTP 500 errors.

HOW TO GET YOUR CREDENTIALS FROM THE BROWSER:
----------------------------------------------

Step 1 — Get the OAS-TOKEN value:
  1. Open your browser and navigate to the OAS portal.
  2. Open DevTools (F12) → Network tab.
  3. Make a successful authenticated request (e.g. load the vessels page).
  4. Click on the request → Headers → Request Headers → find "Cookie:".
  5. In the Cookie header value, locate the segment that starts with "OAS-TOKEN=".
  6. Copy ONLY the token value (the part after "OAS-TOKEN=" up to the next ";").
  7. Paste it when prompted below.

Step 2 — Get the full Cookie header:
  Option A (recommended):
    1. In DevTools → Network, right-click the authenticated request.
    2. Select "Copy" → "Copy as fetch".
    3. In the copied text, find the "cookie" header value.
    4. Copy the ENTIRE value string (everything between the quotes after "cookie":).

  Option B:
    1. In DevTools → Network → click the request → Headers tab.
    2. Find the "Cookie:" request header.
    3. Copy the ENTIRE value line (all cookies separated by "; ").

  Paste the full cookie string when prompted below.

Files saved:
  oas_token.txt    — contains the bare OAS-TOKEN value
  oas_cookies.txt  — contains the full cookie header string
"""

import os
import sys
import requests

OAS_TOKEN_FILE = "oas_token.txt"
OAS_COOKIES_FILE = "oas_cookies.txt"

# OAS test endpoint used to validate credentials
OAS_VALIDATE_URL = "https://oas.shell.com/api/vessels"
OAS_VALIDATE_PARAMS = {"site": "PERNIS"}


def parse_cookie_string(cookie_str):
    """
    Parse a raw 'key=value; key2=value2' cookie header string into a dict.

    Handles the full cookie header value as copied from browser DevTools.
    Leading 'Cookie: ' prefix is stripped automatically if present.
    """
    cookies = {}
    if not cookie_str:
        return cookies

    # Strip leading 'Cookie: ' prefix if the user pasted the full header line
    stripped = cookie_str.strip()
    if stripped.lower().startswith("cookie:"):
        stripped = stripped[len("cookie:"):].strip()

    for part in stripped.split(";"):
        part = part.strip()
        if not part:
            continue
        if "=" in part:
            key, _, value = part.partition("=")
            cookies[key.strip()] = value.strip()
        else:
            # Cookie with no value
            cookies[part] = ""

    return cookies


def prompt_multiline(prompt_text):
    """
    Prompt the user for input that may span multiple lines.
    The user signals completion by entering a blank line.
    """
    print(prompt_text)
    print("(Press Enter twice when done)")
    lines = []
    while True:
        try:
            line = input()
        except EOFError:
            break
        if line == "" and lines:
            break
        lines.append(line)
    return " ".join(lines).strip()


def validate_credentials(cookies_dict):
    """
    Send a test request to the OAS API using the full cookie bundle.
    Returns (success: bool, status_code: int, message: str).
    """
    try:
        response = requests.get(
            OAS_VALIDATE_URL,
            params=OAS_VALIDATE_PARAMS,
            cookies=cookies_dict,
            # Bypass corporate proxy: OAS is an internal Shell service reached directly
            proxies={"http": None, "https": None},
            timeout=15,
            verify=False,
        )
        if response.status_code == 200:
            return True, 200, "✅ Credentials validated successfully."
        else:
            return False, response.status_code, f"❌ Server returned HTTP {response.status_code}: {response.text[:200]}"
    except Exception as exc:
        return False, 0, f"❌ Request failed: {exc}"


def main():
    print("=" * 60)
    print("  OAS Token & Cookie Refresh Helper")
    print("=" * 60)
    print()
    print("This tool updates the OAS authentication credentials used by")
    print("the Email Assistant to fetch live vessel data from the OAS API.")
    print()

    # ── Step 1: OAS-TOKEN value ──────────────────────────────────────────────
    print("-" * 60)
    print("STEP 1 — Paste the OAS-TOKEN value")
    print("-" * 60)
    print("In your browser DevTools → Network → [authenticated request]")
    print("→ Headers → Cookie → find 'OAS-TOKEN=<value>' → copy the VALUE only.")
    print()

    token_value = ""
    while not token_value:
        try:
            token_value = input("Paste OAS-TOKEN value here: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nCancelled.")
            sys.exit(0)
        if not token_value:
            print("Token cannot be empty. Please try again.")

    # ── Step 2: Full Cookie header string ────────────────────────────────────
    print()
    print("-" * 60)
    print("STEP 2 — Paste the FULL Cookie header string")
    print("-" * 60)
    print("Option A (recommended):")
    print("  DevTools → Network → right-click request → Copy → Copy as fetch")
    print("  Then copy the entire value after  \"cookie\": ")
    print()
    print("Option B:")
    print("  DevTools → Network → click request → Headers tab")
    print("  Find 'Cookie:' request header → copy the ENTIRE value line")
    print("  (all cookies separated by '; ')")
    print()
    print("The string should look like:")
    print("  OAS-TOKEN=abc123; _ga=GA1.2.xxx; dtCookie=yyy; ...")
    print()

    cookie_str = ""
    while not cookie_str:
        cookie_str = prompt_multiline("Paste full Cookie header value here:")
        if not cookie_str:
            print("Cookie string cannot be empty. Please try again.")
            print()

    # Parse and verify that OAS-TOKEN is present in the cookie string
    cookies_dict = parse_cookie_string(cookie_str)

    if "OAS-TOKEN" not in cookies_dict:
        print()
        print("⚠️  WARNING: 'OAS-TOKEN' not found in the pasted cookie string.")
        print("   Injecting the token value you provided in Step 1.")
        cookies_dict["OAS-TOKEN"] = token_value
    elif cookies_dict.get("OAS-TOKEN") != token_value:
        print()
        print("ℹ️  NOTE: The OAS-TOKEN in the cookie string differs from Step 1.")
        print("   Using the cookie-string value (it may be more up-to-date).")
        token_value = cookies_dict["OAS-TOKEN"]

    # ── Validate credentials ─────────────────────────────────────────────────
    print()
    print("-" * 60)
    print("Validating credentials against OAS API …")
    print("-" * 60)

    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    success, status_code, message = validate_credentials(cookies_dict)
    print(message)

    if not success:
        print()
        print("⚠️  Validation failed. The credentials may still work in the")
        print("   Email Assistant. Saving anyway — you can retry later.")
        proceed = ""
        while proceed not in ("y", "n"):
            try:
                proceed = input("Save credentials anyway? [y/n]: ").strip().lower()
            except (EOFError, KeyboardInterrupt):
                print("\nCancelled.")
                sys.exit(0)
        if proceed == "n":
            print("Credentials NOT saved.")
            sys.exit(1)

    # ── Save files ────────────────────────────────────────────────────────────
    try:
        with open(OAS_TOKEN_FILE, "w", encoding="utf-8") as f:
            f.write(token_value)
        print(f"\n✅ Token saved  → {OAS_TOKEN_FILE}")
    except Exception as exc:
        print(f"\n❌ Failed to save token: {exc}")
        sys.exit(1)

    # Save the raw cookie string (strip leading 'Cookie: ' prefix if present)
    save_str = cookie_str.strip()
    if save_str.lower().startswith("cookie:"):
        save_str = save_str[len("cookie:"):].strip()

    try:
        with open(OAS_COOKIES_FILE, "w", encoding="utf-8") as f:
            f.write(save_str)
        print(f"✅ Cookies saved → {OAS_COOKIES_FILE}")
    except Exception as exc:
        print(f"\n❌ Failed to save cookies: {exc}")
        sys.exit(1)

    print()
    print("=" * 60)
    print("  Done! The Email Assistant will use these credentials")
    print("  on the next run.")
    print("=" * 60)


if __name__ == "__main__":
    main()
