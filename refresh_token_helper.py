"""
OAS Token Refresh Helper
========================
Run this script each morning (or when the agent reports token expired).
Takes less than 30 seconds.

Usage:
    python refresh_token_helper.py
"""

import os
import sys
import webbrowser
import tkinter as tk
from tkinter import messagebox, scrolledtext

TOKEN_FILE = "oas_token.txt"
OAS_URL = "https://nl.oas.shell.com/OASV6"

def get_existing_token():
    if os.path.exists(TOKEN_FILE):
        try:
            with open(TOKEN_FILE) as f:
                t = f.read().strip()
            return t if t else None
        except Exception:
            return None
    return None

def save_token(token):
    with open(TOKEN_FILE, "w") as f:
        f.write(token.strip())

def main():
    root = tk.Tk()
    root.title("OAS Token Refresh Helper")
    root.geometry("600x420")
    root.resizable(False, False)

    # Center the window
    root.eval('tk::PlaceWindow . center')

    tk.Label(root, text="🔐 OAS Token Refresh Helper", font=("Segoe UI", 14, "bold")).pack(pady=(18, 4))
    tk.Label(root, text="Follow these steps to refresh your OAS session token:", font=("Segoe UI", 10)).pack()

    steps = (
        "1.  Click 'Open OAS in Browser' below\n"
        "2.  OAS will open — press F12 to open DevTools\n"
        "3.  Click the 'Network' tab → make any search in OAS\n"
        "4.  Find 'GetExplorerShipmentSummary' in the request list\n"
        "5.  Click it → click 'Headers' tab\n"
        "6.  Find the 'Cookie:' header → copy the OAS-TOKEN value\n"
        "    (the long string after 'OAS-TOKEN=')\n"
        "7.  Paste it in the box below and click Save"
    )
    tk.Label(root, text=steps, font=("Segoe UI", 9), justify="left", anchor="w",
             bg="#f0f4f8", relief="groove", padx=10, pady=8).pack(fill="x", padx=18, pady=10)

    tk.Button(
        root, text="🌐 Open OAS in Browser", font=("Segoe UI", 10, "bold"),
        bg="#0078d4", fg="white", relief="flat", padx=10, pady=5,
        command=lambda: webbrowser.open(OAS_URL)
    ).pack(pady=(0, 8))

    tk.Label(root, text="Paste OAS-TOKEN value here:", font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=18)
    token_box = scrolledtext.ScrolledText(root, height=4, font=("Consolas", 8), wrap="word")
    token_box.pack(fill="x", padx=18, pady=(2, 10))

    # Pre-fill with existing token if available
    existing = get_existing_token()
    if existing:
        token_box.insert("1.0", existing)

    def on_save():
        token = token_box.get("1.0", "end").strip()
        if not token:
            messagebox.showwarning("Empty", "Please paste the OAS-TOKEN value first.")
            return
        if len(token) < 50:
            messagebox.showwarning("Too short", "This doesn't look like a valid token (too short).\nMake sure you copied the full value.")
            return
        save_token(token)
        messagebox.showinfo("✅ Saved!", f"Token saved to {TOKEN_FILE}\n\nThe Email Assistant will use it automatically.\nToken is valid for today's session.")
        root.destroy()

    tk.Button(
        root, text="💾 Save Token", font=("Segoe UI", 11, "bold"),
        bg="#107c10", fg="white", relief="flat", padx=16, pady=6,
        command=on_save
    ).pack(pady=4)

    root.mainloop()

if __name__ == "__main__":
    main()
