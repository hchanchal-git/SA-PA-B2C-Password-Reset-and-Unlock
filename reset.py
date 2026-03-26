# Created By Himanshu Kumar (Cloud Ops)

import requests
import random
import datetime
import platform
import subprocess
import sys
import time

# ==================================================
# Azure AD B2C - SA / PA Password Reset & Unlock
# ==================================================

print("=" * 50)
print(" Azure AD B2C - SA / PA Password Reset & Unlock ")
print("=" * 50)

# ---------------- INPUT ----------------
TENANT_ID = input("B2C_TENANT_ID_USER (*.onmicrosoft.com): ").strip()
CLIENT_ID = input("B2C_APP_ID_USER (Client ID): ").strip()
CLIENT_SECRET = input("SECRET_VALUE_USER (Client Secret): ").strip()
EGAIN_URL = input("Enter eGain APP URL: ").strip()

ISSUER = TENANT_ID
choice = input("Reset which user? (sa / pa / both): ").lower().strip()

if choice == "sa":
    USERS = ["sa"]
elif choice == "pa":
    USERS = ["pa"]
elif choice == "both":
    USERS = ["sa", "pa"]
else:
    print("❌ Invalid choice")
    sys.exit(1)

# ---------------- HELPERS ----------------
def utc_now():
    return datetime.datetime.now(datetime.timezone.utc).strftime(
        "%Y-%m-%dT%H:%M:%S.%fZ"
    )

def new_strong_password():
    chars = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789@#$%!"
    return "".join(random.choice(chars) for _ in range(16))

# ---------------- TOKEN ----------------
print("\nGenerating Graph API Token...")

token_resp = requests.post(
    f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token",
    data={
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
    },
)

token_resp.raise_for_status()
token = token_resp.json()

expiry = datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(
    seconds=token["expires_in"]
)

print("Graph Token   : GENERATED SUCCESSFULLY")
print(f"Token Expires : {expiry} UTC")

HEADERS = {
    "Authorization": f"Bearer {token['access_token']}",
    "Content-Type": "application/json",
}

# ---------------- MAIL ----------------
def send_outlook_mail(subject, body):
    os_name = platform.system()

    # -------- WINDOWS --------
    if os_name == "Windows":
        try:
            import win32com.client
            import pythoncom

            pythoncom.CoInitialize()

            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            namespace.Logon("", "", False, False)

            mail = outlook.CreateItem(0)
            mail.To = "cbu-tsops@egain.com"
            mail.Subject = subject
            mail.Body = body

            mail.Send()

            # Give Outlook time to push mail
            time.sleep(2)

            return "SENT (Windows Outlook)"

        except Exception as e:
            return f"FAILED (Windows): {e}"

    # -------- macOS --------
    elif os_name == "Darwin":
        try:
            applescript = f'''
            tell application "Microsoft Outlook"
                activate
                set newMessage to make new outgoing message with properties {{subject:"{subject}", content:"{body}"}}
                tell newMessage
                    make new recipient at end of to recipients with properties {{email address:{{address:"cbu-tsops@egain.com"}}}}
                    send
                end tell
            end tell
            '''
            subprocess.run(["osascript", "-e", applescript], check=True)
            return "SENT (Mac Outlook)"

        except Exception as e:
            return f"FAILED (Mac): {e}"

    else:
        return "FAILED (Unsupported OS)"

# ---------------- PROCESS USERS ----------------
for username in USERS:
    print("\n" + "-" * 50)
    print(f"Processing User : {username}")

    user_lookup = requests.get(
        "https://graph.microsoft.com/beta/users",
        headers=HEADERS,
        params={
            "$filter": f"identities/any(c:c/issuerAssignedId eq '{username}' and c/issuer eq '{ISSUER}')"
        },
    )

    user_lookup.raise_for_status()
    users = user_lookup.json().get("value", [])

    if not users:
        print("User Status    : NOT FOUND")
        continue

    user = users[0]
    user_id = user["id"]

    print(f"User ID        : {user_id}")
    print("Account Status : ENABLED")

    last_login_prop = next(
        (k for k in user.keys() if k.startswith("extension_") and k.endswith("lastLoginDateTime")),
        None,
    )

    old_last_login = user.get(last_login_prop) if last_login_prop else None
    if not old_last_login:
        old_last_login = "NEVER LOGGED IN"

    print(f"Old Last Login : {old_last_login}")

    pwd_input = input("Enter new password (ENTER = auto): ").strip()
    password = pwd_input if pwd_input else new_strong_password()

    new_last_login = utc_now()

    reset_body = {
        "accountEnabled": True,
        "passwordProfile": {
            "password": password,
            "forceChangePasswordNextSignIn": False,
        },
        "passwordPolicies": "DisablePasswordExpiration, DisableStrongPassword",
    }

    if last_login_prop:
        reset_body[last_login_prop] = new_last_login

    requests.patch(
        f"https://graph.microsoft.com/beta/users/{user_id}",
        headers=HEADERS,
        json=reset_body,
    ).raise_for_status()

    print("Password Reset : SUCCESS")
    print("Account Unlock : SUCCESS")
    print(f"New Last Login : {new_last_login}")

    mail_body = f"""
================ EXECUTION DETAILS =================

eGain URL Referenced     : {EGAIN_URL}
Tenant ID Used           : {TENANT_ID}
Old Last Login           : {old_last_login}
New Last Login           : {new_last_login} (UTC)
Issuer Used              : {TENANT_ID}
Client ID Used           : {CLIENT_ID}

Account Unlock           : SUCCESS
Password Reset           : SUCCESS

Graph API Version        : beta
Graph Scope              : https://graph.microsoft.com/.default

Users Selected           : {", ".join(USERS)}
Processed User           : {username}
Azure AD User ID         : {user_id}

Mail Method              : Outlook Desktop

====================================================
"""

    mail_status = send_outlook_mail(
        "SA/PA B2C User Password Reset & Unlock - Execution Summary",
        mail_body,
    )

    print(f"Mail Status    : {mail_status}")

print("\n" + "=" * 50)
print(" ALL REQUESTS COMPLETED SUCCESSFULLY ")
print("=" * 50)
