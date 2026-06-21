#!/usr/bin/env python3
"""Re-authenticate Gmail API — metodo local server con puerto fijo."""
import os, json
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

TK = os.path.expanduser("~/.hermes/google_token.json")
CF = os.path.expanduser(
    "~/Downloads/client_secret_311076682142-31ufbb0dc57p6fg8rtb6hc8tjal8aa9a.apps.googleusercontent.com.json"
)

with open(CF) as f:
    client_config = json.load(f)
cinst = client_config.setdefault("installed", {})
cinst["redirect_uris"] = ["http://localhost:8080", "http://localhost"]

SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/gmail.send",
]

if os.path.exists(TK):
    creds = Credentials.from_authorized_user_file(TK, SCOPES)
    if creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            with open(TK, "w") as f:
                f.write(creds.to_json())
            print("Token refreshed!")
            import sys; sys.exit(0)
        except:
            print("Refresh failed, re-authenticating...")

print("Abriendo navegador en http://localhost:8080 ...")
flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
creds = flow.run_local_server(port=8080, open_browser=True)

with open(TK, "w") as f:
    f.write(creds.to_json())
print("Token guardado:", TK)

svc = build("gmail", "v1", credentials=creds)
p = svc.users().getProfile(userId="me").execute()
print("Conectado como:", p["emailAddress"])
