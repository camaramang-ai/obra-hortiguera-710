#!/usr/bin/env python3
import os, base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

t = os.path.expanduser("~/.hermes/google_token.json")
creds = Credentials.from_authorized_user_file(t, [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/gmail.send",
])
if creds.expired and creds.refresh_token:
    creds.refresh(Request())
svc = build("gmail", "v1", credentials=creds)

def send(to, subject, body):
    msg = MIMEMultipart("alternative")
    msg["To"] = to
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    sent = svc.users().messages().send(userId="me", body={"raw": raw}).execute()
    print(f"Enviado a {to} - ID: {sent['id']}")

# ── 1. Caballito Sanitarios ──
send(
    "info@caballitosanitarios.com.ar",
    "Consulta por sanitarios Ferrum y griferia FV para obra en Caballito",
    """Hola, como estan?

Les escribo porque estoy haciendo una remodelacion en Caballito y queria pedirles un presupuesto. Soy Pablo.

  * Inodoro Largo Ferrum Bari Blanco (1 un)
  * Deposito Ferrum Bari Dual de Apoyar (1 un)
  * Bidet Ferrum Bari 3 Agujeros (1 un)
  * Tapa Asiento Ferrum Bari Madera (1 un)
  * Schneider Vanitory City Colgar 80cm (1 un)
  * Schneider Vanitory Rivo Colg 60cm Blanco (1 un)
  * FV Lavatorio Mauna Cromo (3 JG)
  * FV Bidet Mauna Cromo (2 JG)
  * FV Ducha Mana c/Transferencia Cromo (2 JG)
  * FV Duchamatic c/Brazo Cromo (2 un)
  * FV Banera Enoxado 150x70 (1 un)
  * FV Desague Lineal 60cm (1 un)

Si de algunos no tienen el producto exacto, mandenme opciones equivalentes.

Me interesa saber: precio unitario, stock disponible, envio a Caballito, descuento por efectivo/transferencia, y si aceptan cuotas.

Desde ya muchas gracias.

Saludos,
Pablo"""
)

# ── 2. Generacion Sanitaria ──
send(
    "gsrosario@yahoo.com",
    "Consulta por sanitarios Ferrum para obra en Caballito",
    """Hola, como estan?

Les escribo porque estoy haciendo una remodelacion en Caballito, CABA, y queria pedirles presupuesto para el set completo de sanitarios Ferrum Bari. Soy Pablo.

  * Inodoro Largo Ferrum Bari Blanco (1 un)
  * Deposito Ferrum Bari Dual de Apoyar (1 un)
  * Bidet Ferrum Bari 3 Agujeros (1 un)
  * Tapa Asiento Ferrum Bari Madera (1 un)

Si no tienen alguno exacto, opciones equivalentes.

Me interesa saber: precio, stock, envio a Caballito, descuento por efectivo o transferencia.

Desde ya muchas gracias.

Saludos,
Pablo"""
)

print("Listo.")
