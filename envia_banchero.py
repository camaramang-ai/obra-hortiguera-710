#!/usr/bin/env python3
import os, json, base64
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

items_text = """  \u2022 Inodoro Largo Ferrum Bari Blanco (1 un)
  \u2022 Deposito Ferrum Bari Dual de Apoyar (1 un)
  \u2022 Bidet Ferrum Bari 3 Agujeros (1 un)
  \u2022 Tapa Asiento Ferrum Bari Madera (1 un)
  \u2022 FV Lavatorio Mauna Cromo (3 JG)
  \u2022 FV Bidet Mauna Cromo (2 JG)
  \u2022 FV Ducha Mana c/Transferencia Cromo (2 JG)
  \u2022 FV Duchamatic c/Brazo Cromo (2 un)
  \u2022 FV Banera Enoxado 150x70 (1 un)
  \u2022 FV Desague Lineal 60cm (1 un)
  \u2022 FV Portarrollos Cromo (3 un)
  \u2022 FV Toallero Barra Cromo (3 un)
  \u2022 FV Percha Cromo (5 un)
  \u2022 Eliane Ceramica Forma Branco 32x60 (30.1 m2)
  \u2022 Espejo Reflejar Rectang. 100x80 4mm (1 un)
  \u2022 Espejo 20x30 Bordes Acabado Ilumin Oxboro 220V (1 un)
  \u2022 Indoor Largo GAP Blanco Roca (1 un)
  \u2022 Alimentacion Efectiva Cond. Bajo Exc. 3/6 L (1 un)
  \u2022 Bidet 3 Agujeros Roca La Gama (1 un)
  \u2022 Tapa Amortiguador Slim (1 un)
  \u2022 Portobello Zesty Light Grate Mateo 45x120 Text (6.44 m2)"""

body = f"""Hola, como estan?

Les escribo porque estoy haciendo una remodelacion en Caballito y queria pedirles un presupuesto para los siguientes productos. Soy Pablo.

{items_text}

Si de algunos no tienen el producto exacto, por favor mandenme opciones equivalentes que puedan ofrecer.

Me interesaria saber:
  \u2022 Precio unitario y total de cada producto
  \u2022 Stock disponible
  \u2022 Si hacen envio a Caballito y cuanto sale el flete
  \u2022 Si tienen descuento por pago en efectivo o transferencia bancaria
  \u2022 Si aceptan cuotas o financiacion

Desde ya muchas gracias.

Saludos,
Pablo"""

msg = MIMEMultipart("alternative")
msg["To"] = "ventas@bancherosanitarios.com.ar"
msg["Subject"] = "Consulta por sanitarios, griferia FV y revestimientos para obra en Caballito"
msg.attach(MIMEText(body, "plain", "utf-8"))
raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
sent = svc.users().messages().send(userId="me", body={"raw": raw}).execute()
print(f"Enviado a Banchero Sanitarios - ID: {sent['id']}")
print(f"To: ventas@bancherosanitarios.com.ar")
