#!/usr/bin/env python3
"""
Agente de Compras — Obra Hortiguera 710
Contacta proveedores por email, consulta precios, disponibilidad y formas de pago.
Compila resultados y sugiere la mejor opción.

Uso:
  python3 agente-compras.py listar-proveedores
  python3 agente-compras.py enviar --proveedor barugel
  python3 agente-compras.py enviar --proveedor todos
  python3 agente-compras.py enviar --proveedor barugel --items D.1,D.2,K.1
  python3 agente-compras.py revisar
  python3 agente-compras.py resumen
  python3 agente-compras.py estado
"""

import os, sys, json, argparse, re, textwrap
from datetime import datetime, timezone
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import base64

# ── Gmail API ──────────────────────────────────────────────────────────
TOKEN_PATH = os.path.expanduser("~/.hermes/google_token.json")
CLIENT_SECRET_PATH = os.path.expanduser(
    "~/Downloads/client_secret_311076682142-31ufbb0dc57p6fg8rtb6hc8tjal8aa9a.apps.googleusercontent.com.json"
)

SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/gmail.send",
]

def get_gmail_service():
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build

    creds = None
    if os.path.exists(TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_PATH, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_PATH, "w") as f:
            f.write(creds.to_json())
    return build("gmail", "v1", credentials=creds)


# ── ESTADO ─────────────────────────────────────────────────────────────
STATE_FILE = os.path.expanduser(
    "~/.hermes/agente-compras-estado.json"
)

def load_state():
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE) as f:
            return json.load(f)
    return {"enviados": [], "recibidos": [], "resultados": {}}

def save_state(state):
    os.makedirs(os.path.dirname(STATE_FILE), exist_ok=True)
    with open(STATE_FILE, "w") as f:
        json.dump(state, f, indent=2, ensure_ascii=False)


# ── PROVEEDORES ────────────────────────────────────────────────────────
PROVEEDORES = {
    "barugel": {
        "nombre": "Barugel Azulay & Cía",
        "email": "mbadillo@barugel.com.ar",
        "atencion": "Marcos Badillo",
        "rubro": "Revestimientos, sanitarios, grifería",
        "sucursal": "Floresta (Alberdi 3701)",
        "productos": ["D.1","D.2","D.3","D.4","D.5","D.6","D.7",
                      "K.1","P.1","P.2","P.3",
                      "S.1","S.2","S.3","S.4",
                      "V.1","V.2","B.1","B.2",
                      "G.1","G.2","G.3","G.4",
                      "A.1","A.2","A.3"],
        "nota": "Presupuesto base ya emitido (PV01091239). Preguntar por Cerdisa Deep Sun stock."
    },
    "blaisten": {
        "nombre": "Blaisten",
        "email": "ventasweb@blaisten.com",
        "atencion": "Ventas",
        "rubro": "Sanitarios Ferrum, Klaukol, vanitory",
        "sucursal": "Floresta (Alberdi 3928)",
        "productos": ["S.1","S.2","S.3","S.4", "V.1","V.2", "K.1"],
        "nota": "Consultar precios de Ferrum Bari completo y vanitory Schneider."
    },
    "el_amigo": {
        "nombre": "El Amigo",
        "email": "ventas@elamigo.com.ar",
        "atencion": "Ventas",
        "rubro": "Sanitarios Ferrum",
        "sucursal": "CABA",
        "productos": ["S.1","S.2","S.3","S.4"],
        "nota": "Tiene buenos precios en Ferrum Bari según relevamiento."
    },
    "bercomat": {
        "nombre": "Familia Bercomat",
        "email": "consultas@bercomat.com",
        "atencion": "Ventas",
        "rubro": "Grifería FV Mauna, adhesivos, pastinas",
        "sucursal": "Varias GBA",
        "productos": ["K.1", "P.1", "P.2", "P.3", "G.1", "G.2", "G.3"],
        "nota": "Mejor precio en Klaukol adhesivo (~$18.200). Consultar promo vigente."
    },
    "nimat": {
        "nombre": "NIMAT",
        "email": "ventas@nimat.com.ar",
        "atencion": "Ventas",
        "rubro": "Grifería FV Mauna",
        "sucursal": "CABA",
        "productos": ["G.1", "G.2", "G.3"],
        "nota": "Tiene la ducha Mauna más barata que Barugel (-$42.760 c/u)."
    },
    "merlino": {
        "nombre": "Merlino",
        "email": "ventas@merlino.com.ar",
        "atencion": "Ventas",
        "rubro": "Vanitory, muebles de baño",
        "sucursal": "CABA",
        "productos": ["V.2"],
        "nota": "Mejor precio en Vanitory Rivo 60cm ($307.703)."
    },
    "resta": {
        "nombre": "Resta",
        "email": "ventas@resta.com.ar",
        "atencion": "Ventas",
        "rubro": "Grifería, Duchamatic FV",
        "sucursal": "CABA",
        "productos": ["G.4"],
        "nota": "Precio con -5% por transferencia bancaria."
    },
    "edificor": {
        "nombre": "Edificor",
        "email": "ventas@edificor.com.ar",
        "atencion": "Ventas",
        "rubro": "Revestimientos porcelánicos",
        "sucursal": "CABA",
        "productos": ["D.6", "D.7"],
        "nota": "Pinto Vivant ~50% más barato que Barugel."
    },
    "ferrocons": {
        "nombre": "Ferrocons",
        "email": "ventas@ferrocons.com.ar",
        "atencion": "Ventas",
        "rubro": "Revestimientos",
        "sucursal": "CABA",
        "productos": ["D.2", "D.4"],
        "nota": "Buen precio en Portobello Nord Ris y Tendenza."
    },
    "distabile": {
        "nombre": "Distabile",
        "email": "ventas@distabile.com.ar",
        "atencion": "Ventas",
        "rubro": "Adhesivos Klaukol",
        "sucursal": "CABA",
        "productos": ["K.1"],
        "nota": "Precio competitivo en adhesivo ($27.712)."
    },
    "caballito_sanitarios": {
        "nombre": "Caballito Sanitarios",
        "email": "info@caballitosanitarios.com.ar",
        "atencion": "Ventas",
        "rubro": "Sanitarios Ferrum + FV",
        "sucursal": "Donato Álvarez 64 (a 5 cuadras)",
        "productos": ["S.1","S.2","S.3","S.4", "G.1","G.2","G.3"],
        "nota": "PENDIENTE — verificar email exacto. Queda cerca para ir presencial."
    },
    "blaisten": {
        "nombre": "Blaisten",
        "email": "ventasweb@blaisten.com",
        "atencion": "Ventas",
        "rubro": "Revestimientos, sanitarios, grifería, adhesivos",
        "sucursal": "Alberdi 3928, Floresta (~2km)",
        "productos": ["D.1","D.2","D.3","D.4","D.5","D.6","D.7","K.1","P.1","P.2","P.3","S.1","S.2","S.3","S.4","V.1","V.2","B.1","B.2","G.1","G.2","G.3","G.4","A.1","A.2","A.3","H4.1","H4.2","H4.3","H4.4","H4.5","H4.6","H4.7"],
        "nota": "Cadena grande. Pedir alternativas si no tienen algún producto exacto."
    },
    "bercomat": {
        "nombre": "Familia Bercomat",
        "email": "consultas@bercomat.com",
        "atencion": "Ventas",
        "rubro": "Sanitarios, grifería, adhesivos, materiales",
        "sucursal": "Varias GBA",
        "productos": ["D.1","D.2","D.3","D.4","D.5","D.6","D.7","K.1","P.1","P.2","P.3","S.1","S.2","S.3","S.4","V.1","V.2","B.1","B.2","G.1","G.2","G.3","G.4","A.1","A.2","A.3","H4.1","H4.2","H4.3","H4.4","H4.5","H4.6","H4.7"],
        "nota": "65+ años, varias sucursales."
    },
    "easy": {
        "nombre": "Easy",
        "email": "ventasecommerce@cencosud.com.ar",
        "atencion": "Ventas",
        "rubro": "Todo construcción y hogar",
        "sucursal": "Palermo / Haedo",
        "productos": ["D.1","D.2","D.3","D.4","D.5","D.6","D.7","K.1","P.1","P.2","P.3","S.1","S.2","S.3","S.4","V.1","V.2","B.1","B.2","G.1","G.2","G.3","G.4","A.1","A.2","A.3","H4.1","H4.2","H4.3","H4.4","H4.5","H4.6","H4.7"],
        "nota": "Cadena nacional. Usar ventasecommerce@cencosud.com.ar (ventas empresa)."
    },
    "resta": {
        "nombre": "Resta Sanitarios",
        "email": "info@sanitariosresta.com.ar",
        "atencion": "Ventas",
        "rubro": "Sanitarios y grifería (FV, Roca, etc.)",
        "sucursal": "CABA",
        "productos": ["S.1","S.2","S.3","S.4","V.1","V.2","B.1","B.2","G.1","G.2","G.3","G.4","A.1","A.2","A.3","H4.1","H4.2","H4.3","H4.4","H4.5","H4.6","H4.7","K.1","P.1","P.2","P.3"],
        "nota": "Especialista en sanitarios y grifería. Lleva FV, Roca. No vende revestimientos."
    },
}


# ── CATÁLOGO DE PRODUCTOS ──────────────────────────────────────────────
PRODUCTOS = {
    "D.1": {"nombre": "Portobello Art Martase", "unidad": "m²", "cant": 1.44},
    "D.2": {"nombre": "Tendenza Black Cement (porcelanato)", "unidad": "m²", "cant": 4.95},
    "D.3": {"nombre": "Eliane Forma Branco 32x60 (cerámica)", "unidad": "m²", "cant": 30.10},
    "D.4": {"nombre": "Portobello Nord Ris Ripado 30x90", "unidad": "m²", "cant": 4.02},
    "D.5": {"nombre": "Vite Liscio Light Eco Grey 20x120", "unidad": "m²", "cant": 4.32},
    "D.6": {"nombre": "Pinto Vivant Luge Suave 7x24", "unidad": "m²", "cant": 2.04},
    "D.7": {"nombre": "Pinto PBG Vivant Suage 7x24", "unidad": "m²", "cant": 3.06},
    "K.1": {"nombre": "Klaukol Adhesivo Fluido Porcelanato x25kg", "unidad": "un", "cant": 26},
    "P.1": {"nombre": "Klaukol Pastina Fluida SKG05 Blanco 5kg", "unidad": "un", "cant": 1},
    "P.2": {"nombre": "Klaukol Pastina Fluida SKG05 Gris Claro 5kg", "unidad": "un", "cant": 1},
    "P.3": {"nombre": "Klaukol Pastina Fluida AP Perla 5kg", "unidad": "un", "cant": 2},
    "S.1": {"nombre": "Inodoro Largo Ferrum Bari Blanco", "unidad": "un", "cant": 1},
    "S.2": {"nombre": "Depósito Ferrum Bari Dual de Apoyar", "unidad": "un", "cant": 1},
    "S.3": {"nombre": "Bidet Ferrum Bari 3 Agujeros", "unidad": "un", "cant": 1},
    "S.4": {"nombre": "Tapa Asiento Ferrum Bari Madera", "unidad": "un", "cant": 1},
    "V.1": {"nombre": "Schneider Vanitory City 80cm Colgante", "unidad": "un", "cant": 1},
    "V.2": {"nombre": "Schneider Vanitory Rivo 60cm Colgante", "unidad": "un", "cant": 1},
    "B.1": {"nombre": "FV Bañera Enoxado 150x70", "unidad": "un", "cant": 1},
    "B.2": {"nombre": "FV Desagüe Lineal 60cm Rejilla Ciega", "unidad": "un", "cant": 1},
    "G.1": {"nombre": "FV Lavatorio Mauna Cromo (juego grifería)", "unidad": "JG", "cant": 3},
    "G.2": {"nombre": "FV Bidet Mauna Cromo (juego grifería)", "unidad": "JG", "cant": 2},
    "G.3": {"nombre": "FV Ducha Mauna c/Transferencia Cromo", "unidad": "JG", "cant": 2},
    "G.4": {"nombre": "FV Duchamatic c/Brazo Cromo", "unidad": "un", "cant": 2},
    "A.1": {"nombre": "FV Portarrollos Cromo C/0167", "unidad": "un", "cant": 3},
    "A.2": {"nombre": "FV Toallero Barra Cromo C/0405", "unidad": "un", "cant": 3},
    "A.3": {"nombre": "FV Percha Cromo C/0166", "unidad": "un", "cant": 5},
    "H4.1": {"nombre": "Espejo Reflejar Rectang. 100x80 4mm", "unidad": "un", "cant": 1},
    "H4.2": {"nombre": "Espejo 20x30 Bordes Acabado Ilumin Oxboro 220V", "unidad": "un", "cant": 1},
    "H4.3": {"nombre": "Indoor Largo c/Salvaguarda GAP Blanco Roca", "unidad": "un", "cant": 1},
    "H4.4": {"nombre": "Alimentación Efectiva Cond. Bajo Exc. 3/6 L", "unidad": "un", "cant": 1},
    "H4.5": {"nombre": "Bidet 3 Agujeros Roca La Gama Tapas Cindo Pintado", "unidad": "un", "cant": 1},
    "H4.6": {"nombre": "Tapa Amortiguador Slim Cond. Amortiguada", "unidad": "un", "cant": 1},
    "H4.7": {"nombre": "Portobello Zesty Light Grate Mateo 45x120 Text", "unidad": "m²", "cant": 6.44},
}


# ── TEMPLATES DE EMAIL ─────────────────────────────────────────────────
def _build_product_list(items):
    lines = []
    for c in items:
        p = PRODUCTOS.get(c)
        if p:
            lines.append(f"  • {c} — {p['nombre']} ({p['cant']} {p['unidad']})")
    return "\n".join(lines)


TEMPLATES = {}

TEMPLATES["barugel"] = lambda prov, items: f"""\
Hola Marcos, ¿cómo estás?

Soy Pablo, te escribía porque la semana pasada me pasaste el presupuesto (PV01091239) y quería consultarte un par de cosas antes de avanzar con la compra.

Primero quería saber si el Cerdisa Deep Sun 60x120 (código 34616630) sigue teniendo stock en sucursal Alcorta, porque es una de las primeras cosas que quiero comprar.

Después, por los otros productos del presupuesto, quería consultarte:

{_build_product_list(items)}

¿Estos precios están vigentes? ¿Tienen disponibilidad actual?

Y aprovecho a preguntarte:
  • ¿El flete a Caballito tiene algún costo aproximado?
  • ¿Tienen descuento por pago en efectivo o transferencia?
  • ¿Suelen ofrecer financiación (cuotas) o solo contado/transferencia?

Desde ya muchas gracias, cualquier cosa avisame.

Saludos,
Pablo"""

TEMPLATES["nimat"] = lambda prov, items: f"""\
Hola, ¿cómo están?

Les escribo porque estoy haciendo una remodelación en Caballito y estoy buscando grifería FV Mauna línea Cromo. Me llamo Pablo.

Necesito cotizar:

{_build_product_list(items)}

Me interesaba saber:
  • ¿Tienen stock disponible?
  • ¿Precio por unidad y por el total?
  • ¿Hacen envío a Caballito? ¿cobra el flete?
  • ¿Tienen descuento por pago en efectivo o transferencia bancaria?
  • ¿Aceptan cuotas o financiación?

Cualquier cosa, quedo atento a su respuesta. Muchas gracias.

Saludos,
Pablo"""

TEMPLATES["bercomat"] = lambda prov, items: f"""\
Hola, ¿cómo están?

Les escribo porque estoy haciendo una obra en Caballito y me interesó su local para comprar algunos materiales. Me llamo Pablo.

Estoy necesitando:

{_build_product_list(items)}

Si me pueden pasar precios y disponibilidad se los agradezco. También me interesaría saber:
  • ¿Tienen algún descuento por pago en efectivo o transferencia?
  • ¿Hacen envíos a Caballito? ¿cuánto sale el flete?
  • ¿Aceptan cuotas o financiación?

Desde ya muchas gracias.

Saludos,
Pablo"""

TEMPLATES["default"] = lambda prov, items: f"""\
Hola, ¿cómo están?

Les escribo porque estoy haciendo una remodelación en Caballito, CABA, y estoy buscando cotizar algunos productos para el baño y la cocina. Me llamo Pablo.

Necesito precio y disponibilidad de:

{_build_product_list(items)}

Aprovecho a consultarles:
  • ¿Tienen stock de estos productos?
  • ¿Hacen envíos a Caballito? ¿cuánto sale el flete?
  • ¿Tienen descuento por pago en efectivo o transferencia bancaria?
  • ¿Aceptan cuotas o financiación?

Les agradezco desde ya la información. Quedo atento.

Saludos,
Pablo"""


TEMPLATES["general"] = lambda prov, items: f"""\
Hola, ¿cómo están?

Les escribo porque estoy haciendo una remodelación en Caballito y quería pedirles un presupuesto para una serie de productos. Soy Pablo.

{_build_product_list(items)}

Si de algunos no tienen el producto exacto, por favor mándenme opciones equivalentes que puedan ofrecer.

Me interesaría saber:
  • Precio unitario y total de cada producto
  • Stock disponible
  • Si hacen envío a Caballito y cuánto sale el flete
  • Si tienen descuento por pago en efectivo o transferencia bancaria
  • Si aceptan cuotas o financiación

Desde ya muchas gracias.

Saludos,
Pablo"""


def generar_email(proveedor_key, items=None):
    prov = PROVEEDORES[proveedor_key]
    if items is None:
        items = prov["productos"]  # todos los productos
    general_provs = ["blaisten", "bercomat", "easy", "resta"]
    template_key = "general" if proveedor_key in general_provs else proveedor_key
    template_fn = TEMPLATES.get(template_key, TEMPLATES["default"])
    body = template_fn(prov, items)
    subject = f"Consulta por {prov['rubro'].lower()} — Remodelación Caballito"
    if proveedor_key == "barugel":
        subject = f"Hola Marcos, consulta sobre presupuesto PV01091239"
    return {
        "to": prov["email"],
        "cc": "",
        "subject": subject,
        "body": body,
    }


# ── ENVÍO ──────────────────────────────────────────────────────────────
def _create_message(to, subject, body, cc=""):
    msg = MIMEMultipart("alternative")
    msg["To"] = to
    msg["Subject"] = subject
    if cc:
        msg["Cc"] = cc
    # parte texto plano
    msg.attach(MIMEText(body, "plain", "utf-8"))
    # parte HTML simple
    html = body.replace("\n", "<br>\n")
    msg.attach(MIMEText(f"<html><body><pre style='font-family:sans-serif;font-size:14px;line-height:1.6'>{body}</pre></body></html>", "html", "utf-8"))
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    return {"raw": raw}


def enviar_consulta(proveedor_key, items=None, dry_run=False):
    prov = PROVEEDORES[proveedor_key]
    email_data = generar_email(proveedor_key, items)

    print(f"─── Consulta a: {prov['nombre']} ({prov['email']}) ───")
    print(f"Asunto: {email_data['subject']}")
    print()
    print(email_data["body"])
    print()

    if dry_run:
        print("🔹 DRY RUN — No se envió")
        return {"proveedor": proveedor_key, "status": "dry_run", "asunto": email_data["subject"]}

    try:
        service = get_gmail_service()
        msg = _create_message(email_data["to"], email_data["subject"], email_data["body"], email_data.get("cc", ""))
        sent = service.users().messages().send(userId="me", body=msg).execute()
        print(f"✅ Enviado — ID: {sent['id']}")
        print(f"   a {email_data['to']}")

        state = load_state()
        state["enviados"].append({
            "proveedor": proveedor_key,
            "to": email_data["to"],
            "subject": email_data["subject"],
            "items": items or prov["productos"][:8],
            "gmail_id": sent["id"],
            "fecha": datetime.now(timezone.utc).isoformat(),
        })
        save_state(state)
        return {"proveedor": proveedor_key, "status": "sent", "gmail_id": sent["id"]}

    except Exception as e:
        print(f"❌ Error al enviar: {e}")
        return {"proveedor": proveedor_key, "status": "error", "error": str(e)}


# ── REVISAR RESPUESTAS ─────────────────────────────────────────────────
def revisar_respuestas():
    print("Revisando respuestas en la bandeja de entrada...\n")
    try:
        service = get_gmail_service()
        state = load_state()

        # Buscar emails de proveedores en inbox
        proveedor_emails = [p["email"].split("@")[1] for p in PROVEEDORES.values()]
        query = " OR ".join(f"from:{d}" for d in set(proveedor_emails))
        query = f"in:inbox ({query})"

        results = service.users().messages().list(userId="me", q=query, maxResults=50).execute()
        messages = results.get("messages", [])

        if not messages:
            print("📭 No hay respuestas nuevas de proveedores.")
            return

        print(f"📬 Se encontraron {len(messages)} mensajes de proveedores:\n")

        for msg_data in messages:
            msg = service.users().messages().get(userId="me", id=msg_data["id"], format="metadata",
                metadataHeaders=["From", "Subject", "Date"]).execute()
            headers = {h["name"]: h["value"] for h in msg["payload"]["headers"]}
            sender = headers.get("From", "")
            subject = headers.get("Subject", "")
            date = headers.get("Date", "")

            # Identificar proveedor
            prov_key = None
            for key, p in PROVEEDORES.items():
                if p["email"].split("@")[1] in sender.lower():
                    prov_key = key
                    break

            print(f"  📧 De: {sender}")
            print(f"     Asunto: {subject}")
            print(f"     Fecha: {date}")
            if prov_key:
                print(f"     → Proveedor: {PROVEEDORES[prov_key]['nombre']}")
            print()

            # Registrar en estado
            if prov_key and msg_data["id"] not in [r.get("gmail_id") for r in state["recibidos"]]:
                state["recibidos"].append({
                    "proveedor": prov_key,
                    "gmail_id": msg_data["id"],
                    "sender": sender,
                    "subject": subject,
                    "fecha": date,
                    "leido": False,
                })
                save_state(state)

        print("💡 Usá 'python3 agente-compras.py leer <gmail_id>' para ver el contenido completo.")
        print("   Después 'python3 agente-compras.py resumen' para compilar resultados.")

    except Exception as e:
        print(f"❌ Error al revisar: {e}")


def leer_email(gmail_id):
    try:
        service = get_gmail_service()
        msg = service.users().messages().get(userId="me", id=gmail_id, format="full").execute()
        headers = {h["name"]: h["value"] for h in msg["payload"]["headers"]}
        
        print(f"De: {headers.get('From', '?')}")
        print(f"Asunto: {headers.get('Subject', '?')}")
        print(f"Fecha: {headers.get('Date', '?')}")
        print(f"─" * 50)
        print()

        # Extraer texto plano
        def _get_text(part):
            if part.get("mimeType") == "text/plain":
                data = part.get("body", {}).get("data", "")
                if data:
                    return base64.urlsafe_b64decode(data).decode("utf-8", errors="replace")
            if part.get("parts"):
                for p in part["parts"]:
                    result = _get_text(p)
                    if result:
                        return result
            return ""

        body = _get_text(msg["payload"])
        if body:
            print(body)
        
        # Marcar como leído
        service.users().messages().modify(userId="me", id=gmail_id,
            body={"removeLabelIds": ["UNREAD"]}).execute()

        state = load_state()
        for r in state["recibidos"]:
            if r["gmail_id"] == gmail_id:
                r["leido"] = True
        save_state(state)

    except Exception as e:
        print(f"❌ Error al leer email: {e}")


# ── RESUMEN ────────────────────────────────────────────────────────────
def generar_resumen():
    state = load_state()
    print("=" * 60)
    print("  RESUMEN DE COTIZACIONES — Obra Hortiguera 710")
    print("=" * 60)

    if not state["enviados"]:
        print("\n📭 No se han enviado consultas aún.")
        print("   Usá: python3 agente-compras.py enviar --proveedor todos")
        return

    print(f"\n📤 Consultas enviadas: {len(state['enviados'])}")
    for e in state["enviados"]:
        prov = PROVEEDORES.get(e["proveedor"], {})
        fecha = e.get("fecha", "")[:19].replace("T", " ")
        print(f"  ✅ {prov.get('nombre', e['proveedor']):30s} → {fecha}")

    print(f"\n📬 Respuestas recibidas: {len(state['recibidos'])}")
    for r in state["recibidos"]:
        prov = PROVEEDORES.get(r["proveedor"], {})
        leido = "📖" if r.get("leido") else "📕"
        print(f"  {leido} {prov.get('nombre', r['proveedor']):30s} — {r.get('subject','')[:50]}")

    print(f"\n💡 Tips:")
    print(f"   • Revisá respuestas: python3 agente-compras.py revisar")
    print(f"   • Leé un email:      python3 agente-compras.py leer <gmail_id>")
    print(f"   • Ver estado:        python3 agente-compras.py estado")


def mostrar_estado():
    state = load_state()
    enviados = len(state["enviados"])
    recibidos = len(state["recibidos"])
    pendientes = max(0, enviados - recibidos)
    print(f"Consultas enviadas:     {enviados}")
    print(f"Respuestas recibidas:   {recibidos}")
    print(f"Pendientes de respuesta: {pendientes}")
    resultados = state.get("resultados", {})
    if resultados:
        print(f"\nResultados compilados: {len(resultados)} productos")
        for code, data in sorted(resultados.items()):
            p = PRODUCTOS.get(code, {})
            mejor = data.get("mejor_precio", "?")
            proveedor = data.get("mejor_proveedor", "?")
            print(f"  {code} {p.get('nombre','')[:40]:40s} → ${mejor} ({proveedor})")


# ── CLI ────────────────────────────────────────────────────────────────
def listar_proveedores():
    print("Proveedores disponibles:\n")
    print(f"{'Clave':20s} {'Nombre':30s} {'Rubro':35s} {'Email'}")
    print("-" * 110)
    for key, p in sorted(PROVEEDORES.items()):
        print(f"{key:20s} {p['nombre']:30s} {p['rubro']:35s} {p['email']}")
    print(f"\nProductos en catálogo: {len(PRODUCTOS)} items")
    print("Usá 'python3 agente-compras.py enviar --proveedor CLAVE' para enviar consulta")


def main():
    parser = argparse.ArgumentParser(description="Agente de Compras — Obra Hortiguera 710")
    sub = parser.add_subparsers(dest="comando")

    p_listar = sub.add_parser("listar-proveedores", help="Mostrar proveedores disponibles")

    p_enviar = sub.add_parser("enviar", help="Enviar consulta a un proveedor")
    p_enviar.add_argument("--proveedor", "-p", required=True,
        help="Clave del proveedor o 'todos'")
    p_enviar.add_argument("--items", "-i", help="Items específicos separados por coma (ej: D.1,D.2)")
    p_enviar.add_argument("--dry-run", action="store_true", help="Solo mostrar el email sin enviar")

    p_revisar = sub.add_parser("revisar", help="Revisar respuestas en inbox")

    p_leer = sub.add_parser("leer", help="Leer contenido de un email")
    p_leer.add_argument("gmail_id", help="ID del mensaje en Gmail")

    p_resumen = sub.add_parser("resumen", help="Mostrar resumen de cotizaciones")
    p_estado = sub.add_parser("estado", help="Mostrar estado actual")

    args = parser.parse_args()

    if args.comando == "listar-proveedores":
        listar_proveedores()

    elif args.comando == "enviar":
        items = args.items.split(",") if args.items else None

        if args.proveedor == "todos":
            for key in sorted(PROVEEDORES.keys()):
                print()
                enviar_consulta(key, items=None if items else PROVEEDORES[key]["productos"][:8],
                               dry_run=args.dry_run)
        else:
            if args.proveedor not in PROVEEDORES:
                print(f"❌ Proveedor '{args.proveedor}' no encontrado.")
                print("   Usá 'python3 agente-compras.py listar-proveedores' para ver los disponibles.")
                sys.exit(1)
            enviar_consulta(args.proveedor, items, dry_run=args.dry_run)

    elif args.comando == "revisar":
        revisar_respuestas()

    elif args.comando == "leer":
        leer_email(args.gmail_id)

    elif args.comando == "resumen":
        generar_resumen()

    elif args.comando == "estado":
        mostrar_estado()

    else:
        parser.print_help()


if __name__ == "__main__":
    main()
