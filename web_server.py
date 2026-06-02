"""
KENNEDY EMAIL MARKETING — Servidor Web
=======================================
Backend Flask que reemplaza la app de escritorio.
Corre en Render y permite usar la plataforma desde cualquier dispositivo.

DEPLOY EN RENDER:
    1. Crear nuevo repo GitHub con este archivo + requirements_web.txt
    2. New Web Service → conectar repo
    3. Build: pip install -r requirements_web.txt
    4. Start: python web_server.py
    5. Variables de entorno:
       - SECRET_KEY  (cualquier string largo aleatorio)
       - CLIENT_ID   6e508ef7-b92b-415c-adda-c66b500c244f
"""

import os, json, base64, threading, time, math, re
from datetime import datetime
from collections import Counter
from functools import wraps
import urllib.parse, urllib.request

from flask import (Flask, request, jsonify, session, redirect,
                   url_for, render_template_string, Response)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "kennedy-secret-2026")

# ── Contraseña de acceso ──────────────────────────────────────────────────────
APP_PASSWORD = os.environ.get("APP_PASSWORD", "juanpiolmos2005")

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect("/login")
        return f(*args, **kwargs)
    return decorated

LOGIN_HTML = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Kennedy — Acceso</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{background:#0A0F1E;display:flex;align-items:center;justify-content:center;
     min-height:100vh;font-family:'Segoe UI',sans-serif;}
.box{background:#131D35;border:1px solid #1E3050;border-radius:12px;
     padding:40px 32px;width:90%;max-width:360px;text-align:center;}
.icon{font-size:3em;margin-bottom:12px;}
h2{color:#EEF2FF;font-size:18px;margin-bottom:6px;}
p{color:#7B8DB0;font-size:13px;margin-bottom:24px;}
input{width:100%;background:#1A2540;border:1px solid #1E3050;color:#EEF2FF;
      border-radius:8px;padding:12px;font-size:14px;outline:none;margin-bottom:14px;
      text-align:center;letter-spacing:2px;}
input:focus{border-color:#4F8EF7;}
button{width:100%;background:#4F8EF7;color:#fff;border:none;border-radius:8px;
       padding:12px;font-size:14px;font-weight:700;cursor:pointer;}
button:hover{opacity:.9;}
.err{color:#FF5370;font-size:12px;margin-bottom:10px;}
</style>
</head>
<body>
<div class="box">
  <div class="icon">✉</div>
  <h2>Email Marketing</h2>
  <p>Universidad Kennedy</p>
  {% if error %}<div class="err">⚠ Clave incorrecta</div>{% endif %}
  <form method="POST" action="/login">
    <input type="password" name="password" placeholder="Ingresá la clave" autofocus>
    <button type="submit">Ingresar</button>
  </form>
</div>
</body>
</html>"""

@app.route("/login", methods=["GET","POST"])
def login():
    from flask import render_template_string as rts
    error = False
    if request.method == "POST":
        if request.form.get("password") == APP_PASSWORD:
            session["logged_in"] = True
            return redirect("/")
        error = True
    return rts(LOGIN_HTML, error=error)

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")

# ── Config Microsoft ──────────────────────────────────────────────────────────
CLIENT_ID  = os.environ.get("CLIENT_ID", "6e508ef7-b92b-415c-adda-c66b500c244f")
AUTHORITY  = "https://login.microsoftonline.com/common"
SCOPES     = ["https://graph.microsoft.com/Mail.Send",
              "https://graph.microsoft.com/User.Read",
              "offline_access"]
GRAPH_URL  = "https://graph.microsoft.com/v1.0"
REDIRECT_URI = os.environ.get("REDIRECT_URI", "")  # se setea en Render

# ── Almacenamiento ────────────────────────────────────────────────────────────
DATA_DIR      = os.environ.get("DATA_DIR", "/tmp/kennedy")
os.makedirs(DATA_DIR, exist_ok=True)

ASESORES_FILE = os.path.join(DATA_DIR, "asesores.json")
TOKENS_FILE   = os.path.join(DATA_DIR, "tokens.json")
TRACKING_FILE = os.path.join(DATA_DIR, "tracking.json")
VENTAS_FILE   = os.path.join(DATA_DIR, "ventas.json")
HISTORIAL_FILE= os.path.join(DATA_DIR, "historial.json")

# GitHub backup (opcional)
GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN", "")
GITHUB_REPO  = os.environ.get("GITHUB_REPO", "")

# ── Helpers de archivos ───────────────────────────────────────────────────────
def load_json(path, default=None):
    if os.path.exists(path):
        try:
            with open(path) as f: return json.load(f)
        except: pass
    return default if default is not None else []

def save_json(path, data):
    with open(path, "w") as f: json.dump(data, f, indent=2)

# ── MSAL helpers ──────────────────────────────────────────────────────────────
def get_msal_app():
    try:
        import msal
        return msal.ConfidentialClientApplication(
            CLIENT_ID,
            authority=AUTHORITY,
            client_credential=None,  # Public client
        )
    except:
        return None

def refresh_token_for(email):
    """Intenta renovar el access_token usando el refresh_token guardado."""
    try:
        import msal
        tokens = load_json(TOKENS_FILE, {})
        tok = tokens.get(email.lower(), {})
        if not tok.get("refresh_token"): return None
        app_msal = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
        result = app_msal.acquire_token_by_refresh_token(tok["refresh_token"], scopes=SCOPES)
        if "access_token" in result:
            tokens[email.lower()]["access_token"]  = result["access_token"]
            tokens[email.lower()]["refresh_token"] = result.get("refresh_token", tok["refresh_token"])
            save_json(TOKENS_FILE, tokens)
            return result["access_token"]
    except Exception as e:
        print(f"[token] Error renovando {email}: {e}")
    return None

def get_access_token(email):
    tokens = load_json(TOKENS_FILE, {})
    tok = tokens.get(email.lower(), {})
    if tok.get("access_token"):
        # Intentar renovar por si venció
        new_tok = refresh_token_for(email)
        return new_tok or tok["access_token"]
    return None

# ── Graph API helpers ─────────────────────────────────────────────────────────
def graph_post(token, endpoint, payload):
    data = json.dumps(payload).encode()
    req = urllib.request.Request(
        f"{GRAPH_URL}{endpoint}",
        data=data,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        },
        method="POST"
    )
    try:
        with urllib.request.urlopen(req, timeout=20) as r:
            return r.status
    except urllib.error.HTTPError as e:
        raise Exception(f"HTTP {e.code}: {e.read()[:200].decode()}")

def send_mail_graph(token, to_email, subject, html_body):
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html_body},
            "toRecipients": [{"emailAddress": {"address": to_email}}]
        },
        "saveToSentItems": True
    }
    status = graph_post(token, "/me/sendMail", payload)
    if status != 202:
        raise Exception(f"Error enviando mail: status {status}")

# ── Tracking HTML ─────────────────────────────────────────────────────────────
TRACKING_BASE_URL = os.environ.get("TRACKING_URL", "https://tracking-kennedy.onrender.com")
GIF_PIXEL = bytes([71,73,70,56,57,97,1,0,1,0,0,255,0,44,0,0,0,0,1,0,1,0,0,2,0,59])

def build_email_html(nombre, carrera, mail, base_id, body_html,
                     wa_number="5491154221134",
                     incluir_form=True, incluir_tel=True):
    tb = TRACKING_BASE_URL.rstrip("/")
    tel_numero = "08002223340"
    wa_url   = f"https://wa.me/{wa_number}?text=" + urllib.parse.quote(
                f"Hola, quiero hablar con un asesor. Carrera: {carrera}")
    form_url = ("https://forms.office.com/Pages/ResponsePage.aspx"
                "?id=Xa8CpRwOJ0u30AACXk6AaVU6WFRNGIpFmNQvKZwGC0NUN09PVkpUM1I3VDBPN1hYREdQR1BLTVRFNS4u")

    def turl(a, dest=""):
        params = {"a":a,"m":mail,"n":nombre,"c":carrera,"b":base_id}
        if dest: params["dest"] = dest
        return tb + "/t?" + urllib.parse.urlencode(params)

    pixel_url  = turl("open")
    wa_final   = turl("wa",   dest=wa_url)
    form_final = turl("form", dest=form_url)
    tel_final  = turl("tel",  dest=f"tel:{tel_numero}")

    BOTONES = (
        '<div style="margin:20px 0;">'
        '<div style="margin-bottom:12px;">'
        f'<a href="{wa_final}" style="background:#25D366;color:#fff;padding:12px 20px;'
        'text-decoration:none;border-radius:5px;font-weight:bold;display:inline-block;">'
        'Contactar Asesor por WhatsApp</a></div>'
        + (f'<div style="margin-bottom:12px;"><a href="{tel_final}" style="background:#D10135;'
           'color:#fff;padding:12px 20px;text-decoration:none;border-radius:5px;'
           'font-weight:bold;display:inline-block;">📞 Llamar al 0800-222-3340</a></div>'
           if incluir_tel else '')
        + (f'<div><a href="{form_final}" style="background:#333;color:#fff;padding:12px 20px;'
           'text-decoration:none;border-radius:5px;font-weight:bold;display:inline-block;">'
           'Completar Formulario de Inscripción</a></div>'
           if incluir_form else '')
        + '</div>'
    )
    FIRMA = ('<p>Atentamente,<br><strong>Admisiones | Universidad Kennedy</strong><br>'
             '<small style="color:#888;">Este email fue enviado a ' + mail + '</small></p>')
    PIXEL = f'<img src="{pixel_url}" width="1" height="1" style="display:none;" alt="">'

    return (
        '<div style="font-family:Arial,sans-serif;color:#333;line-height:1.6;max-width:600px;">'
        + body_html + BOTONES + FIRMA + PIXEL + '</div>'
    )

# ── Cola de envíos ────────────────────────────────────────────────────────────
_job_status = {}   # job_id -> {...}
_jobs_lock  = threading.Lock()

def run_send_job(job_id, contacts, asesores, subject_tpl, body_html_tpl,
                 wa_number, incluir_form, incluir_tel, delay, base_id, base_name):
    with _jobs_lock:
        _job_status[job_id] = {
            "status": "running", "total": len(contacts),
            "sent": 0, "errors": 0, "log": [], "start": datetime.now().isoformat()
        }

    tokens = load_json(TOKENS_FILE, {})
    historial = load_json(HISTORIAL_FILE, [])
    activos = [a for a in asesores if a.get("active", True)]
    if not activos:
        with _jobs_lock:
            _job_status[job_id]["status"] = "error"
            _job_status[job_id]["log"].append("Sin asesores activos")
        return

    def log(msg, tipo="info"):
        with _jobs_lock:
            _job_status[job_id]["log"].append({"ts": datetime.now().strftime("%H:%M:%S"),
                                               "msg": msg, "tipo": tipo})

    for i, row in enumerate(contacts):
        if _job_status[job_id].get("stop"):
            log("Detenido por el usuario", "warn"); break

        # Distribuir por round-robin
        adv = activos[i % len(activos)]
        adv_email = adv["email"]

        token = get_access_token(adv_email)
        if not token:
            log(f"Sin token para {adv_email} — omitiendo contacto {row.get('mail','')}", "err")
            with _jobs_lock: _job_status[job_id]["errors"] += 1
            continue

        nombre  = str(row.get("nombre","")).strip()
        mail    = str(row.get("mail","")).strip()
        carrera = str(row.get("carrera","")).strip()

        def fill(t):
            return (t.replace("{Nombre}", nombre)
                     .replace("{Asesor}", adv.get("name",""))
                     .replace("{Carrera}", carrera)
                     .replace("{Apellido}", str(row.get("apellido","")))
                     .replace("{Provincia}", str(row.get("provincia","")))
                     .replace("{Legajo}", str(row.get("legajo",""))))

        subject  = fill(subject_tpl)
        body_html = fill(build_email_html(nombre, carrera, mail, base_id,
                                          fill(body_html_tpl),
                                          wa_number=wa_number,
                                          incluir_form=incluir_form,
                                          incluir_tel=incluir_tel))
        try:
            send_mail_graph(token, mail, subject, body_html)
            with _jobs_lock: _job_status[job_id]["sent"] += 1
            log(f"✅ {mail} → {adv.get('name','')}", "ok")
            historial.append({
                "ts": datetime.now().isoformat(), "asesor": adv.get("name",""),
                "to": mail, "nombre": nombre, "carrera": carrera,
                "status": "enviado", "base_id": base_id, "base_name": base_name
            })
        except Exception as e:
            with _jobs_lock: _job_status[job_id]["errors"] += 1
            log(f"❌ {mail}: {str(e)[:80]}", "err")
            historial.append({
                "ts": datetime.now().isoformat(), "asesor": adv.get("name",""),
                "to": mail, "nombre": nombre, "carrera": carrera,
                "status": "error", "base_id": base_id, "base_name": base_name
            })

        save_json(HISTORIAL_FILE, historial)
        if i < len(contacts) - 1:
            time.sleep(delay)

    with _jobs_lock:
        _job_status[job_id]["status"] = "done"
        log(f"Finalizado — Enviados: {_job_status[job_id]['sent']} | Errores: {_job_status[job_id]['errors']}", "ok")

# ══════════════════════════════════════════════════════════════════════════════
#  API ENDPOINTS
# ══════════════════════════════════════════════════════════════════════════════

# ── Asesores ──────────────────────────────────────────────────────────────────
@app.route("/api/asesores", methods=["GET"])
def api_asesores():
    return jsonify(load_json(ASESORES_FILE, []))

@app.route("/api/asesores", methods=["POST"])
def api_asesores_save():
    save_json(ASESORES_FILE, request.json)
    return jsonify({"ok": True})

# ── Auth Microsoft ────────────────────────────────────────────────────────────
@app.route("/api/auth/url")
def api_auth_url():
    """Devuelve la URL de login de Microsoft para un asesor."""
    email = request.args.get("email","")
    redirect_uri = request.host_url.rstrip("/") + "/api/auth/callback"
    state = base64.b64encode(email.encode()).decode()
    params = {
        "client_id":     CLIENT_ID,
        "response_type": "code",
        "redirect_uri":  redirect_uri,
        "scope":         " ".join(SCOPES),
        "state":         state,
        "prompt":        "select_account",
        "login_hint":    email,
    }
    url = f"{AUTHORITY}/oauth2/v2.0/authorize?" + urllib.parse.urlencode(params)
    return jsonify({"url": url})

@app.route("/api/auth/callback")
def api_auth_callback():
    """Recibe el código de Microsoft y guarda el token."""
    code  = request.args.get("code","")
    state = request.args.get("state","")
    try:
        email = base64.b64decode(state).decode()
    except:
        email = ""

    redirect_uri = request.host_url.rstrip("/") + "/api/auth/callback"
    data = urllib.parse.urlencode({
        "client_id":    CLIENT_ID,
        "grant_type":   "authorization_code",
        "code":         code,
        "redirect_uri": redirect_uri,
        "scope":        " ".join(SCOPES),
    }).encode()

    req = urllib.request.Request(
        f"{AUTHORITY}/oauth2/v2.0/token",
        data=data,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        method="POST"
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as r:
            result = json.loads(r.read())
    except Exception as e:
        return f"<h2>Error obteniendo token: {e}</h2>", 400

    if "access_token" in result:
        tokens = load_json(TOKENS_FILE, {})
        tokens[email.lower()] = {
            "access_token":  result["access_token"],
            "refresh_token": result.get("refresh_token", ""),
            "ts": datetime.now().isoformat()
        }
        save_json(TOKENS_FILE, tokens)
        return """<html><body style="font-family:sans-serif;background:#0A0F1E;color:#fff;
                  display:flex;align-items:center;justify-content:center;height:100vh;margin:0;">
                  <div style="text-align:center;">
                  <div style="font-size:3em;">✅</div>
                  <h2 style="color:#00C896;">Login exitoso</h2>
                  <p style="color:#7B8DB0;">Podés cerrar esta pestaña y volver a la app.</p>
                  </div></body></html>"""
    return f"<h2>Error: {result.get('error_description','desconocido')}</h2>", 400

@app.route("/api/auth/status")
def api_auth_status():
    tokens = load_json(TOKENS_FILE, {})
    return jsonify({email: bool(tok.get("access_token")) for email, tok in tokens.items()})

# ── Envío ─────────────────────────────────────────────────────────────────────
@app.route("/api/send", methods=["POST"])
def api_send():
    data       = request.json
    contacts   = data.get("contacts", [])
    asesores   = load_json(ASESORES_FILE, [])
    subject    = data.get("subject", "")
    body_html  = data.get("body_html", "")
    wa_number  = data.get("wa_number", "5491154221134")
    incluir_form = data.get("incluir_form", True)
    incluir_tel  = data.get("incluir_tel", True)
    delay        = float(data.get("delay", 2))
    base_id      = data.get("base_id", "")
    base_name    = data.get("base_name", "")

    if not contacts:
        return jsonify({"error": "Sin contactos"}), 400

    job_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    threading.Thread(
        target=run_send_job,
        args=(job_id, contacts, asesores, subject, body_html,
              wa_number, incluir_form, incluir_tel, delay, base_id, base_name),
        daemon=True
    ).start()
    return jsonify({"job_id": job_id})

@app.route("/api/send/status/<job_id>")
def api_send_status(job_id):
    with _jobs_lock:
        return jsonify(_job_status.get(job_id, {"status": "not_found"}))

@app.route("/api/send/stop/<job_id>", methods=["POST"])
def api_send_stop(job_id):
    with _jobs_lock:
        if job_id in _job_status:
            _job_status[job_id]["stop"] = True
    return jsonify({"ok": True})

# ── Historial ─────────────────────────────────────────────────────────────────
@app.route("/api/historial")
def api_historial():
    return jsonify(load_json(HISTORIAL_FILE, []))

# ── Tracking ─────────────────────────────────────────────────────────────────
@app.route("/api/tracking_data")
def api_tracking_data():
    try:
        req = urllib.request.Request(
            TRACKING_BASE_URL.rstrip("/") + "/api/tracking",
            headers={"User-Agent": "kennedy-web"})
        with urllib.request.urlopen(req, timeout=10) as r:
            return jsonify(json.loads(r.read()))
    except:
        return jsonify([])

# ── Ventas ────────────────────────────────────────────────────────────────────
@app.route("/api/ventas", methods=["GET"])
def api_ventas():
    return jsonify(load_json(VENTAS_FILE, []))

@app.route("/api/ventas", methods=["POST"])
def api_ventas_save():
    save_json(VENTAS_FILE, request.json)
    return jsonify({"ok": True})

# ── Health ─────────────────────────────────────────────────────────────────────
@app.route("/health")
def health():
    return jsonify({"status": "ok", "ts": datetime.now().isoformat()})

# ══════════════════════════════════════════════════════════════════════════════
#  FRONTEND — Single Page App
# ══════════════════════════════════════════════════════════════════════════════
@app.route("/")
@app.route("/<path:path>")
@login_required
def frontend(path=""):
    return render_template_string(HTML_APP)

HTML_APP = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1">
<title>Kennedy Email Marketing</title>
<style>
:root {
  --bg:    #0A0F1E; --surf:  #0F1629; --card:  #131D35; --card2: #1A2540;
  --bord:  #1E3050; --blue:  #4F8EF7; --green: #00C896; --orange:#FFB547;
  --red:   #FF5370; --purple:#9B72F7; --teal:  #00D4CC;
  --text:  #EEF2FF; --muted: #7B8DB0; --dim:   #2D3F60;
}
* { box-sizing:border-box; margin:0; padding:0; }
body { background:var(--bg); color:var(--text); font-family:'Segoe UI',sans-serif;
       min-height:100vh; }

/* Header */
.header { background:var(--surf); border-bottom:2px solid var(--blue);
          padding:0 20px; display:flex; align-items:center; height:56px;
          position:sticky; top:0; z-index:100; }
.header-logo { font-size:22px; margin-right:10px; }
.header-title { font-size:14px; font-weight:700; color:var(--text); }
.header-sub { font-size:10px; color:var(--blue); }

/* Nav tabs */
.nav { display:flex; background:var(--surf); border-bottom:1px solid var(--bord);
       overflow-x:auto; -webkit-overflow-scrolling:touch; }
.nav-tab { padding:12px 18px; font-size:12px; font-weight:600; color:var(--muted);
           cursor:pointer; border-bottom:2px solid transparent; white-space:nowrap;
           transition:all .2s; }
.nav-tab.active { color:var(--blue); border-bottom-color:var(--blue); }
.nav-tab:hover { color:var(--text); }

/* Pages */
.page { display:none; padding:16px; max-width:1200px; margin:0 auto; }
.page.active { display:block; }

/* Cards */
.card { background:var(--card); border:1px solid var(--bord); border-radius:8px;
        margin-bottom:12px; overflow:hidden; }
.card-header { background:var(--card2); padding:10px 14px; font-size:12px;
               font-weight:700; color:var(--muted); display:flex; align-items:center;
               gap:8px; border-left:3px solid var(--blue); }
.card-body { padding:14px; }

/* KPI grid */
.kpi-grid { display:grid; gap:8px; margin-bottom:12px; }
.kpi-grid-2 { grid-template-columns:repeat(2,1fr); }
.kpi-grid-4 { grid-template-columns:repeat(4,1fr); }
@media(max-width:600px){ .kpi-grid-4 { grid-template-columns:repeat(2,1fr); } }
.kpi { background:var(--card2); border-radius:8px; overflow:hidden; text-align:center; }
.kpi-bar { height:3px; }
.kpi-label { font-size:10px; color:var(--muted); padding:8px 4px 2px; }
.kpi-val { font-size:22px; font-weight:700; padding:2px 4px 4px; }
.kpi-sub { font-size:10px; color:var(--dim); padding-bottom:8px; }

/* Buttons */
.btn { border:none; border-radius:6px; padding:8px 16px; font-size:12px;
       font-weight:700; cursor:pointer; transition:opacity .15s; }
.btn:hover { opacity:.85; }
.btn:disabled { opacity:.4; cursor:not-allowed; }
.btn-blue   { background:var(--blue);   color:#fff; }
.btn-green  { background:var(--green);  color:#000; }
.btn-orange { background:var(--orange); color:#000; }
.btn-red    { background:var(--red);    color:#fff; }
.btn-ghost  { background:var(--card2);  color:var(--muted); }
.btn-sm     { padding:5px 10px; font-size:11px; }
.btn-row    { display:flex; gap:8px; flex-wrap:wrap; margin-top:10px; }

/* Forms */
label { font-size:11px; color:var(--muted); display:block; margin-bottom:4px; margin-top:10px; }
input, select, textarea {
  width:100%; background:var(--card2); border:1px solid var(--bord);
  color:var(--text); border-radius:6px; padding:8px 10px; font-size:13px;
  outline:none; transition:border .15s; font-family:inherit; }
input:focus, select:focus, textarea:focus { border-color:var(--blue); }
textarea { resize:vertical; min-height:80px; }

/* Table */
.table-wrap { overflow-x:auto; border-radius:6px; }
table { width:100%; border-collapse:collapse; font-size:12px; }
th { background:var(--card2); color:var(--muted); font-weight:600;
     padding:8px 10px; text-align:left; white-space:nowrap; }
td { padding:8px 10px; border-bottom:1px solid var(--bord); }
tr:hover td { background:var(--card2); }
.badge { display:inline-block; padding:2px 8px; border-radius:12px;
         font-size:10px; font-weight:700; }
.badge-green  { background:#00C89622; color:var(--green); }
.badge-blue   { background:#4F8EF722; color:var(--blue); }
.badge-orange { background:#FFB54722; color:var(--orange); }
.badge-red    { background:#FF537022; color:var(--red); }

/* Log */
.log-box { background:var(--surf); border-radius:6px; padding:10px;
           height:220px; overflow-y:auto; font-family:'Courier New',monospace;
           font-size:11px; border:1px solid var(--bord); }
.log-ok   { color:var(--green); }
.log-err  { color:var(--red); }
.log-info { color:var(--blue); }
.log-warn { color:var(--orange); }

/* Progress */
progress { width:100%; height:6px; border-radius:3px; margin:8px 0; }
progress::-webkit-progress-bar { background:var(--card2); border-radius:3px; }
progress::-webkit-progress-value { background:var(--green); border-radius:3px; }

/* Toggle */
.toggle-row { display:flex; align-items:center; gap:10px; padding:8px 0; }
.toggle { position:relative; width:42px; height:22px; }
.toggle input { opacity:0; width:0; height:0; }
.toggle-slider { position:absolute; inset:0; background:var(--dim);
                 border-radius:22px; cursor:pointer; transition:.3s; }
.toggle input:checked + .toggle-slider { background:var(--green); }
.toggle-slider:before { content:""; position:absolute; width:16px; height:16px;
  left:3px; bottom:3px; background:#fff; border-radius:50%; transition:.3s; }
.toggle input:checked + .toggle-slider:before { transform:translateX(20px); }

/* File drop */
.drop-zone { border:2px dashed var(--bord2); border-radius:8px; padding:30px;
             text-align:center; cursor:pointer; transition:all .2s; color:var(--muted); }
.drop-zone:hover, .drop-zone.drag { border-color:var(--blue); color:var(--text); }
.drop-zone input { display:none; }

/* Status dot */
.dot { display:inline-block; width:8px; height:8px; border-radius:50%; margin-right:4px; }
.dot-green { background:var(--green); }
.dot-red   { background:var(--red); }
.dot-orange { background:var(--orange); }

/* Search result */
.result-card { background:var(--card2); border-left:3px solid var(--blue);
               border-radius:6px; padding:12px; margin:10px 0; }

/* Spinner */
.spinner { display:inline-block; width:16px; height:16px; border:2px solid var(--dim);
           border-top-color:var(--blue); border-radius:50%; animation:spin .6s linear infinite; }
@keyframes spin { to { transform:rotate(360deg); } }

/* Mobile adjustments */
@media(max-width:600px) {
  .page { padding:10px; }
  .card-body { padding:10px; }
  .hide-mobile { display:none; }
}
</style>
</head>
<body>

<div class="header">
  <div class="header-logo">✉</div>
  <div>
    <div class="header-title">Email Marketing</div>
    <div class="header-sub">Universidad Kennedy</div>
  </div>
  <div style="margin-left:auto;font-size:11px;color:var(--muted)" id="headerStatus"></div>
</div>

<div class="nav">
  <div class="nav-tab active" onclick="showPage('asesores')">⚙ Asesores</div>
  <div class="nav-tab" onclick="showPage('envio')">📤 Envío</div>
  <div class="nav-tab" onclick="showPage('historial')">📋 Historial</div>
  <div class="nav-tab" onclick="showPage('tracking')">🎯 Tracking</div>
  <div class="nav-tab" onclick="showPage('ventas')">💼 Ventas</div>
</div>

<!-- ══════════ ASESORES ══════════ -->
<div class="page active" id="page-asesores">
  <div class="card">
    <div class="card-header" style="border-left-color:var(--green)">⚙ Asesores de envío</div>
    <div class="card-body">
      <div id="asesores-list"></div>
      <div class="btn-row">
        <button class="btn btn-green" onclick="dlgAsesor()">＋ Agregar asesor</button>
      </div>
    </div>
  </div>
</div>

<!-- ══════════ ENVÍO ══════════ -->
<div class="page" id="page-envio">
  <div class="card">
    <div class="card-header" style="border-left-color:var(--orange)">📁 Base de contactos</div>
    <div class="card-body">
      <div class="drop-zone" onclick="document.getElementById('fileInput').click()"
           id="dropZone">
        <input type="file" id="fileInput" accept=".xlsx,.xls,.csv" onchange="loadFile(this)">
        <div style="font-size:2em;margin-bottom:8px;">📂</div>
        <div>Tocá para subir base (Excel / CSV)</div>
        <div id="fileInfo" style="color:var(--blue);margin-top:6px;font-size:12px;"></div>
      </div>
      <div id="preview" style="margin-top:10px;font-size:12px;color:var(--muted);"></div>
    </div>
  </div>

  <div class="card">
    <div class="card-header" style="border-left-color:var(--blue)">✏ Contenido del email</div>
    <div class="card-body">
      <label>Asunto:</label>
      <input type="text" id="subject" placeholder="Ej: Información sobre {Carrera} — Universidad Kennedy">
      <label>Cuerpo del email:</label>
      <div style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:6px;">
        <button class="btn btn-ghost btn-sm" onclick="insertVar('{Nombre}')">Nombre</button>
        <button class="btn btn-ghost btn-sm" onclick="insertVar('{Carrera}')">Carrera</button>
        <button class="btn btn-ghost btn-sm" onclick="insertVar('{Asesor}')">Asesor</button>
        <button class="btn btn-ghost btn-sm" onclick="insertVar('{Provincia}')">Provincia</button>
      </div>
      <textarea id="bodyText" rows="8"
        placeholder="Hola {Nombre},&#10;&#10;Te contactamos desde la Universidad Kennedy para contarte sobre la carrera de {Carrera}..."></textarea>
      <small style="color:var(--dim)">Las variables se reemplazan automáticamente al enviar.</small>

      <div style="margin-top:12px;">
        <div class="toggle-row">
          <label class="toggle"><input type="checkbox" id="incluirForm" checked>
          <div class="toggle-slider"></div></label>
          <span style="font-size:12px;">📝 Incluir botón de Formulario</span>
        </div>
        <div class="toggle-row">
          <label class="toggle"><input type="checkbox" id="incluirTel" checked>
          <div class="toggle-slider"></div></label>
          <span style="font-size:12px;">📞 Incluir botón 0800-222-3340</span>
        </div>
      </div>

      <label>WhatsApp (número sin +):</label>
      <input type="text" id="waNumber" value="5491154221134" style="max-width:200px;">
      <label>Demora entre envíos (segundos):</label>
      <input type="number" id="delay" value="2" min="1" max="30" style="max-width:100px;">
    </div>
  </div>

  <div class="card">
    <div class="card-header" style="border-left-color:var(--teal)">🚀 Control de envío</div>
    <div class="card-body">
      <div class="kpi-grid kpi-grid-4" style="margin-bottom:10px;">
        <div class="kpi"><div class="kpi-bar" style="background:var(--text)"></div>
          <div class="kpi-label">Total</div><div class="kpi-val" id="kpiTotal">0</div></div>
        <div class="kpi"><div class="kpi-bar" style="background:var(--green)"></div>
          <div class="kpi-label">Enviados</div><div class="kpi-val" style="color:var(--green)" id="kpiSent">0</div></div>
        <div class="kpi"><div class="kpi-bar" style="background:var(--red)"></div>
          <div class="kpi-label">Errores</div><div class="kpi-val" style="color:var(--red)" id="kpiErr">0</div></div>
        <div class="kpi"><div class="kpi-bar" style="background:var(--blue)"></div>
          <div class="kpi-label">Progreso</div><div class="kpi-val" style="color:var(--blue)" id="kpiPct">0%</div></div>
      </div>
      <progress id="progress" value="0" max="100"></progress>
      <div class="btn-row">
        <button class="btn btn-green" id="btnStart" onclick="startSend()">▶ INICIAR ENVÍO</button>
        <button class="btn btn-red" id="btnStop" onclick="stopSend()" disabled>■ DETENER</button>
      </div>
      <div class="log-box" id="logBox" style="margin-top:10px;"></div>
    </div>
  </div>
</div>

<!-- ══════════ HISTORIAL ══════════ -->
<div class="page" id="page-historial">
  <div class="card">
    <div class="card-header" style="border-left-color:var(--teal)">📋 Historial de envíos</div>
    <div class="card-body">
      <div class="btn-row" style="margin-bottom:10px;">
        <button class="btn btn-ghost btn-sm" onclick="loadHistorial()">🔄 Actualizar</button>
        <button class="btn btn-ghost btn-sm" onclick="exportHistorial()">⬇ Exportar CSV</button>
      </div>
      <div class="table-wrap">
        <table id="histTable">
          <thead><tr>
            <th>Fecha</th><th>Asesor</th><th>Email</th>
            <th>Nombre</th><th>Carrera</th><th>Estado</th>
          </tr></thead>
          <tbody id="histBody"></tbody>
        </table>
      </div>
    </div>
  </div>
</div>

<!-- ══════════ TRACKING ══════════ -->
<div class="page" id="page-tracking">
  <div class="kpi-grid kpi-grid-4" id="trackingKpis"></div>
  <div class="card">
    <div class="card-header" style="border-left-color:var(--blue)">🎯 Eventos recientes</div>
    <div class="card-body">
      <div class="btn-row" style="margin-bottom:10px;">
        <button class="btn btn-ghost btn-sm" onclick="loadTracking()">🔄 Actualizar</button>
      </div>
      <div class="table-wrap">
        <table id="trackTable">
          <thead><tr>
            <th>Fecha/Hora</th><th>Acción</th><th>Nombre</th>
            <th>Email</th><th class="hide-mobile">Carrera</th><th class="hide-mobile">Base</th>
          </tr></thead>
          <tbody id="trackBody"></tbody>
        </table>
      </div>
    </div>
  </div>
</div>

<!-- ══════════ VENTAS ══════════ -->
<div class="page" id="page-ventas">
  <div class="card">
    <div class="card-header" style="border-left-color:var(--blue)">🔍 Buscar contacto</div>
    <div class="card-body">
      <div style="display:flex;gap:8px;">
        <input type="text" id="ventaSearch" placeholder="Email o teléfono..."
               onkeydown="if(event.key==='Enter')buscarContacto()">
        <button class="btn btn-blue btn-sm" onclick="buscarContacto()">Buscar</button>
      </div>
      <div id="ventaResultado" style="margin-top:10px;"></div>
    </div>
  </div>

  <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;">
    <div class="card" style="margin-bottom:0">
      <div class="card-header" style="border-left-color:var(--green)">📋 Registros</div>
      <div class="card-body" style="padding:0;">
        <div class="table-wrap" style="max-height:400px;overflow-y:auto;">
          <table id="ventasTable">
            <thead><tr>
              <th>Fecha</th><th>Nombre</th><th>Estado</th><th>Base</th>
            </tr></thead>
            <tbody id="ventasBody"></tbody>
          </table>
        </div>
      </div>
    </div>

    <div class="card" style="margin-bottom:0">
      <div class="card-header" style="border-left-color:var(--orange)">✏ Registrar / Editar</div>
      <div class="card-body">
        <input type="hidden" id="vId">
        <label>Nombre:</label><input type="text" id="vNombre">
        <label>Email:</label><input type="text" id="vEmail">
        <label>Teléfono:</label><input type="text" id="vTel">
        <label>Base:</label><input type="text" id="vBase">
        <label>Carrera:</label><input type="text" id="vCarrera">
        <label>Asesor:</label><input type="text" id="vAsesor">
        <label>Legajo:</label><input type="text" id="vLegajo">
        <label>Estado:</label>
        <select id="vEstado">
          <option>En seguimiento</option>
          <option>Venta</option>
          <option>Tiene legajo</option>
          <option>Perdido</option>
        </select>
        <label>Notas:</label><textarea id="vNotas" rows="3"></textarea>
        <div class="btn-row">
          <button class="btn btn-green btn-sm" onclick="guardarVenta()">💾 Guardar</button>
          <button class="btn btn-ghost btn-sm" onclick="limpiarVenta()">Limpiar</button>
          <button class="btn btn-red btn-sm" onclick="eliminarVenta()">🗑</button>
        </div>
      </div>
    </div>
  </div>
</div>

<!-- ══════════ MODAL ASESOR ══════════ -->
<div id="modal" style="display:none;position:fixed;inset:0;background:#0008;z-index:200;
     display:none;align-items:center;justify-content:center;">
  <div style="background:var(--card);border-radius:12px;padding:24px;width:90%;max-width:400px;
              border:1px solid var(--bord2);">
    <h3 style="margin-bottom:16px;color:var(--text)">Agregar/Editar Asesor</h3>
    <input type="hidden" id="aId">
    <label>Nombre:</label><input type="text" id="aNombre" placeholder="Juan García">
    <label>Email Microsoft:</label><input type="text" id="aEmail" placeholder="juan@kennedy.edu.ar">
    <div id="authStatus" style="margin:10px 0;font-size:12px;"></div>
    <div class="btn-row">
      <button class="btn btn-green btn-sm" onclick="authAsesor()">🔑 Autenticar cuenta</button>
    </div>
    <div class="btn-row" style="margin-top:16px;">
      <button class="btn btn-blue" onclick="saveAsesor()">Guardar</button>
      <button class="btn btn-ghost" onclick="closeModal()">Cancelar</button>
    </div>
  </div>
</div>

<script>
// ── Estado global ─────────────────────────────────────────────────────────────
let state = {
  asesores: [], contacts: [], baseId: "", baseName: "",
  currentJob: null, tracking: [], ventas: [], historial: [],
  editingVenta: null
};

// ── Navegación ────────────────────────────────────────────────────────────────
function showPage(name) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-tab').forEach(t => t.classList.remove('active'));
  document.getElementById('page-' + name).classList.add('active');
  event.currentTarget.classList.add('active');
  if (name === 'tracking') loadTracking();
  if (name === 'historial') loadHistorial();
  if (name === 'ventas') { loadVentas(); }
  if (name === 'asesores') loadAsesores();
}

// ── API helper ─────────────────────────────────────────────────────────────────
async function api(path, opts={}) {
  const r = await fetch(path, {
    headers: {'Content-Type':'application/json'},
    ...opts,
    body: opts.body ? JSON.stringify(opts.body) : undefined
  });
  return r.json();
}

// ── Asesores ─────────────────────────────────────────────────────────────────
async function loadAsesores() {
  state.asesores = await api('/api/asesores');
  const authStatus = await api('/api/auth/status');
  const el = document.getElementById('asesores-list');
  if (!state.asesores.length) {
    el.innerHTML = '<p style="color:var(--muted);font-size:13px;">No hay asesores configurados.</p>';
    return;
  }
  el.innerHTML = state.asesores.map((a,i) => `
    <div style="display:flex;align-items:center;gap:10px;padding:10px;
         background:var(--card2);border-radius:8px;margin-bottom:6px;">
      <span class="dot ${authStatus[a.email?.toLowerCase()] ? 'dot-green' : 'dot-red'}"></span>
      <div style="flex:1;">
        <div style="font-size:13px;font-weight:600;">${a.name||''}</div>
        <div style="font-size:11px;color:var(--muted);">${a.email||''}</div>
      </div>
      <div style="display:flex;gap:6px;">
        <button class="btn btn-ghost btn-sm" onclick="dlgAsesor(${i})">✏</button>
        <button class="btn btn-red btn-sm" onclick="delAsesor(${i})">🗑</button>
        <button class="btn btn-${a.active!==false?'green':'ghost'} btn-sm"
                onclick="toggleAsesor(${i})">${a.active!==false?'✓ Activo':'Pausado'}</button>
      </div>
    </div>`).join('');
}

function dlgAsesor(idx=null) {
  document.getElementById('modal').style.display = 'flex';
  document.getElementById('aId').value = idx ?? '';
  const a = idx !== null ? state.asesores[idx] : {};
  document.getElementById('aNombre').value = a.name || '';
  document.getElementById('aEmail').value  = a.email || '';
  document.getElementById('authStatus').innerHTML = '';
}

function closeModal() { document.getElementById('modal').style.display = 'none'; }

async function saveAsesor() {
  const idx   = document.getElementById('aId').value;
  const name  = document.getElementById('aNombre').value.trim();
  const email = document.getElementById('aEmail').value.trim().toLowerCase();
  if (!name || !email) { alert('Completá nombre y email'); return; }
  const a = { name, email, active: true };
  if (idx !== '') state.asesores[parseInt(idx)] = a;
  else state.asesores.push(a);
  await api('/api/asesores', { method:'POST', body: state.asesores });
  closeModal(); loadAsesores();
}

async function delAsesor(i) {
  if (!confirm('¿Eliminar asesor?')) return;
  state.asesores.splice(i, 1);
  await api('/api/asesores', { method:'POST', body: state.asesores });
  loadAsesores();
}

async function toggleAsesor(i) {
  state.asesores[i].active = !(state.asesores[i].active !== false);
  await api('/api/asesores', { method:'POST', body: state.asesores });
  loadAsesores();
}

async function authAsesor() {
  const email = document.getElementById('aEmail').value.trim();
  if (!email) { alert('Ingresá el email primero'); return; }
  const {url} = await api(`/api/auth/url?email=${encodeURIComponent(email)}`);
  window.open(url, '_blank', 'width=500,height=650');
  document.getElementById('authStatus').innerHTML =
    '<span style="color:var(--orange)">⏳ Esperando autenticación... Refrescá asesores cuando termines.</span>';
}

// ── Carga de archivo ──────────────────────────────────────────────────────────
async function loadFile(input) {
  const file = input.files[0]; if (!file) return;
  document.getElementById('fileInfo').textContent = '⏳ Procesando...';

  const formData = new FormData();
  formData.append('file', file);

  // Leer CSV/XLSX en el cliente usando FileReader + parsing básico
  const ext = file.name.split('.').pop().toLowerCase();
  if (ext === 'csv') {
    const text = await file.text();
    parseCSV(text, file.name);
  } else {
    // Para XLSX mandamos al servidor
    const fd = new FormData(); fd.append('file', file);
    const r = await fetch('/api/upload', {method:'POST', body: fd});
    const data = await r.json();
    if (data.contacts) {
      state.contacts  = data.contacts;
      state.baseId    = data.base_id;
      state.baseName  = file.name;
      showPreview();
    }
  }
}

function parseCSV(text, fname) {
  const lines = text.trim().split('\n');
  const headers = lines[0].split(/[,;]/).map(h => h.trim().toLowerCase()
    .replace(/['"]/g,'').normalize('NFD').replace(/[\u0300-\u036f]/g,''));
  const findCol = (...opts) => {
    for (const o of opts) { const i = headers.indexOf(o); if(i>=0) return i; }
    return -1;
  };
  const iMail = findCol('mail','email','correo','e-mail');
  const iNom  = findCol('nombre','name','first_name','firstname');
  const iApe  = findCol('apellido','apellidos','last_name','lastname');
  const iCar  = findCol('carrera','career');
  const iProv = findCol('provincia','province','region');
  const iLeg  = findCol('legajo','cli_idoriginal','id');

  state.contacts = lines.slice(1).map(l => {
    const cols = l.split(/[,;]/).map(c => c.trim().replace(/^["']|["']$/g,''));
    return {
      mail:     iMail>=0 ? cols[iMail] : '',
      nombre:   iNom>=0  ? cols[iNom]  : '',
      apellido: iApe>=0  ? cols[iApe]  : '',
      carrera:  iCar>=0  ? cols[iCar]  : '',
      provincia:iProv>=0 ? cols[iProv] : '',
      legajo:   iLeg>=0  ? cols[iLeg]  : '',
    };
  }).filter(c => c.mail && c.mail.includes('@'));

  state.baseId   = fname.replace(/\.[^.]+$/,'');
  state.baseName = fname;
  showPreview();
}

function showPreview() {
  document.getElementById('fileInfo').textContent =
    `✅ ${state.baseName} — ${state.contacts.length} contactos cargados`;
  document.getElementById('preview').textContent =
    state.contacts.slice(0,3).map(c => `${c.nombre} ${c.apellido} <${c.mail}> — ${c.carrera}`).join('\n')
    + (state.contacts.length > 3 ? `\n... y ${state.contacts.length-3} más` : '');
  document.getElementById('kpiTotal').textContent = state.contacts.length;
}

function insertVar(v) {
  const ta = document.getElementById('bodyText');
  const s  = ta.selectionStart, e = ta.selectionEnd;
  ta.value = ta.value.slice(0,s) + v + ta.value.slice(e);
  ta.selectionStart = ta.selectionEnd = s + v.length;
  ta.focus();
}

// ── Envío ─────────────────────────────────────────────────────────────────────
function textToHtml(txt) {
  return txt.split('\n').map(l => l.trim()
    ? `<p style="margin:0 0 4px 0">${l.replace(/\*\*(.+?)\*\*/g,'<strong>$1</strong>')}</p>`
    : '<br>').join('');
}

async function startSend() {
  if (!state.contacts.length) { alert('Cargá una base primero'); return; }
  if (!state.asesores.filter(a=>a.active!==false).length) {
    alert('Configurá al menos un asesor activo'); return; }
  const subject = document.getElementById('subject').value.trim();
  if (!subject) { alert('Completá el asunto'); return; }

  document.getElementById('btnStart').disabled = true;
  document.getElementById('btnStop').disabled  = false;
  clearLog();

  const body_html = textToHtml(document.getElementById('bodyText').value);
  const res = await api('/api/send', { method:'POST', body: {
    contacts:    state.contacts,
    subject:     subject,
    body_html:   body_html,
    wa_number:   document.getElementById('waNumber').value,
    incluir_form:document.getElementById('incluirForm').checked,
    incluir_tel: document.getElementById('incluirTel').checked,
    delay:       parseFloat(document.getElementById('delay').value),
    base_id:     state.baseId,
    base_name:   state.baseName,
  }});

  if (res.job_id) {
    state.currentJob = res.job_id;
    pollJob();
  }
}

async function stopSend() {
  if (state.currentJob)
    await api(`/api/send/stop/${state.currentJob}`, {method:'POST'});
}

async function pollJob() {
  if (!state.currentJob) return;
  const s = await api(`/api/send/status/${state.currentJob}`);
  const total = s.total || 1;
  document.getElementById('kpiSent').textContent = s.sent || 0;
  document.getElementById('kpiErr').textContent  = s.errors || 0;
  const pct = Math.round(((s.sent||0)+(s.errors||0))/total*100);
  document.getElementById('kpiPct').textContent  = pct + '%';
  document.getElementById('progress').value = pct;

  // Log
  const lb = document.getElementById('logBox');
  lb.innerHTML = (s.log||[]).map(l =>
    `<div class="log-${l.tipo||'info'}">[${l.ts}] ${l.msg}</div>`
  ).join('');
  lb.scrollTop = lb.scrollHeight;

  if (s.status === 'running') {
    setTimeout(pollJob, 1500);
  } else {
    document.getElementById('btnStart').disabled = false;
    document.getElementById('btnStop').disabled  = true;
    state.currentJob = null;
  }
}

function clearLog() {
  document.getElementById('logBox').innerHTML = '';
  ['kpiSent','kpiErr','kpiPct'].forEach(id => document.getElementById(id).textContent = id==='kpiPct'?'0%':'0');
  document.getElementById('progress').value = 0;
}

// ── Historial ─────────────────────────────────────────────────────────────────
async function loadHistorial() {
  state.historial = await api('/api/historial');
  const tbody = document.getElementById('histBody');
  tbody.innerHTML = [...state.historial].reverse().slice(0,200).map(h => `
    <tr>
      <td>${(h.ts||'').slice(0,16).replace('T',' ')}</td>
      <td>${h.asesor||''}</td>
      <td style="font-size:11px">${h.to||''}</td>
      <td>${h.nombre||''}</td>
      <td class="hide-mobile">${h.carrera||''}</td>
      <td><span class="badge badge-${h.status==='enviado'?'green':'red'}">${h.status||''}</span></td>
    </tr>`).join('');
}

function exportHistorial() {
  const rows = [['Fecha','Asesor','Email','Nombre','Carrera','Estado'],
    ...state.historial.map(h=>[h.ts,h.asesor,h.to,h.nombre,h.carrera,h.status])];
  downloadCSV(rows, 'historial_kennedy.csv');
}

// ── Tracking ──────────────────────────────────────────────────────────────────
async function loadTracking() {
  const data = await api('/api/tracking_data');
  state.tracking = data;

  const aperturas  = data.filter(e=>e.accion==='open');
  const wa         = data.filter(e=>e.accion==='wa');
  const tel        = data.filter(e=>e.accion==='tel');
  const form       = data.filter(e=>e.accion==='form');
  const uniq = (arr,k='mail') => new Set(arr.map(e=>e[k]?.toLowerCase()).filter(Boolean)).size;

  document.getElementById('trackingKpis').innerHTML = [
    ['👁 Aperturas','var(--green)',  uniq(aperturas),  aperturas.length,  'únicas','totales'],
    ['💬 WhatsApp', 'var(--orange)', uniq(wa),         wa.length,         'únicos','totales'],
    ['📞 0800',     'var(--red)',    uniq(tel),         tel.length,        'únicos','totales'],
    ['📝 Formulario','var(--blue)',  uniq(form),        form.length,       'únicos','totales'],
  ].map(([t,c,v1,v2,l1,l2])=>`
    <div class="kpi">
      <div class="kpi-bar" style="background:${c}"></div>
      <div class="kpi-label">${t}</div>
      <div class="kpi-val" style="color:${c}">${v1}</div>
      <div class="kpi-sub">${l1} | ${v2} ${l2}</div>
    </div>`).join('');

  const icons = {open:'📧 Apertura',wa:'💬 WhatsApp',tel:'📞 0800',form:'📝 Formulario'};
  const colors= {open:'var(--green)',wa:'var(--orange)',tel:'var(--red)',form:'var(--blue)'};
  document.getElementById('trackBody').innerHTML =
    [...data].reverse().slice(0,100).map(e=>`
    <tr>
      <td>${(e.ts||'').slice(0,16).replace('T',' ')}</td>
      <td style="color:${colors[e.accion]||'inherit'}">${icons[e.accion]||e.accion}</td>
      <td>${e.nombre||''}</td>
      <td style="font-size:11px">${e.mail||''}</td>
      <td class="hide-mobile">${e.carrera||''}</td>
      <td class="hide-mobile" style="font-size:11px">${e.base_id||''}</td>
    </tr>`).join('');
}

// ── Ventas ────────────────────────────────────────────────────────────────────
async function loadVentas() {
  state.ventas = await api('/api/ventas');
  renderVentas();
}

function renderVentas() {
  const colors = {'Venta':'green','Tiene legajo':'blue','En seguimiento':'orange','Perdido':'red'};
  document.getElementById('ventasBody').innerHTML =
    [...state.ventas].reverse().map(v=>`
    <tr onclick="editarVenta(${v.id})" style="cursor:pointer">
      <td>${(v.ts||'').slice(0,10)}</td>
      <td>${v.nombre||''}</td>
      <td><span class="badge badge-${colors[v.estado]||'orange'}">${v.estado||''}</span></td>
      <td style="font-size:11px">${v.base||''}</td>
    </tr>`).join('');
}

async function buscarContacto() {
  const q = document.getElementById('ventaSearch').value.trim().toLowerCase();
  if (!q) return;
  const hist = await api('/api/historial');
  const track = await api('/api/tracking_data');

  const hm = hist.filter(h =>
    h.to?.toLowerCase().includes(q) || h.nombre?.toLowerCase().includes(q));
  const tm = track.filter(e => e.mail?.toLowerCase().includes(q));

  const el = document.getElementById('ventaResultado');
  if (!hm.length && !tm.length) {
    el.innerHTML = '<p style="color:var(--orange);font-size:12px;">⚠ No se encontró ningún contacto.</p>';
    return;
  }

  // Autocompletar
  if (hm.length) {
    const h = hm[0];
    document.getElementById('vNombre').value  = h.nombre || '';
    document.getElementById('vEmail').value   = h.to || '';
    document.getElementById('vBase').value    = h.base_id || h.base_name || '';
    document.getElementById('vCarrera').value = h.carrera || '';
    document.getElementById('vAsesor').value  = h.asesor || '';
  }

  const acciones = {};
  tm.forEach(e => { acciones[e.accion] = (acciones[e.accion]||0)+1; });
  const iconos = {open:'📧 Apertura',wa:'💬 WhatsApp',tel:'📞 0800',form:'📝 Form'};
  const resumen = Object.entries(acciones)
    .map(([k,v])=>`${iconos[k]||k}: ${v}`).join('  ·  ');

  el.innerHTML = `<div class="result-card">
    <div style="font-size:12px;font-weight:700;color:var(--green);margin-bottom:6px;">
      ✅ Encontrado — ${hm.length} envíos · ${tm.length} eventos</div>
    ${hm.length ? `<div style="font-size:12px;">📁 Base: ${hm[0].base_id||hm[0].base_name||'—'}</div>` : ''}
    ${resumen ? `<div style="font-size:12px;color:var(--green);margin-top:4px;">📊 ${resumen}</div>` : ''}
  </div>`;
}

function editarVenta(id) {
  const v = state.ventas.find(x=>x.id===id); if(!v) return;
  state.editingVenta = id;
  document.getElementById('vId').value      = id;
  document.getElementById('vNombre').value  = v.nombre||'';
  document.getElementById('vEmail').value   = v.email||'';
  document.getElementById('vTel').value     = v.telefono||'';
  document.getElementById('vBase').value    = v.base||'';
  document.getElementById('vCarrera').value = v.carrera||'';
  document.getElementById('vAsesor').value  = v.asesor||'';
  document.getElementById('vLegajo').value  = v.legajo||'';
  document.getElementById('vEstado').value  = v.estado||'En seguimiento';
  document.getElementById('vNotas').value   = v.notas||'';
}

async function guardarVenta() {
  const email = document.getElementById('vEmail').value.trim();
  const tel   = document.getElementById('vTel').value.trim();
  if (!email && !tel) { alert('Ingresá email o teléfono'); return; }
  const id = parseInt(document.getElementById('vId').value) || Date.now();
  const venta = {
    id, ts: new Date().toISOString(),
    nombre:   document.getElementById('vNombre').value,
    email, telefono: tel,
    base:     document.getElementById('vBase').value,
    carrera:  document.getElementById('vCarrera').value,
    asesor:   document.getElementById('vAsesor').value,
    legajo:   document.getElementById('vLegajo').value,
    estado:   document.getElementById('vEstado').value,
    notas:    document.getElementById('vNotas').value,
  };
  const idx = state.ventas.findIndex(v=>v.id===id);
  if (idx>=0) state.ventas[idx] = venta; else state.ventas.push(venta);
  await api('/api/ventas', {method:'POST', body: state.ventas});
  limpiarVenta(); renderVentas();
}

async function eliminarVenta() {
  const id = parseInt(document.getElementById('vId').value);
  if (!id || !confirm('¿Eliminar?')) return;
  state.ventas = state.ventas.filter(v=>v.id!==id);
  await api('/api/ventas', {method:'POST', body: state.ventas});
  limpiarVenta(); renderVentas();
}

function limpiarVenta() {
  state.editingVenta = null;
  ['vId','vNombre','vEmail','vTel','vBase','vCarrera','vAsesor','vLegajo','vNotas']
    .forEach(id => document.getElementById(id).value='');
  document.getElementById('vEstado').value = 'En seguimiento';
  document.getElementById('ventaResultado').innerHTML = '';
}

// ── Utilidades ────────────────────────────────────────────────────────────────
function downloadCSV(rows, fname) {
  const csv = rows.map(r=>r.map(c=>`"${(c||'').toString().replace(/"/g,'""')}"`).join(',')).join('\n');
  const a = document.createElement('a');
  a.href = 'data:text/csv;charset=utf-8,\uFEFF'+encodeURIComponent(csv);
  a.download = fname; a.click();
}

// ── Init ─────────────────────────────────────────────────────────────────────
loadAsesores();
</script>
</body>
</html>"""


# ── Upload endpoint para XLSX ─────────────────────────────────────────────────
@app.route("/api/upload", methods=["POST"])
def api_upload():
    try:
        import pandas as pd
        f = request.files.get("file")
        if not f: return jsonify({"error": "No file"}), 400
        ext = f.filename.rsplit(".", 1)[-1].lower()
        df  = pd.read_excel(f, engine="openpyxl") if ext == "xlsx" else pd.read_excel(f, engine="xlrd")
        df.columns = [c.strip().lower().replace(" ","_") for c in df.columns]

        # Mapear columnas
        def fcol(*opts):
            for o in opts:
                if o in df.columns: return o
            return None

        mail_col = fcol("mail","email","correo","e-mail")
        nom_col  = fcol("nombre","name")
        ape_col  = fcol("apellido","apellidos")
        car_col  = fcol("carrera","career")
        prov_col = fcol("provincia")
        leg_col  = fcol("legajo","cli_idoriginal")

        contacts = []
        for _, row in df.iterrows():
            mail = str(row.get(mail_col,"") if mail_col else "").strip()
            if "@" not in mail: continue
            contacts.append({
                "mail":     mail,
                "nombre":   str(row.get(nom_col,"") if nom_col else ""),
                "apellido": str(row.get(ape_col,"") if ape_col else ""),
                "carrera":  str(row.get(car_col,"") if car_col else ""),
                "provincia":str(row.get(prov_col,"") if prov_col else ""),
                "legajo":   str(row.get(leg_col,"") if leg_col else ""),
            })
        base_id = f.filename.rsplit(".",1)[0]
        return jsonify({"contacts": contacts, "base_id": base_id, "total": len(contacts)})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    print(f"Kennedy Web Server iniciando en puerto {port}...")
    app.run(host="0.0.0.0", port=port, debug=False)
