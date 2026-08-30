import json
import os
import re
import secrets
import smtplib
import string
import tempfile
import shutil
import uuid
import zipfile
from email.mime.text import MIMEText
from datetime import datetime, timedelta, timezone
from functools import wraps

import psycopg2
import psycopg2.extras
import requests
from flask import Flask, request, send_file, jsonify
from itsdangerous import URLSafeTimedSerializer, BadSignature, SignatureExpired
from werkzeug.security import generate_password_hash, check_password_hash

from converters import convert, text_to_pptx, extract_text, ConversionError

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25MB upload cap

MIME_TYPES = {
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "pdf": "application/pdf",
}

# CORS: locked to the live frontend.
ALLOWED_ORIGIN = "https://masterconvert-tau.vercel.app"
FRONTEND_URL = ALLOWED_ORIGIN

# AI features (Smart Summarize / drafting) call Claude directly.
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
ANTHROPIC_MODEL = "claude-sonnet-5"

# ---------- accounts / auth ----------
DATABASE_URL = os.environ.get("DATABASE_URL", "")
SECRET_KEY = os.environ.get("SECRET_KEY", "")
GOOGLE_CLIENT_ID = os.environ.get("GOOGLE_CLIENT_ID", "")

# Password-reset emails, sent via Gmail SMTP with an app password —
# console.google.com -> Security -> 2-Step Verification -> App passwords.
EMAIL_ADDRESS = os.environ.get("EMAIL_ADDRESS", "")
EMAIL_APP_PASSWORD = os.environ.get("EMAIL_APP_PASSWORD", "")

TOKEN_MAX_AGE = 60 * 60 * 24 * 30  # 30 days
RESET_TOKEN_MAX_AGE = 60 * 60  # 1 hour — short-lived on purpose
FREE_CONVERSIONS_LIMIT = 2
FREE_WINDOW_DAYS = 30

_serializer = URLSafeTimedSerializer(SECRET_KEY) if SECRET_KEY else None


def get_db():
    if not DATABASE_URL:
        raise ConversionError("Accounts need DATABASE_URL set on the server")
    conn = psycopg2.connect(DATABASE_URL, cursor_factory=psycopg2.extras.RealDictCursor)
    conn.autocommit = True
    return conn


def init_db():
    if not DATABASE_URL:
        return
    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS users (
                    id SERIAL PRIMARY KEY,
                    name TEXT NOT NULL,
                    email TEXT UNIQUE NOT NULL,
                    password_hash TEXT,
                    google_sub TEXT UNIQUE,
                    plan TEXT NOT NULL DEFAULT 'free',
                    referral_code TEXT UNIQUE NOT NULL,
                    bonus_credit_months INTEGER NOT NULL DEFAULT 0,
                    conversions_used INTEGER NOT NULL DEFAULT 0,
                    conversions_reset_at TIMESTAMPTZ NOT NULL DEFAULT now(),
                    created_at TIMESTAMPTZ NOT NULL DEFAULT now()
                )
                """
            )
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS conversions (
                    id SERIAL PRIMARY KEY,
                    user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
                    from_format TEXT NOT NULL,
                    to_format TEXT NOT NULL,
                    created_at TIMESTAMPTZ NOT NULL DEFAULT now()
                )
                """
            )
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS promo_codes (
                    id SERIAL PRIMARY KEY,
                    code TEXT UNIQUE NOT NULL,
                    created_by INTEGER REFERENCES users(id) ON DELETE SET NULL,
                    discount_desc TEXT NOT NULL,
                    expires_at TIMESTAMPTZ NOT NULL,
                    created_at TIMESTAMPTZ NOT NULL DEFAULT now()
                )
                """
            )
    finally:
        conn.close()


EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def random_suffix(n=4):
    alphabet = string.ascii_uppercase + string.digits
    return "".join(secrets.choice(alphabet) for _ in range(n))


def generate_referral_code(name):
    base = re.sub(r"[^A-Za-z]", "", name)[:8].upper() or "USER"
    return f"{base}{random_suffix()}"


def generate_promo_code():
    return "".join(secrets.choice(string.ascii_uppercase + string.digits) for _ in range(10))


def make_token(user_id):
    return _serializer.dumps({"user_id": user_id})


def verify_token(token):
    if not _serializer:
        return None
    try:
        data = _serializer.loads(token, max_age=TOKEN_MAX_AGE)
        return data.get("user_id")
    except (BadSignature, SignatureExpired):
        return None


def make_reset_token(user_id):
    return _serializer.dumps({"reset_user_id": user_id}, salt="password-reset")


def verify_reset_token(token):
    if not _serializer:
        return None
    try:
        data = _serializer.loads(token, max_age=RESET_TOKEN_MAX_AGE, salt="password-reset")
        return data.get("reset_user_id")
    except (BadSignature, SignatureExpired):
        return None


def send_email(to_email, subject, body):
    if not EMAIL_ADDRESS or not EMAIL_APP_PASSWORD:
        raise ConversionError("Email sending isn't configured on the server yet")
    msg = MIMEText(body)
    msg["Subject"] = subject
    msg["From"] = EMAIL_ADDRESS
    msg["To"] = to_email
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=15) as server:
        server.login(EMAIL_ADDRESS, EMAIL_APP_PASSWORD)
        server.sendmail(EMAIL_ADDRESS, to_email, msg.as_string())


def user_row_to_dict(row):
    now = datetime.now(timezone.utc)
    used = row["conversions_used"]
    reset_at = row["conversions_reset_at"]
    if reset_at and now - reset_at > timedelta(days=FREE_WINDOW_DAYS):
        used = 0  # window has rolled over; reflect that even before the next write

    is_free = row["plan"] == "free"
    return {
        "id": row["id"],
        "name": row["name"],
        "email": row["email"],
        "plan": row["plan"],
        "referral_code": row["referral_code"],
        "bonus_credit_months": row["bonus_credit_months"],
        "conversions_remaining": max(0, FREE_CONVERSIONS_LIMIT - used) if is_free else None,
        "conversions_limit": FREE_CONVERSIONS_LIMIT if is_free else None,
    }


def auth_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if request.method == "OPTIONS":
            return "", 204
        header = request.headers.get("Authorization", "")
        token = header.split(" ", 1)[1] if header.startswith("Bearer ") else None
        user_id = verify_token(token) if token else None
        if not user_id:
            return jsonify({"error": "Not authenticated"}), 401
        conn = get_db()
        try:
            with conn.cursor() as cur:
                cur.execute("SELECT * FROM users WHERE id = %s", (user_id,))
                row = cur.fetchone()
        finally:
            conn.close()
        if not row:
            return jsonify({"error": "Not authenticated"}), 401
        request.current_user = row
        return fn(*args, **kwargs)

    return wrapper


def call_claude(system_prompt, user_message, max_tokens=600, use_search=False):
    if not ANTHROPIC_API_KEY:
        raise ConversionError("AI features need ANTHROPIC_API_KEY set on the server")
    payload = {
        "model": ANTHROPIC_MODEL,
        "max_tokens": max_tokens,
        "system": system_prompt,
        "messages": [{"role": "user", "content": user_message}],
    }
    if use_search:
        payload["tools"] = [{"type": "web_search_20260318", "name": "web_search", "max_uses": 5}]
    try:
        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": ANTHROPIC_API_KEY,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json",
            },
            json=payload,
            timeout=140 if use_search else 45,
        )
    except requests.exceptions.Timeout:
        raise ConversionError("AI request timed out — try again, or try a narrower topic")
    except requests.exceptions.RequestException as e:
        raise ConversionError(f"AI request failed: {e}")
    if resp.status_code != 200:
        raise ConversionError(f"AI request failed ({resp.status_code}): {resp.text[:200]}")
    data = resp.json()
    return "".join(b.get("text", "") for b in data.get("content", []) if b.get("type") == "text")


OUTLINE_LINE_RE = re.compile(r'^\s*(\d+(?:\.\d+)*)\.?\s+(.+?)\s*$')


def parse_outline(text):
    """Recognizes lines like '5.1 Discussion' or '5.1.1 Knowledge of Dietary
    Sources' as a heading with a nesting level equal to how many dot-separated
    numbers it has (5 -> level 1, 5.1 -> level 2, 5.1.1 -> level 3). Plain
    topic text with no such lines simply parses to an empty list."""
    sections = []
    for line in text.strip().split("\n"):
        m = OUTLINE_LINE_RE.match(line)
        if m:
            number = m.group(1)
            heading = m.group(2).strip()
            if heading:
                sections.append({"number": number, "heading": heading, "level": number.count(".") + 1})
    return sections


def search_for_sources(topic, max_results=4):
    """Runs a search-only Claude call and reads real title/url pairs straight out of
    Anthropic's own web_search_tool_result blocks — never from the model's text output,
    so a source can never be fabricated: it either came from a real search hit or it
    isn't in the list at all."""
    if not ANTHROPIC_API_KEY:
        return []
    payload = {
        "model": ANTHROPIC_MODEL,
        "max_tokens": 200,
        "system": (
            "Search the web for 3-4 real, credible sources (academic papers, reputable "
            "educational or scientific publications) relevant to the given topic. Once you've "
            "searched, just reply with the word Done — no summary needed."
        ),
        "messages": [{"role": "user", "content": topic}],
        "tools": [{"type": "web_search_20260318", "name": "web_search", "max_uses": 3}],
    }
    try:
        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": ANTHROPIC_API_KEY,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json",
            },
            json=payload,
            timeout=60,
        )
    except requests.exceptions.RequestException:
        return []
    if resp.status_code != 200:
        return []
    data = resp.json()
    sources, seen = [], set()
    for block in data.get("content", []):
        if block.get("type") != "web_search_tool_result":
            continue
        content = block.get("content")
        if not isinstance(content, list):
            continue
        for item in content:
            if item.get("type") != "web_search_result":
                continue
            url = (item.get("url") or "").strip()
            title = (item.get("title") or "").strip()
            if url and title and url not in seen:
                seen.add(url)
                sources.append({"title": title, "url": url})
    return sources[:max_results]


@app.after_request
def add_cors_headers(resp):
    resp.headers["Access-Control-Allow-Origin"] = ALLOWED_ORIGIN
    resp.headers["Access-Control-Allow-Methods"] = "POST, GET, OPTIONS"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization"
    resp.headers["Access-Control-Expose-Headers"] = "X-Batch-Summary"
    return resp


@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


# ---------- auth ----------


@app.route("/api/auth/signup", methods=["POST", "OPTIONS"])
def signup():
    if request.method == "OPTIONS":
        return "", 204
    if not DATABASE_URL or not SECRET_KEY:
        return jsonify({"error": "Accounts aren't configured on the server yet"}), 500

    data = request.get_json(silent=True) or {}
    name = (data.get("name") or "").strip()
    email = (data.get("email") or "").strip().lower()
    password = data.get("password") or ""

    if not name or not email or not password:
        return jsonify({"error": "Name, email, and password are all required"}), 400
    if not EMAIL_RE.match(email):
        return jsonify({"error": "That doesn't look like a valid email"}), 400
    if len(password) < 8:
        return jsonify({"error": "Password must be at least 8 characters"}), 400

    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT id FROM users WHERE email = %s", (email,))
            if cur.fetchone():
                return jsonify({"error": "An account with that email already exists"}), 409

            referral_code = generate_referral_code(name)
            cur.execute(
                """
                INSERT INTO users (name, email, password_hash, referral_code)
                VALUES (%s, %s, %s, %s)
                RETURNING *
                """,
                (name, email, generate_password_hash(password), referral_code),
            )
            row = cur.fetchone()
    finally:
        conn.close()

    return jsonify({"token": make_token(row["id"]), "user": user_row_to_dict(row)})


@app.route("/api/auth/login", methods=["POST", "OPTIONS"])
def login():
    if request.method == "OPTIONS":
        return "", 204
    if not DATABASE_URL or not SECRET_KEY:
        return jsonify({"error": "Accounts aren't configured on the server yet"}), 500

    data = request.get_json(silent=True) or {}
    email = (data.get("email") or "").strip().lower()
    password = data.get("password") or ""

    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM users WHERE email = %s", (email,))
            row = cur.fetchone()
    finally:
        conn.close()

    if not row or not row["password_hash"] or not check_password_hash(row["password_hash"], password):
        return jsonify({"error": "Incorrect email or password"}), 401

    return jsonify({"token": make_token(row["id"]), "user": user_row_to_dict(row)})


@app.route("/api/auth/forgot-password", methods=["POST", "OPTIONS"])
def forgot_password():
    if request.method == "OPTIONS":
        return "", 204
    if not DATABASE_URL or not SECRET_KEY:
        return jsonify({"error": "Accounts aren't configured on the server yet"}), 500
    if not EMAIL_ADDRESS or not EMAIL_APP_PASSWORD:
        return jsonify({"error": "Password reset emails aren't configured on the server yet"}), 500

    data = request.get_json(silent=True) or {}
    email = (data.get("email") or "").strip().lower()
    if not email:
        return jsonify({"error": "Email required"}), 400

    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT id, name FROM users WHERE email = %s", (email,))
            row = cur.fetchone()
    finally:
        conn.close()

    # Always the same response whether or not the email has an account —
    # never confirm or deny that in the response itself.
    if row:
        token = make_reset_token(row["id"])
        reset_link = f"{FRONTEND_URL}/?reset_token={token}"
        try:
            send_email(
                email,
                "Reset your MasterConvert password",
                f"Hi {row['name'] or ''},\n\n"
                "We received a request to reset your MasterConvert password. "
                "This link expires in 1 hour:\n\n"
                f"{reset_link}\n\n"
                "If you didn't request this, you can safely ignore this email — "
                "your password won't change unless you click the link above.",
            )
        except ConversionError:
            pass  # already validated config above; a transient send failure shouldn't leak state

    return jsonify({"message": "If that email has an account, a reset link has been sent."})


@app.route("/api/auth/reset-password", methods=["POST", "OPTIONS"])
def reset_password():
    if request.method == "OPTIONS":
        return "", 204
    if not DATABASE_URL or not SECRET_KEY:
        return jsonify({"error": "Accounts aren't configured on the server yet"}), 500

    data = request.get_json(silent=True) or {}
    token = data.get("token") or ""
    new_password = data.get("password") or ""

    if len(new_password) < 8:
        return jsonify({"error": "Password must be at least 8 characters"}), 400

    user_id = verify_reset_token(token)
    if not user_id:
        return jsonify({"error": "This reset link is invalid or has expired — request a new one"}), 400

    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute(
                "UPDATE users SET password_hash = %s WHERE id = %s",
                (generate_password_hash(new_password), user_id),
            )
    finally:
        conn.close()

    return jsonify({"message": "Password updated — you can log in with your new password now"})


@app.route("/api/auth/google", methods=["POST", "OPTIONS"])
def google_signin():
    if request.method == "OPTIONS":
        return "", 204
    if not DATABASE_URL or not SECRET_KEY:
        return jsonify({"error": "Accounts aren't configured on the server yet"}), 500
    if not GOOGLE_CLIENT_ID:
        return jsonify({"error": "Google sign-in isn't configured on the server yet"}), 500

    from google.oauth2 import id_token as google_id_token
    from google.auth.transport import requests as google_requests

    data = request.get_json(silent=True) or {}
    credential = data.get("credential") or ""
    try:
        payload = google_id_token.verify_oauth2_token(
            credential, google_requests.Request(), GOOGLE_CLIENT_ID
        )
    except ValueError:
        return jsonify({"error": "Google sign-in failed"}), 401

    google_sub = payload["sub"]
    email = (payload.get("email") or "").strip().lower()
    name = payload.get("name") or email.split("@")[0]

    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM users WHERE google_sub = %s OR email = %s", (google_sub, email))
            row = cur.fetchone()
            if row:
                if not row["google_sub"]:
                    cur.execute("UPDATE users SET google_sub = %s WHERE id = %s RETURNING *", (google_sub, row["id"]))
                    row = cur.fetchone()
            else:
                referral_code = generate_referral_code(name)
                cur.execute(
                    """
                    INSERT INTO users (name, email, google_sub, referral_code)
                    VALUES (%s, %s, %s, %s)
                    RETURNING *
                    """,
                    (name, email, google_sub, referral_code),
                )
                row = cur.fetchone()
    finally:
        conn.close()

    return jsonify({"token": make_token(row["id"]), "user": user_row_to_dict(row)})


@app.route("/api/auth/me", methods=["GET", "OPTIONS"])
@auth_required
def me():
    return jsonify({"user": user_row_to_dict(request.current_user)})


@app.route("/api/auth/update-profile", methods=["POST", "OPTIONS"])
@auth_required
def update_profile():
    data = request.get_json(silent=True) or {}
    name = (data.get("name") or "").strip()
    email = (data.get("email") or "").strip().lower()
    if not name or not email:
        return jsonify({"error": "Name and email are required"}), 400
    if not EMAIL_RE.match(email):
        return jsonify({"error": "That doesn't look like a valid email"}), 400

    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT id FROM users WHERE email = %s AND id != %s",
                (email, request.current_user["id"]),
            )
            if cur.fetchone():
                return jsonify({"error": "That email is already in use"}), 409
            cur.execute(
                "UPDATE users SET name = %s, email = %s WHERE id = %s RETURNING *",
                (name, email, request.current_user["id"]),
            )
            row = cur.fetchone()
    finally:
        conn.close()

    return jsonify({"user": user_row_to_dict(row)})


# ---------- plans / promo ----------


@app.route("/api/plan/upgrade", methods=["POST", "OPTIONS"])
@auth_required
def plan_upgrade():
    data = request.get_json(silent=True) or {}
    plan = (data.get("plan") or "").strip()
    referral_code = (data.get("referral_code") or "").strip().upper() or None
    if plan not in ("free", "payperuse", "monthly", "yearly"):
        return jsonify({"error": "Unknown plan"}), 400

    bonus_applied = False
    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("UPDATE users SET plan = %s WHERE id = %s RETURNING *", (plan, request.current_user["id"]))
            row = cur.fetchone()

            if referral_code and plan in ("monthly", "yearly"):
                cur.execute(
                    "SELECT * FROM users WHERE referral_code = %s AND id != %s",
                    (referral_code, row["id"]),
                )
                referrer = cur.fetchone()
                if referrer:
                    cur.execute(
                        "UPDATE users SET bonus_credit_months = bonus_credit_months + 1 WHERE id = %s",
                        (referrer["id"],),
                    )
                    cur.execute(
                        "UPDATE users SET bonus_credit_months = bonus_credit_months + 1 WHERE id = %s RETURNING *",
                        (row["id"],),
                    )
                    row = cur.fetchone()
                    bonus_applied = True
    finally:
        conn.close()

    return jsonify({"user": user_row_to_dict(row), "result": {"referral_bonus_applied": bonus_applied}})


@app.route("/api/promo/generate", methods=["POST", "OPTIONS"])
@auth_required
def promo_generate():
    data = request.get_json(silent=True) or {}
    discount_desc = (data.get("discount_desc") or "").strip() or "1 free month"
    try:
        valid_days = int(data.get("valid_days") or 7)
    except (TypeError, ValueError):
        valid_days = 7
    valid_days = max(1, min(valid_days, 90))

    code = generate_promo_code()
    expires_at = datetime.now(timezone.utc) + timedelta(days=valid_days)

    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                INSERT INTO promo_codes (code, created_by, discount_desc, expires_at)
                VALUES (%s, %s, %s, %s)
                """,
                (code, request.current_user["id"], discount_desc, expires_at),
            )
    finally:
        conn.close()

    return jsonify({"code": code, "expires_at": expires_at.isoformat()})


@app.route("/api/conversions/history", methods=["GET", "OPTIONS"])
@auth_required
def conversions_history():
    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT from_format, to_format, created_at
                FROM conversions WHERE user_id = %s
                ORDER BY created_at DESC LIMIT 50
                """,
                (request.current_user["id"],),
            )
            rows = cur.fetchall()
    finally:
        conn.close()

    return jsonify(
        {
            "conversions": [
                {
                    "from_format": r["from_format"],
                    "to_format": r["to_format"],
                    "created_at": r["created_at"].isoformat(),
                }
                for r in rows
            ]
        }
    )


def _log_conversion_and_consume(user_row, from_fmt, to_fmt):
    """Enforce the free-tier cap and record the conversion. Raises ConversionError if over the limit.

    Runs as one explicit transaction (not autocommit) so the row lock from
    SELECT ... FOR UPDATE actually holds until the increment is written —
    otherwise two conversions fired at once could both slip past the cap.
    """
    conn = get_db()
    conn.autocommit = False
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM users WHERE id = %s FOR UPDATE", (user_row["id"],))
            row = cur.fetchone()
            now = datetime.now(timezone.utc)
            used = row["conversions_used"]
            reset_at = row["conversions_reset_at"]
            if reset_at and now - reset_at > timedelta(days=FREE_WINDOW_DAYS):
                used = 0
                reset_at = now

            if row["plan"] == "free" and used >= FREE_CONVERSIONS_LIMIT:
                conn.rollback()
                raise ConversionError(
                    "You've used your 2 free conversions this month. Upgrade for unlimited."
                )

            cur.execute(
                "UPDATE users SET conversions_used = %s, conversions_reset_at = %s WHERE id = %s",
                (used + 1, reset_at, row["id"]),
            )
            cur.execute(
                "INSERT INTO conversions (user_id, from_format, to_format) VALUES (%s, %s, %s)",
                (row["id"], from_fmt, to_fmt),
            )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


# ---------- conversion + AI endpoints (unchanged logic, now behind auth) ----------


@app.route("/api/convert", methods=["POST", "OPTIONS"])
@auth_required
def convert_endpoint():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    from_fmt = (request.form.get("from") or "").lower().strip()
    to_fmt = (request.form.get("to") or "").lower().strip()
    style = (request.form.get("style") or "").lower().strip() or None

    if not from_fmt or not to_fmt:
        return jsonify({"error": "Missing 'from' or 'to' format"}), 400
    if file.filename == "":
        return jsonify({"error": "Empty filename"}), 400

    ext = file.filename.rsplit(".", 1)[-1].lower() if "." in file.filename else ""
    if ext != from_fmt:
        return jsonify({"error": f"File extension .{ext} doesn't match declared source format '{from_fmt}'"}), 400

    try:
        _log_conversion_and_consume(request.current_user, from_fmt, to_fmt)
    except ConversionError as e:
        return jsonify({"error": str(e)}), 402

    work_dir = tempfile.mkdtemp(prefix=f"mc_{uuid.uuid4().hex[:8]}_")
    try:
        src_path = os.path.join(work_dir, file.filename)
        file.save(src_path)

        try:
            result_path = convert(src_path, from_fmt, to_fmt, work_dir, style=style)
        except ConversionError as e:
            return jsonify({"error": str(e)}), 422
        except FileNotFoundError as e:
            return jsonify({"error": f"Required conversion tool missing on server: {e}"}), 500

        base_name = file.filename.rsplit(".", 1)[0]
        download_name = f"{base_name}.{to_fmt}"

        return send_file(
            result_path,
            mimetype=MIME_TYPES.get(to_fmt, "application/octet-stream"),
            as_attachment=True,
            download_name=download_name,
        )
    except Exception as e:
        return jsonify({"error": f"Conversion failed: {e}"}), 500
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


@app.route("/api/convert-batch", methods=["POST", "OPTIONS"])
@auth_required
def convert_batch_endpoint():
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "No files uploaded"}), 400
    if len(files) > 20:
        return jsonify({"error": "Maximum 20 files per batch"}), 400

    from_fmt = (request.form.get("from") or "").lower().strip()
    to_fmt = (request.form.get("to") or "").lower().strip()
    style = (request.form.get("style") or "").lower().strip() or None
    if not from_fmt or not to_fmt:
        return jsonify({"error": "Missing 'from' or 'to' format"}), 400

    work_dir = tempfile.mkdtemp(prefix=f"mc_batch_{uuid.uuid4().hex[:8]}_")
    details = []
    converted = []  # (result_path, download_name)
    try:
        for file in files:
            filename = file.filename or "unnamed"
            ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
            if ext != from_fmt:
                details.append({"filename": filename, "success": False, "error": f"Not a .{from_fmt} file"})
                continue

            try:
                _log_conversion_and_consume(request.current_user, from_fmt, to_fmt)
            except ConversionError as e:
                details.append({"filename": filename, "success": False, "error": str(e)})
                continue

            src_path = os.path.join(work_dir, filename)
            file.save(src_path)
            try:
                result_path = convert(src_path, from_fmt, to_fmt, work_dir, style=style)
                base_name = filename.rsplit(".", 1)[0]
                converted.append((result_path, f"{base_name}.{to_fmt}"))
                details.append({"filename": filename, "success": True})
            except ConversionError as e:
                details.append({"filename": filename, "success": False, "error": str(e)})
            except Exception:
                details.append({"filename": filename, "success": False, "error": "Conversion failed"})

        if not converted:
            return jsonify({"error": "No files could be converted", "details": details}), 422

        zip_path = os.path.join(work_dir, "converted_files.zip")
        used_names = set()
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for result_path, download_name in converted:
                name, n = download_name, 1
                while name in used_names:
                    base, _, extn = download_name.rpartition(".")
                    name = f"{base} ({n}).{extn}"
                    n += 1
                used_names.add(name)
                zf.write(result_path, arcname=name)

        summary = {
            "total": len(files),
            "succeeded": len(converted),
            "failed": len(files) - len(converted),
            "details": details,
        }
        response = send_file(
            zip_path, mimetype="application/zip", as_attachment=True, download_name="converted_files.zip"
        )
        response.headers["X-Batch-Summary"] = json.dumps(summary)
        return response
    except Exception as e:
        return jsonify({"error": f"Batch conversion failed: {e}"}), 500
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


@app.route("/api/text-to-pptx", methods=["POST", "OPTIONS"])
@auth_required
def text_to_pptx_endpoint():
    data = request.get_json(silent=True) or {}
    text = (data.get("text") or "").strip()
    if not text:
        return jsonify({"error": "No text provided"}), 400
    if len(text) > 50000:
        return jsonify({"error": "Text too long (50,000 character limit)"}), 400

    work_dir = tempfile.mkdtemp(prefix=f"mc_txt_{uuid.uuid4().hex[:8]}_")
    try:
        result_path = text_to_pptx(text, work_dir)
        return send_file(
            result_path,
            mimetype=MIME_TYPES["pptx"],
            as_attachment=True,
            download_name="Presentation.pptx",
        )
    except ConversionError as e:
        return jsonify({"error": str(e)}), 422
    except Exception as e:
        return jsonify({"error": f"Conversion failed: {e}"}), 500
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


@app.route("/api/extract-text", methods=["POST", "OPTIONS"])
@auth_required
def extract_text_endpoint():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "Empty filename"}), 400
    ext = file.filename.rsplit(".", 1)[-1].lower() if "." in file.filename else ""

    work_dir = tempfile.mkdtemp(prefix=f"mc_ext_{uuid.uuid4().hex[:8]}_")
    try:
        src_path = os.path.join(work_dir, file.filename)
        file.save(src_path)
        text = extract_text(src_path, ext)
        if len(text) > 20000:
            text = text[:20000]
        return jsonify({"text": text, "filename": file.filename})
    except ConversionError as e:
        return jsonify({"error": str(e)}), 422
    except Exception as e:
        return jsonify({"error": f"Couldn't read file: {e}"}), 500
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


@app.route("/api/summarize", methods=["POST", "OPTIONS"])
@auth_required
def summarize_endpoint():
    data = request.get_json(silent=True) or {}
    text = (data.get("text") or "").strip()
    if not text:
        return jsonify({"error": "No text provided"}), 400
    if len(text) > 20000:
        return jsonify({"error": "Text too long (20,000 character limit)"}), 400
    try:
        result = call_claude(
            system_prompt=(
                "You condense study notes into slide-ready bullet points for a tutoring app. "
                "Return 4 to 7 short bullet points capturing the material, one per line, no "
                "numbering, no markdown, no preamble or closing remarks — plain lines only."
            ),
            user_message=text,
            max_tokens=500,
        )
        bullets = [ln.strip(" -•\t") for ln in result.strip().split("\n") if ln.strip()]
        return jsonify({"bullets": bullets})
    except ConversionError as e:
        return jsonify({"error": str(e)}), 502


@app.route("/api/write", methods=["POST", "OPTIONS"])
@auth_required
def write_endpoint():
    data = request.get_json(silent=True) or {}
    topic = (data.get("topic") or "").strip()
    if not topic:
        return jsonify({"error": "No topic provided"}), 400
    if len(topic) > 4000:
        return jsonify({"error": "Topic too long (4,000 character limit)"}), 400

    outline = parse_outline(topic)

    try:
        if len(outline) >= 2:
            # Outline mode: e.g. "5.1 Discussion / 5.1.1 General Awareness / 5.1.2 ..."
            # Cap section count so one request can't run unbounded.
            outline = outline[:14]
            combined_headings = "; ".join(s["heading"] for s in outline)

            # One shared search across the whole outline — a bigger, single pool
            # of real sources every section draws from, instead of one search per
            # section (slower and more likely to fragment/duplicate sources).
            sources = search_for_sources(combined_headings, max_results=8)

            if sources:
                sources_block = "\n".join(f"[{i + 1}] {s['title']}" for i, s in enumerate(sources))
                base_prompt = (
                    "You are writing one subsection of a longer academic piece on "
                    f"\"{combined_headings}\". Write ONLY the subsection given below as the "
                    "user message — 100-160 words, plain prose. Below is a numbered list of "
                    "real sources relevant to the overall piece — cite them naturally with "
                    "their bracketed number, e.g. [1], [2], where genuinely relevant to this "
                    "particular subsection; leave sources unused if they don't fit here. Do "
                    "not invent any author names or additional sources beyond this list. No "
                    f"heading, no reference list — just the paragraph.\n\nSOURCES:\n{sources_block}"
                )
            else:
                sources = []
                base_prompt = (
                    "You are writing one subsection of a longer academic piece on "
                    f"\"{combined_headings}\". Write ONLY the subsection given below as the "
                    "user message — 100-160 words, plain prose. No verified sources were "
                    "found for this topic, so do not invent any citations, authors, or "
                    "sources. No heading, no reference list — just the paragraph."
                )

            for sec in outline:
                sec_text = call_claude(
                    system_prompt=base_prompt,
                    user_message=f"{sec['number']} {sec['heading']}",
                    max_tokens=400,
                    use_search=False,
                ).strip()
                sec["text"] = sec_text if len(sec_text) >= 25 else "(Couldn't generate this subsection — try again.)"

            return jsonify({
                "sections": [
                    {"number": s["number"], "heading": s["heading"], "level": s["level"], "text": s["text"]}
                    for s in outline
                ],
                "sources": sources,
            })

        # Simple mode: a single flat topic, unchanged from before.
        # Step 1: search only. Sources come straight from Anthropic's own structured
        # search-result blocks, not from anything the model writes — so every title/url
        # here is guaranteed to be something that was actually found, never invented.
        sources = search_for_sources(topic)

        # Step 2: a plain, tool-free writing call. No search loop to hang, run long,
        # or narrate about — just a single bounded completion, grounded in exactly the
        # real sources from step 1 (or told plainly that none were found).
        if sources:
            sources_block = "\n".join(f"[{i + 1}] {s['title']}" for i, s in enumerate(sources))
            system_prompt = (
                "Write a clear, well-informed paragraph, 150-220 words, for a student on the "
                "given topic. Below is a numbered list of real sources found for this exact "
                "topic — refer to them naturally in the text using their bracketed number, "
                "e.g. [1], [2], at points where that source's subject matter is genuinely "
                "relevant. Do not invent any author names, years, or additional sources beyond "
                "this list — the numbers are the only citation format to use. Not every "
                "sentence needs one, and it's fine to leave a source unused if it doesn't fit. "
                "Write plain prose only, no reference list at the end — the sources are shown "
                f"separately.\n\nSOURCES:\n{sources_block}"
            )
        else:
            system_prompt = (
                "Write a clear, well-informed paragraph, 150-220 words, for a student on the "
                "given topic. No verified sources were found for this specific topic, so do "
                "not invent any citations, authors, or sources — write in plain, well-informed "
                "prose instead, with no citation markers at all."
            )

        text = call_claude(
            system_prompt=system_prompt,
            user_message=topic,
            max_tokens=500,
            use_search=False,
        ).strip()

        if len(text) < 40:
            raise ConversionError(
                "Couldn't finish a draft for this topic — try again, or make the topic a bit narrower"
            )

        return jsonify({"text": text, "sources": sources})
    except ConversionError as e:
        return jsonify({"error": str(e)}), 502


init_db()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    app.run(host="0.0.0.0", port=port)
