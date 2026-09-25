import json
import logging
import base64
import os
import random
import re
import secrets
import smtplib
import string
import tempfile
import shutil
import time
import uuid
import zipfile
from collections import Counter
from urllib.parse import urlparse
from email.mime.text import MIMEText
from datetime import datetime, timedelta, timezone
from functools import wraps

import psycopg2
import psycopg2.extras
import requests
import jwt as pyjwt  # PyJWT — aliased since this file also uses "jwt" as a short variable name in a couple of places below
from flask import Flask, request, send_file, jsonify
from flask_compress import Compress
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from itsdangerous import URLSafeTimedSerializer, BadSignature, SignatureExpired
from werkzeug.security import generate_password_hash, check_password_hash

from converters import (
    convert, text_to_pptx, summary_slides_to_pptx, academic_essay_to_docx, extract_text, ConversionError,
    apply_pdf_operations, pdf_split, images_to_pdf, pdf_get_page_thumbnails, quiz_to_docx, preview_file,
    warm_up_libreoffice,
)
from slidestudio import build_template_deck, parse_pptx_to_deck, template_catalog, StudioError

# Load LibreOffice into memory in the background as each worker starts, so
# the first conversion after a deploy or restart isn't the slow one.
import threading as _threading
_threading.Thread(target=warm_up_libreoffice, daemon=True).start()

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB upload cap — raised from 25MB: confirmed directly that scanned/image-heavy PDFs (the PDF editor's primary use case) routinely exceed 25MB in ways a plain text document never would, and 25MB was rejecting genuinely legitimate files.

# Gzip text-ish responses (JSON API replies, Sage/chat history, quiz data) —
# cheap on this box compared to a LibreOffice conversion, and meaningfully
# cuts payload size on slow mobile connections. Deliberately NOT compressing
# docx/pptx/xlsx/pdf downloads: those formats are already internally
# compressed (zip-based), so re-gzipping them burns CPU for near-zero size
# gain.
app.config["COMPRESS_MIMETYPES"] = [
    "application/json", "text/html", "text/css", "text/javascript",
    "application/javascript", "text/plain", "application/manifest+json",
]
app.config["COMPRESS_MIN_SIZE"] = 500  # skip tiny replies — not worth the CPU
Compress(app)

# Rate limiting — login/signup/password-reset had no throttling at all, so
# nothing stopped a scripted credential-stuffing run or a mass password-reset
# email-bomb against a real account beyond the natural cost of the password
# hash itself. In-memory storage (the default) is genuinely fine here, not a
# shortcut: this service runs as a single Render instance (numInstances: 1),
# so there's no second process with its own separate counters to get out of
# sync with. If this ever moves to multiple instances, this needs a shared
# backend (e.g. Redis) instead, or each instance enforces its own limit
# independently and the effective limit multiplies by instance count.
# The same applies within this one instance: gunicorn runs 2 worker
# processes, each keeping its own count, so a limit here can let through
# up to twice its number. That's fine for what these limits are for —
# stopping scripts making hundreds or thousands of requests — not precise
# quotas (Remy's monthly allowance is exact; it's counted in the database).
limiter = Limiter(get_remote_address, app=app, storage_uri="memory://")


def _account_or_ip_key():
    """Counts requests per signed-in account rather than per network, so a
    whole school computer lab sharing one internet connection isn't treated
    as one person. Falls back to the IP address when there's no valid login."""
    auth = request.headers.get("Authorization", "")
    if auth.startswith("Bearer "):
        uid = verify_token(auth[7:])
        if uid:
            return f"user:{uid}"
    return get_remote_address()


# Per-account ceilings for the features that cost real money or heavy
# server time. They sit far above what a person uses — they're there to
# stop a script on one free account from running up the AI bill (Evidence-
# Based Writing makes up to three AI calls per message) or tying up the
# half-CPU server's LibreOffice conversions for everyone else.
AI_LIMIT = "30 per hour;100 per day"
HEAVY_LIMIT = "60 per hour"
BATCH_LIMIT = "10 per hour"   # each batch can be up to 20 files


@limiter.request_filter
def _exempt_cors_preflight():
    """Every route below handles OPTIONS as a CORS preflight (returns 204
    with no auth or rate-limit logic run). Flask-Limiter applies a route's
    limit to all of that route's methods by default, so without this, a
    client that had already used up its POST limit could then also get
    its OPTIONS preflight rejected — which would break the real request
    the browser was about to make right after, not just the one being
    limited."""
    return request.method == "OPTIONS"



logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("docently")

MIME_TYPES = {
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "pdf": "application/pdf",
}

# File-upload safety: a user-supplied filename must never be used
# directly to build a server-side path. Confirmed directly, not just in
# theory — a filename like "../../../../tmp/x.docx" actually escaped
# the intended temp working directory and wrote a file elsewhere on
# disk when tested against this exact codebase. The fix separates two
# things that were conflated before: the server-side disk path (which
# the user never sees, so it can just be a safe random name) from the
# download filename (which should stay human-readable, including
# non-ASCII scripts — werkzeug's own secure_filename was tested here
# too and rejected for this second purpose specifically, since it
# strips non-ASCII entirely: a Chinese filename like "简历.docx" comes
# back as just "docx", which would be a real, confusing regression for
# any user not writing in Latin script).
def safe_upload_filename(original_filename):
    """A safe, random server-side filename for a saved upload,
    preserving only a sanitized (lowercased, alphanumeric-only)
    extension from the original — never the original name itself."""
    ext = original_filename.rsplit(".", 1)[-1].lower() if original_filename and "." in original_filename else ""
    ext = re.sub(r"[^a-z0-9]", "", ext)[:10]
    return "upload_" + uuid.uuid4().hex + (f".{ext}" if ext else ""), ext


def safe_download_name(original_filename, fallback="download"):
    """Sanitizes a filename for safe use in a Content-Disposition
    header — strips path separators and control characters (which
    could break the header or inject something unintended) while
    preserving Unicode letters, spaces, and punctuation for
    readability, unlike secure_filename."""
    if not original_filename:
        return fallback
    cleaned = re.sub(r"[/\\\x00-\x1f\x7f]", "", original_filename).strip()
    return cleaned or fallback


# CORS: allow the app's known frontend origins. This has broken three times
# now on a too-narrow allowlist — first the MasterConvert->Docently rename,
# then a second Vercel deployment landing on docently-1.vercel.app instead
# of the original docently.vercel.app, then the Docently->Docente rename
# landing on docente-1-dusky.vercel.app. Note "docente" is NOT "docently"
# with a suffix chopped off — they share only the root "docent" and then
# diverge ("-e" vs "-ly"), so the pattern below matches on that shared
# root plus either ending, covering both brand generations at once. That
# way neither a future suffix change nor another rebrand locks everyone
# out again.
# Any *.vercel.app address is accepted. Every Vercel upload can land on a new
# address — docently-1, docente-1-dusky, and "manifest-2"/"manifest-3" when a
# project got named after the first file dragged in — and each time the old
# name-based rule refused it, so logins failed with "Couldn't connect" until
# the permanent link was re-pointed. This is safe for this app: sign-in
# travels as a token in the Authorization header (never a cookie) and a site
# can only read its own stored token, so another website can't act as a
# student even when its requests are accepted — the name rule was never
# guarding accounts, only blocking our own new uploads. A future custom
# domain can be added with EXTRA_ALLOWED_ORIGINS.
_ALLOWED_ORIGIN_PATTERN = re.compile(r"^https://[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.vercel\.app$")
ALLOWED_ORIGINS = {
    "https://docently.vercel.app",
    "https://masterconvert-tau.vercel.app",
}
# Additional origins can be added without a code change via this env var
# (comma-separated) — set on Render if a custom domain is added later.
ALLOWED_ORIGINS |= {o.strip() for o in os.environ.get("EXTRA_ALLOWED_ORIGINS", "").split(",") if o.strip()}
FRONTEND_URL = "https://docently.vercel.app"


def _origin_is_allowed(origin):
    if not origin:
        return False
    return origin in ALLOWED_ORIGINS or bool(_ALLOWED_ORIGIN_PATTERN.match(origin))

# AI features (Smart Summarize / drafting) call Claude directly.
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
ANTHROPIC_MODEL = "claude-sonnet-5"
# The strongest currently-available model, reserved for Sage (the app's
# general-purpose tutoring assistant, /api/sage below) specifically —
# Sonnet already handles the other three tools' narrower, structured-
# output tasks (a slide plan, a quiz, a citation-checked draft) more
# than capably, so there's little to gain there from the slower, more
# expensive model; Sage's job is open-ended tutoring help across
# whatever a student or tutor brings to it, which is exactly the kind
# of broad, unpredictable reasoning this model is worth spending on.
ANTHROPIC_MODEL_STRONG = "claude-opus-5"

# Optional — without this, Semantic Scholar's public API shares a very tight
# anonymous rate limit across every caller on Render's IP range, which is what
# was causing 429 Too Many Requests. Get a free key at
# https://www.semanticscholar.org/product/api#api-key-form and set it as this
# env var on Render to raise the limit substantially.
SEMANTIC_SCHOLAR_API_KEY = os.environ.get("SEMANTIC_SCHOLAR_API_KEY", "")

# ---------- accounts / auth ----------
DATABASE_URL = os.environ.get("DATABASE_URL", "")
SECRET_KEY = os.environ.get("SECRET_KEY", "")
GOOGLE_CLIENT_ID = os.environ.get("GOOGLE_CLIENT_ID", "")

# Password-reset emails, sent via Gmail SMTP with an app password —
# console.google.com -> Security -> 2-Step Verification -> App passwords.
EMAIL_ADDRESS = os.environ.get("EMAIL_ADDRESS", "")
EMAIL_APP_PASSWORD = os.environ.get("EMAIL_APP_PASSWORD", "")

# ---------- App Store / Play Store purchase verification ----------
# Neither store's client-side purchase flow can be built or tested from
# here — that requires a real native (or TWA) app wrapper, real
# registered in-app products in App Store Connect / Play Console, and
# these real platform credentials deployed as env vars, none of which
# exist yet. What's built below is the server-side half of IAP
# compliance: verifying a client-submitted purchase against Apple's/
# Google's own servers before granting a plan, which is required
# regardless of how the client-side purchase UI ends up being built,
# and is the part that was genuinely missing (see plan_upgrade below,
# which previously trusted a bare plan name from the client with no
# verification at all — a real, exploitable gap independent of either
# store's policies).
APPLE_APP_STORE_KEY_ID = os.environ.get("APPLE_APP_STORE_KEY_ID", "")
APPLE_APP_STORE_ISSUER_ID = os.environ.get("APPLE_APP_STORE_ISSUER_ID", "")
APPLE_APP_STORE_PRIVATE_KEY = os.environ.get("APPLE_APP_STORE_PRIVATE_KEY", "")  # .p8 file contents, PEM format
APPLE_BUNDLE_ID = os.environ.get("APPLE_BUNDLE_ID", "")
APPLE_APP_STORE_ENVIRONMENT = os.environ.get("APPLE_APP_STORE_ENVIRONMENT", "production")  # or "sandbox" for TestFlight/dev testing

GOOGLE_PLAY_SERVICE_ACCOUNT_JSON = os.environ.get("GOOGLE_PLAY_SERVICE_ACCOUNT_JSON", "")  # full service-account key JSON, as a string
GOOGLE_PLAY_PACKAGE_NAME = os.environ.get("GOOGLE_PLAY_PACKAGE_NAME", "")

# Maps each store's product/subscription id to this app's own internal
# plan name. These id strings are placeholders matching a conventional
# reverse-DNS naming scheme (com.<company>.<app>.<plan>) — they must be
# changed to whatever ids are actually registered as in-app products in
# App Store Connect and Play Console, which is console configuration,
# not something settable from here.
STORE_PRODUCT_TO_PLAN = {
    "com.docently.app.payperuse": "payperuse",
    "com.docently.app.monthly": "monthly",
    "com.docently.app.yearly": "yearly",
}

TOKEN_MAX_AGE = 60 * 60 * 24 * 30  # 30 days
RESET_TOKEN_MAX_AGE = 60 * 60  # 1 hour — short-lived on purpose
FREE_CONVERSIONS_LIMIT = 2
FREE_WINDOW_DAYS = 30

# Sage runs on the strongest available model with web search enabled —
# real per-message cost confirmed directly against Anthropic's current
# published rates: roughly $0.01-$0.02 for a typical message, more once
# a conversation's accumulated history or a triggered search is
# factored in. That's meaningfully more expensive than the flat,
# one-shot cost of a file conversion, and unlike a conversion, nothing
# about a chat naturally stops a user at one message — a monthly cap is
# the only thing standing between "a tutor asks Sage a few genuine
# questions" and "an unbounded, unmetered bill." Free tier gets enough
# to genuinely try it; paid tiers get a high ceiling meant to catch
# runaway or automated use, not to be hit by normal tutoring use.
SAGE_FREE_MESSAGES_LIMIT = 10
SAGE_PAID_MESSAGES_LIMIT = 200
SAGE_WINDOW_DAYS = 30

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
    try:
        conn = get_db()
    except Exception as e:
        logger.error(
            "init_db: could not connect to the database — the server will still "
            "boot, but every DB-backed route (auth, conversions, promo codes, "
            "memory) will fail until this is fixed. This is often an expired "
            "Render free Postgres instance. Error: %s", e,
        )
        return
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
            # This file has never had a real migration mechanism — every
            # column added after the table's first deployment (including
            # the ones already above, going by how bare CREATE TABLE IF
            # NOT EXISTS can't retroactively alter a table that already
            # exists in production) must have been added by hand outside
            # this code, which is a genuine gap: without something like
            # this, the two new columns below only ever land in a fresh
            # database, never the live one already running. Postgres's
            # own ADD COLUMN IF NOT EXISTS is safe to run unconditionally
            # on every startup, whether the columns already exist or not.
            cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS sage_messages_used INTEGER NOT NULL DEFAULT 0")
            cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS sage_messages_reset_at TIMESTAMPTZ NOT NULL DEFAULT now()")
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
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS user_memory (
                    id SERIAL PRIMARY KEY,
                    user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
                    note TEXT NOT NULL,
                    created_at TIMESTAMPTZ NOT NULL DEFAULT now()
                )
                """
            )
    except Exception as e:
        logger.error("init_db: table setup failed: %s", e)
    finally:
        conn.close()


def add_user_memory(user_id, note):
    """Appends one durable note the AI has picked up about a student's ongoing
    work — additive only, existing notes are never edited or removed."""
    note = (note or "").strip()[:400]
    if not (DATABASE_URL and user_id and note):
        return
    try:
        conn = get_db()
        try:
            with conn.cursor() as cur:
                cur.execute(
                    "INSERT INTO user_memory (user_id, note) VALUES (%s, %s)",
                    (user_id, note),
                )
        finally:
            conn.close()
    except Exception:
        pass  # memory is a nice-to-have — never let it break the actual response


def get_user_memory(user_id, limit=12):
    """Returns this student's most recent accumulated notes, oldest first, so
    later context reads as a running history rather than a jumbled list."""
    if not (DATABASE_URL and user_id):
        return []
    try:
        conn = get_db()
        try:
            with conn.cursor() as cur:
                cur.execute(
                    "SELECT note FROM user_memory WHERE user_id = %s ORDER BY created_at DESC LIMIT %s",
                    (user_id, limit),
                )
                rows = cur.fetchall()
        finally:
            conn.close()
        return [r["note"] for r in reversed(rows)]
    except Exception:
        return []


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
    return {
        "id": row["id"],
        "name": row["name"],
        "email": row["email"],
        "plan": row["plan"],
        "referral_code": row["referral_code"],
        "bonus_credit_months": row.get("bonus_credit_months", 0) or 0,
        # File conversions are free and unlimited on every plan now — no
        # limit/remaining count to report for any plan, free included,
        # rather than the free tier's old fixed cap.
        "conversions_remaining": None,
        "conversions_limit": None,
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


def call_claude(system_prompt, user_message=None, max_tokens=600, use_search=False, return_meta=False, model=None, messages=None):
    """messages, when given, is sent to the API exactly as the native
    multi-turn conversation it already is (alternating user/assistant
    turns) — the correct way to give Claude real conversational
    context, rather than flattening prior turns into one large string
    inside user_message the way the existing /api/write endpoint's own
    history handling does elsewhere in this file. Sage (the assistant
    this exists for) is a genuine back-and-forth conversation, so it
    gets the native form; user_message alone still works unchanged for
    every existing single-shot caller."""
    if not ANTHROPIC_API_KEY:
        raise ConversionError("AI features need ANTHROPIC_API_KEY set on the server")
    payload = {
        "model": model or ANTHROPIC_MODEL,
        "max_tokens": max_tokens,
        "system": system_prompt,
        "messages": messages if messages is not None else [{"role": "user", "content": user_message}],
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
            timeout=140 if use_search else 110,
        )
    except requests.exceptions.Timeout:
        raise ConversionError("AI request timed out — try again, or try a narrower topic")
    except requests.exceptions.RequestException as e:
        raise ConversionError(f"AI request failed: {e}")
    if resp.status_code != 200:
        raise ConversionError(f"AI request failed ({resp.status_code}): {resp.text[:200]}")
    data = resp.json()
    text = "".join(b.get("text", "") for b in data.get("content", []) if b.get("type") == "text")
    if return_meta:
        return {"text": text, "stop_reason": data.get("stop_reason")}
    return text


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


def search_web_sources(topic, max_results=4):
    """Runs a search-only Claude call and reads real title/url pairs straight out of
    Anthropic's own web_search_tool_result blocks — never from the model's text output,
    so a source can never be fabricated: it either came from a real search hit or it
    isn't in the list at all. Returns (sources, diagnostic) — diagnostic explains
    exactly why the list is empty, if it is, so that reason can reach the student
    instead of getting lost in server logs no one can see."""
    if not ANTHROPIC_API_KEY:
        logger.warning("search_web_sources: no ANTHROPIC_API_KEY set, skipping")
        return [], "web search skipped (no API key configured on the server)"
    payload = {
        "model": ANTHROPIC_MODEL,
        "max_tokens": 800,
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
    except requests.exceptions.RequestException as e:
        logger.warning("search_web_sources: request failed for %r: %s", topic[:80], e)
        return [], f"web search request failed: {e}"
    if resp.status_code != 200:
        logger.warning(
            "search_web_sources: non-200 (%s) for %r — body: %s",
            resp.status_code, topic[:80], resp.text[:500],
        )
        return [], f"web search returned HTTP {resp.status_code}: {resp.text[:200]}"
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
                sources.append({"title": title, "url": url, "authors": None, "year": None, "venue": None, "abstract": None})
    logger.info("search_web_sources: %r -> %d source(s)", topic[:80], len(sources))
    diag = None if sources else "web search ran but found no usable results for this topic"
    return sources[:max_results], diag


def search_semantic_scholar(query, max_results=4):
    """Searches Semantic Scholar's free public API for real peer-reviewed papers.
    Sends SEMANTIC_SCHOLAR_API_KEY when configured — without one, the anonymous
    rate limit is shared across every caller on Render's IP range and gets
    exhausted easily. On a 429, retries once after a short backoff before
    giving up, since these limits are usually per-minute and often clear fast.
    Keeps any hit with a real title and URL; an abstract is used as a grounding
    excerpt when present but is not required — many real papers, especially in
    regional/non-mainstream journals, have no abstract indexed, and excluding
    them was emptying the candidate pool entirely for niche topics. Fetches a
    larger relevance-matched pool than max_results, then sorts by publication
    year descending so the most recent matching papers are preferred — falls
    back to older ones only when too few recent matches exist for the topic.
    Returns (sources, diagnostic) — sources is [] on any failure (timeout,
    rate limit, bad response), with diagnostic explaining why."""
    headers = {"x-api-key": SEMANTIC_SCHOLAR_API_KEY} if SEMANTIC_SCHOLAR_API_KEY else {}
    params = {"query": query[:300], "limit": max(max_results * 3, 15), "fields": "title,url,year,authors,abstract,venue"}
    resp = None
    for attempt in range(2):
        try:
            resp = requests.get(
                "https://api.semanticscholar.org/graph/v1/paper/search",
                params=params,
                headers=headers,
                timeout=12,
            )
        except requests.exceptions.RequestException as e:
            logger.warning("search_semantic_scholar: request failed for %r: %s", query[:80], e)
            return [], f"Semantic Scholar request failed: {e}"
        if resp.status_code == 429 and attempt == 0:
            logger.warning("search_semantic_scholar: 429 on first attempt for %r, retrying after backoff", query[:80])
            time.sleep(2.5)
            continue
        break
    if resp.status_code != 200:
        logger.warning(
            "search_semantic_scholar: non-200 (%s) for %r — body: %s",
            resp.status_code, query[:80], resp.text[:500],
        )
        reason = "rate-limited (too many requests)" if resp.status_code == 429 else f"HTTP {resp.status_code}"
        return [], f"Semantic Scholar {reason}: {resp.text[:200]}"
    try:
        data = resp.json()
    except ValueError:
        logger.warning("search_semantic_scholar: non-JSON response for %r", query[:80])
        return [], "Semantic Scholar returned a non-JSON response"
    sources = []
    for paper in data.get("data", []):
        title = (paper.get("title") or "").strip()
        url = (paper.get("url") or "").strip()
        abstract = (paper.get("abstract") or "").strip()
        if not (title and url):
            continue
        author_names = [a.get("name", "").strip() for a in (paper.get("authors") or []) if a.get("name")]
        sources.append({
            "title": title,
            "url": url,
            "abstract": abstract or None,
            "authors": author_names or None,
            "year": paper.get("year"),
            "venue": (paper.get("venue") or "").strip() or None,
        })
    sources.sort(key=lambda s: s["year"] or 0, reverse=True)
    raw_count = len(data.get("data", []))
    logger.info("search_semantic_scholar: %r -> %d source(s) (raw hits: %d)", query[:80], len(sources), raw_count)
    diag = None if sources else (
        f"Semantic Scholar ran but returned {raw_count} raw hits, none usable" if raw_count
        else "Semantic Scholar found zero matching papers for this topic"
    )
    return sources[:max_results], diag


def search_for_sources(topic, max_results=8):
    """Combines two source lookups: Semantic Scholar for real peer-reviewed papers
    (with abstracts and real author/year, sorted most-recent-first) plus general
    web search to fill any remaining slots only when Semantic Scholar can't cover
    the topic on its own — general web sources carry no verifiable author or
    publication date, so they're a last resort, not the default. Semantic Scholar
    hits come first when both find matches for the same topic; deduplicated by URL.
    Returns (sources, diagnostic) — diagnostic is only set when sources ends up
    empty, combining both underlying reasons so the failure is never silent."""
    sources, seen = [], set()
    ss_sources, ss_diag = search_semantic_scholar(topic, max_results=max_results)
    for s in ss_sources:
        if s["url"] not in seen:
            seen.add(s["url"])
            sources.append(s)
    web_diag = None
    if len(sources) < max_results:
        web_sources, web_diag = search_web_sources(topic, max_results=max_results - len(sources))
        for s in web_sources:
            if s["url"] not in seen:
                seen.add(s["url"])
                sources.append(s)
    if not sources:
        combined_diag = "; ".join(d for d in (ss_diag, web_diag) if d) or "no sources found (unknown reason)"
        logger.warning("search_for_sources: ZERO total sources found for %r — %s", topic[:120], combined_diag)
        return [], combined_diag
    logger.info("search_for_sources: %r -> %d total source(s)", topic[:80], len(sources))
    return sources[:max_results], None


def _format_author_apa(full_name):
    """'John Smith' -> 'Smith, J.' for a References-list author entry."""
    parts = full_name.strip().split()
    if not parts:
        return full_name
    surname = parts[-1]
    initials = " ".join(p[0].upper() + "." for p in parts[:-1] if p)
    return f"{surname}, {initials}" if initials else surname


def _domain_org_name(url):
    """Falls back to a readable site name (e.g. 'Healthline') when a web
    source has no identifiable author — standard APA practice for websites."""
    try:
        host = urlparse(url).netloc.lower()
        if host.startswith("www."):
            host = host[4:]
        base = host.split(".")[0] if host else "Source"
        return base.replace("-", " ").title() or "Source"
    except Exception:
        return "Source"


def enrich_sources_for_apa(sources):
    """Computes a real APA in-text citation key and a full References-list
    entry for each source, from actual metadata only — nothing here is
    invented. Sources with no known individual author fall back to the
    organization/site name (standard APA practice); sources with no known
    year use 'n.d.'. Duplicate (author, year) pairs get a/b/c suffixes, as
    APA requires, applied consistently to both the in-text key and the
    reference entry."""
    enriched = []
    for s in sources:
        authors = s.get("authors")
        if authors:
            surnames = [a.split()[-1] for a in authors if a.split()]
            if len(surnames) == 1:
                author_intext = surnames[0]
            elif len(surnames) == 2:
                author_intext = f"{surnames[0]} & {surnames[1]}"
            else:
                author_intext = f"{surnames[0]} et al."
            author_refs_list = [_format_author_apa(a) for a in authors[:20]]
            if len(author_refs_list) == 1:
                author_refs = author_refs_list[0]
            else:
                author_refs = ", ".join(author_refs_list[:-1]) + ", & " + author_refs_list[-1]
        else:
            org = _domain_org_name(s["url"])
            author_intext = org
            author_refs = org

        year = s.get("year")
        year_display = str(year) if year else "n.d."
        enriched.append({**s, "author_intext": author_intext, "author_refs": author_refs,
                          "year_display": year_display, "base_key": f"{author_intext}|{year_display}"})

    key_counts = Counter(e["base_key"] for e in enriched)
    running = {}
    for e in enriched:
        if key_counts[e["base_key"]] > 1:
            idx = running.get(e["base_key"], 0)
            running[e["base_key"]] = idx + 1
            e["year_display"] += "abcdefghij"[idx] if idx < 10 else str(idx)
        e["citation_key"] = f"({e['author_intext']}, {e['year_display']})"

    for e in enriched:
        pieces = [f"{e['author_refs']} ({e['year_display']})."]
        title = e["title"].rstrip(".")
        pieces.append(f"{title}.")
        if e.get("venue"):
            pieces.append(f"{e['venue']}.")
        e["ref_text"] = " ".join(pieces)

    return enriched


def build_sources_block(sources):
    """Source list for the writing prompt, keyed by the exact APA citation
    the model must copy verbatim. Sources with an abstract (from Semantic
    Scholar) include a short excerpt, so the model can cite what the paper
    actually found instead of guessing from the title alone."""
    lines = []
    for s in sources:
        line = f"{s['citation_key']} {s['title']}"
        if s.get("abstract"):
            line += f"\n    Abstract: {s['abstract'][:400]}"
        lines.append(line)
    return "\n".join(lines)


def build_references_list(sources):
    """The client-facing References section: plain reference text plus its
    URL as a separate field (so the frontend can render the URL as a live
    link), alphabetized by author/organization the way APA requires."""
    ordered = sorted(sources, key=lambda s: s["author_intext"].lower())
    return [{"text": s["ref_text"], "url": s["url"]} for s in ordered]


@app.after_request
def add_cors_headers(resp):
    origin = request.headers.get("Origin", "")
    if _origin_is_allowed(origin):
        resp.headers["Access-Control-Allow-Origin"] = origin
        resp.headers["Vary"] = "Origin"
    resp.headers["Access-Control-Allow-Methods"] = "POST, GET, DELETE, OPTIONS"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization"
    resp.headers["Access-Control-Expose-Headers"] = "X-Batch-Summary"
    return resp


# Consistent JSON error responses for every status code the API can
# return — without these, Flask's own default HTML error pages leak
# through for cases the code never explicitly handles (an oversized
# upload, a bad route, a wrong HTTP method), which breaks every
# frontend call expecting resp.json() to work, confirmed directly: an
# oversized upload was actually returning a raw HTML page before this.
@app.errorhandler(413)
def handle_too_large(e):
    max_mb = app.config.get("MAX_CONTENT_LENGTH", 0) // (1024 * 1024)
    return jsonify({"error": f"That file is too large — the limit is {max_mb}MB."}), 413


@app.errorhandler(429)
def handle_rate_limited(e):
    # Flask-Limiter's default reply is an HTML page; the app reads replies
    # as JSON, so without this a limited user saw a parsing error ("Unexpected
    # token '<'") instead of an explanation. Remy's monthly allowance returns
    # its own 429 message directly and never reaches this handler.
    return jsonify({"error": "You've made a lot of requests in a short time. Please wait a few minutes, then try again."}), 429


@app.errorhandler(404)
def handle_not_found(e):
    return jsonify({"error": "That endpoint doesn't exist."}), 404


@app.errorhandler(405)
def handle_method_not_allowed(e):
    return jsonify({"error": "That method isn't allowed on this endpoint."}), 405


@app.errorhandler(500)
def handle_server_error(e):
    logger.error(f"Unhandled server error: {e}", exc_info=True)
    return jsonify({"error": "Something went wrong on our end. Please try again."}), 500


@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


# ---------- auth ----------


@app.route("/api/auth/signup", methods=["POST", "OPTIONS"])
@limiter.limit("10 per hour")
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
    if len(name) > 200:
        return jsonify({"error": "That name is too long (200 characters max)"}), 400
    if len(email) > 254:  # RFC 5321's own limit on a valid email address
        return jsonify({"error": "That email is too long"}), 400
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
            try:
                cur.execute(
                    """
                    INSERT INTO users (name, email, password_hash, referral_code)
                    VALUES (%s, %s, %s, %s)
                    RETURNING *
                    """,
                    (name, email, generate_password_hash(password), referral_code),
                )
                row = cur.fetchone()
            except psycopg2.errors.UniqueViolation:
                # The explicit check above has a real, if narrow, race: two
                # signups for the same email arriving at nearly the same
                # moment can both pass it before either commits. The
                # database's own unique constraint is the actual backstop,
                # and its violation is translated to the same clear
                # message the explicit check gives, rather than falling
                # through to a generic 500.
                conn.rollback()
                return jsonify({"error": "An account with that email already exists"}), 409
    finally:
        conn.close()

    return jsonify({"token": make_token(row["id"]), "user": user_row_to_dict(row)})


@app.route("/api/auth/login", methods=["POST", "OPTIONS"])
@limiter.limit("10 per minute; 30 per hour")
def login():
    if request.method == "OPTIONS":
        return "", 204
    if not DATABASE_URL or not SECRET_KEY:
        return jsonify({"error": "Accounts aren't configured on the server yet"}), 500

    data = request.get_json(silent=True) or {}
    email = (data.get("email") or "").strip().lower()
    password = data.get("password") or ""

    if not email or not password:
        return jsonify({"error": "Email and password are both required"}), 400

    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM users WHERE email = %s", (email,))
            row = cur.fetchone()
    finally:
        conn.close()

    # check_password_hash always runs, even for an email that doesn't
    # exist at all (against a fixed dummy hash in that case) — hashing
    # is deliberately the slow part of this check, and skipping it
    # specifically when the email is unknown would make that one case
    # measurably faster than a wrong-password case, which is exactly
    # the kind of timing gap that lets an attacker enumerate which
    # emails have accounts without ever seeing a password.
    dummy_hash = "pbkdf2:sha256:600000$0000000000000000$0000000000000000000000000000000000000000000000000000000000000000"
    stored_hash = (row["password_hash"] if row and row["password_hash"] else dummy_hash)
    password_ok = check_password_hash(stored_hash, password)

    if not row or not row["password_hash"] or not password_ok:
        return jsonify({"error": "Incorrect email or password"}), 401

    return jsonify({"token": make_token(row["id"]), "user": user_row_to_dict(row)})


@app.route("/api/auth/forgot-password", methods=["POST", "OPTIONS"])
@limiter.limit("5 per hour")
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
                "Reset your Docente password",
                f"Hi {row['name'] or ''},\n\n"
                "We received a request to reset your Docente password. "
                "This link expires in 1 hour:\n\n"
                f"{reset_link}\n\n"
                "If you didn't request this, you can safely ignore this email — "
                "your password won't change unless you click the link above.",
            )
        except ConversionError:
            pass  # already validated config above; a transient send failure shouldn't leak state

    return jsonify({"message": "If that email has an account, a reset link has been sent."})


@app.route("/api/auth/reset-password", methods=["POST", "OPTIONS"])
@limiter.limit("20 per hour")
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
            if cur.rowcount == 0:
                # The token's signature was valid, but the account it
                # points to no longer exists (deleted after the reset
                # link was sent, or similar) — telling the user their
                # password was updated when nothing was actually
                # changed would be a false success.
                return jsonify({"error": "This account no longer exists."}), 400
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
    if not credential:
        return jsonify({"error": "No Google credential was provided"}), 400
    try:
        payload = google_id_token.verify_oauth2_token(
            credential, google_requests.Request(), GOOGLE_CLIENT_ID
        )
    except ValueError:
        return jsonify({"error": "Google sign-in failed — that credential wasn't valid"}), 401
    except Exception as e:
        # verify_oauth2_token can also fail on a transient network/transport
        # problem reaching Google's own servers, not just an invalid
        # credential (a ValueError) — worth its own message, since "isn't
        # valid" would be misleading for what's actually a connectivity
        # blip on our end.
        logger.error(f"Google sign-in verification failed unexpectedly: {e}", exc_info=True)
        return jsonify({"error": "Couldn't reach Google to verify sign-in. Please try again."}), 503

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
                try:
                    cur.execute(
                        """
                        INSERT INTO users (name, email, google_sub, referral_code)
                        VALUES (%s, %s, %s, %s)
                        RETURNING *
                        """,
                        (name, email, google_sub, referral_code),
                    )
                    row = cur.fetchone()
                except psycopg2.errors.UniqueViolation:
                    # The same narrow race as signup's — two sign-in
                    # attempts for the same brand-new Google account
                    # landing at nearly the same moment. Unlike signup,
                    # this isn't really the user's fault and isn't a
                    # "duplicate account" in the same sense — the other
                    # concurrent request already created the account by
                    # now, so re-querying and continuing normally
                    # fulfills the same "sign in with Google" intent
                    # instead of surfacing an error for it.
                    conn.rollback()
                    cur.execute("SELECT * FROM users WHERE google_sub = %s OR email = %s", (google_sub, email))
                    row = cur.fetchone()
                    if not row:
                        return jsonify({"error": "Google sign-in failed. Please try again."}), 500
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
    if len(name) > 200:
        return jsonify({"error": "That name is too long (200 characters max)"}), 400
    if len(email) > 254:
        return jsonify({"error": "That email is too long"}), 400
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
            try:
                cur.execute(
                    "UPDATE users SET name = %s, email = %s WHERE id = %s RETURNING *",
                    (name, email, request.current_user["id"]),
                )
                row = cur.fetchone()
            except psycopg2.errors.UniqueViolation:
                # Same narrow race as signup/google sign-in: another
                # account could claim this exact email between the
                # check above and this update.
                conn.rollback()
                return jsonify({"error": "That email is already in use"}), 409
    finally:
        conn.close()

    return jsonify({"user": user_row_to_dict(row)})


# ---------- plans / promo ----------


class PurchaseVerificationError(Exception):
    """Raised for any reason a submitted purchase can't be trusted —
    wrong app, expired, refunded, or the platform's own API rejected
    it — so plan_upgrade has exactly one place to catch and turn into
    a clean error response, rather than three different failure shapes
    from two different platforms' verification functions."""
    pass


def _apple_verify_transaction(transaction_id):
    """Verifies one In-App Purchase transaction against Apple's App
    Store Server API and returns its product id — the modern
    replacement for the deprecated /verifyReceipt endpoint. Requires
    an App Store Connect API key (APPLE_APP_STORE_KEY_ID / _ISSUER_ID /
    _PRIVATE_KEY) with the "App Manager" or "Customer Support" role,
    generated in App Store Connect -> Users and Access -> Integrations
    -> App Store Connect API — none of which can be generated from
    here, since it requires an actual registered app and an Apple
    Developer account.

    Two steps: (1) sign a short-lived ES256 JWT to authenticate this
    server to Apple, per Apple's documented auth scheme for this API;
    (2) call GET /inApps/v1/transactions/{id}, which returns the
    transaction as a signed JWT (JWS) in its own right. This decodes
    that response's payload without additionally verifying its x5c
    certificate chain against Apple's root CA — a deliberate,
    documented scope limit, not an oversight: the request that
    fetched it was itself authenticated to Apple's own server over
    TLS using our own signed credential, which is the actual trust
    boundary here, and full chain-of-trust verification is a
    meaningfully larger cryptographic undertaking that belongs in its
    own careful pass rather than folded into this one silently.
    """
    if not (APPLE_APP_STORE_KEY_ID and APPLE_APP_STORE_ISSUER_ID and APPLE_APP_STORE_PRIVATE_KEY):
        raise PurchaseVerificationError(
            "Apple purchase verification isn't configured on this server yet "
            "(APPLE_APP_STORE_KEY_ID / _ISSUER_ID / _PRIVATE_KEY)."
        )
    now = int(time.time())
    auth_token = pyjwt.encode(
        {
            "iss": APPLE_APP_STORE_ISSUER_ID,
            "iat": now,
            "exp": now + 300,  # Apple caps this token at 60 minutes; 5 is plenty for one call
            "aud": "appstoreconnect-v1",
            "bid": APPLE_BUNDLE_ID,
        },
        APPLE_APP_STORE_PRIVATE_KEY,
        algorithm="ES256",
        headers={"kid": APPLE_APP_STORE_KEY_ID, "typ": "JWT"},
    )

    host = (
        "https://api.storekit-sandbox.itunes.apple.com"
        if APPLE_APP_STORE_ENVIRONMENT == "sandbox"
        else "https://api.storekit.itunes.apple.com"
    )
    try:
        resp = requests.get(
            f"{host}/inApps/v1/transactions/{transaction_id}",
            headers={"Authorization": f"Bearer {auth_token}"},
            timeout=15,
        )
    except requests.RequestException as e:
        raise PurchaseVerificationError(f"Couldn't reach Apple to verify this purchase: {e}")

    if resp.status_code != 200:
        raise PurchaseVerificationError(f"Apple rejected this transaction (status {resp.status_code}).")

    signed_transaction = resp.json().get("signedTransactionInfo")
    if not signed_transaction:
        raise PurchaseVerificationError("Apple's response didn't include transaction data.")

    # options={"verify_signature": False}: intentional, and only safe because
    # of the trust boundary explained above — this decodes the payload Apple
    # already returned over an authenticated connection, it does not accept
    # an arbitrary client-supplied JWS as truth.
    payload = pyjwt.decode(signed_transaction, options={"verify_signature": False})

    if APPLE_BUNDLE_ID and payload.get("bundleId") != APPLE_BUNDLE_ID:
        raise PurchaseVerificationError("This transaction belongs to a different app.")
    if payload.get("revocationDate"):
        raise PurchaseVerificationError("This purchase was refunded or revoked.")
    expires_ms = payload.get("expiresDate")
    if expires_ms and expires_ms < now * 1000:
        raise PurchaseVerificationError("This subscription has expired.")

    product_id = payload.get("productId")
    if product_id not in STORE_PRODUCT_TO_PLAN:
        raise PurchaseVerificationError(f"Unrecognized product id '{product_id}'.")
    return STORE_PRODUCT_TO_PLAN[product_id]


def _google_verify_purchase(product_id, purchase_token, is_subscription):
    """Verifies one Google Play purchase against the Play Developer API
    and returns the mapped internal plan name. Requires a service
    account (GOOGLE_PLAY_SERVICE_ACCOUNT_JSON) granted access to this
    app in Play Console -> Setup -> API access — again, console
    configuration against a real registered app, not buildable here.

    A PWA wrapped as a TWA (the path noted in this app's own earlier
    plan, not a native Android app) surfaces purchases through the
    browser's Digital Goods API / Payment Request API rather than the
    native Play Billing Library, but the purchase token those APIs
    return is verified server-side exactly the same way as any other
    Play purchase — this function doesn't need to know or care which
    client-side API produced the token.

    Two steps, the standard OAuth2 service-account flow: (1) sign a
    JWT with the service account's private key and exchange it for an
    access token; (2) call the Play Developer API's
    purchases.subscriptionsv2.get (subscriptions) or
    purchases.products.get (one-time purchases) with that token.
    """
    if not (GOOGLE_PLAY_SERVICE_ACCOUNT_JSON and GOOGLE_PLAY_PACKAGE_NAME):
        raise PurchaseVerificationError(
            "Google Play purchase verification isn't configured on this server yet "
            "(GOOGLE_PLAY_SERVICE_ACCOUNT_JSON / GOOGLE_PLAY_PACKAGE_NAME)."
        )
    try:
        service_account = json.loads(GOOGLE_PLAY_SERVICE_ACCOUNT_JSON)
    except json.JSONDecodeError:
        raise PurchaseVerificationError("Google Play service account credentials are malformed.")

    now = int(time.time())
    assertion = pyjwt.encode(
        {
            "iss": service_account["client_email"],
            "scope": "https://www.googleapis.com/auth/androidpublisher",
            "aud": "https://oauth2.googleapis.com/token",
            "iat": now,
            "exp": now + 3600,
        },
        service_account["private_key"],
        algorithm="RS256",
    )
    try:
        token_resp = requests.post(
            "https://oauth2.googleapis.com/token",
            data={"grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer", "assertion": assertion},
            timeout=15,
        )
        token_resp.raise_for_status()
        access_token = token_resp.json()["access_token"]
    except (requests.RequestException, KeyError) as e:
        raise PurchaseVerificationError(f"Couldn't authenticate to Google Play: {e}")

    pkg = GOOGLE_PLAY_PACKAGE_NAME
    if is_subscription:
        url = f"https://androidpublisher.googleapis.com/androidpublisher/v3/applications/{pkg}/purchases/subscriptionsv2/tokens/{purchase_token}"
    else:
        url = f"https://androidpublisher.googleapis.com/androidpublisher/v3/applications/{pkg}/purchases/products/{product_id}/tokens/{purchase_token}"

    try:
        resp = requests.get(url, headers={"Authorization": f"Bearer {access_token}"}, timeout=15)
    except requests.RequestException as e:
        raise PurchaseVerificationError(f"Couldn't reach Google Play to verify this purchase: {e}")
    if resp.status_code != 200:
        raise PurchaseVerificationError(f"Google Play rejected this purchase (status {resp.status_code}).")

    data = resp.json()
    if is_subscription:
        state = data.get("subscriptionState")
        if state not in ("SUBSCRIPTION_STATE_ACTIVE", "SUBSCRIPTION_STATE_IN_GRACE_PERIOD"):
            raise PurchaseVerificationError(f"This subscription isn't currently active (state: {state}).")
    else:
        if data.get("purchaseState") != 0:  # 0 == purchased; 1 == canceled, 2 == pending
            raise PurchaseVerificationError("This purchase isn't in a completed state.")

    if product_id not in STORE_PRODUCT_TO_PLAN:
        raise PurchaseVerificationError(f"Unrecognized product id '{product_id}'.")
    return STORE_PRODUCT_TO_PLAN[product_id]


@app.route("/api/plan/upgrade", methods=["POST", "OPTIONS"])
@auth_required
def plan_upgrade():
    """Downgrading to free needs no verification at all — giving up paid
    access is always safe to grant on request. Every paid plan, on the
    other hand, is only ever granted after _apple_verify_transaction or
    _google_verify_purchase has independently confirmed it against that
    store's own servers; the plan name that ends up written to the
    database comes from that verified mapping, never from whatever the
    client claims it purchased. This replaces an endpoint that
    previously wrote any client-supplied plan string straight into the
    database with no purchase check at all — a real, exploitable gap
    (any authenticated user could grant themselves any paid plan with a
    single request) independent of either store's own policies, and the
    reason this rewrite exists at all, not just a compliance nicety."""
    data = request.get_json(silent=True) or {}
    platform = (data.get("platform") or "").strip().lower()
    referral_code = (data.get("referral_code") or "").strip().upper() or None

    if platform == "free":
        plan = "free"
    elif platform == "ios":
        transaction_id = (data.get("transaction_id") or "").strip()
        if not transaction_id:
            return jsonify({"error": "Missing transaction_id."}), 400
        try:
            plan = _apple_verify_transaction(transaction_id)
        except PurchaseVerificationError as e:
            return jsonify({"error": str(e)}), 402
    elif platform == "android":
        product_id = (data.get("product_id") or "").strip()
        purchase_token = (data.get("purchase_token") or "").strip()
        is_subscription = bool(data.get("is_subscription", True))
        if not (product_id and purchase_token):
            return jsonify({"error": "Missing product_id or purchase_token."}), 400
        try:
            plan = _google_verify_purchase(product_id, purchase_token, is_subscription)
        except PurchaseVerificationError as e:
            return jsonify({"error": str(e)}), 402
    else:
        return jsonify({"error": "platform must be 'free', 'ios', or 'android'."}), 400

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


@app.route("/api/conversions/history", methods=["DELETE", "OPTIONS"])
@auth_required
def clear_conversions_history():
    """Deletes this user's own conversion history rows — the record of
    past conversions used to populate the home screen's Recent Activity
    list, not the converted files themselves (those were never stored
    server-side in the first place; each download happens directly from
    the temporary work directory of that request and is gone once it
    completes). A real, working action, not another cosmetic button —
    unlike the pre-existing "Clear Local File Cache" button beside it in
    Account Settings, which has nothing real to act on in a web app."""
    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM conversions WHERE user_id = %s", (request.current_user["id"],))
    finally:
        conn.close()
    return jsonify({"success": True})


def _log_conversion_and_consume(user_row, from_fmt, to_fmt):
    """Records the conversion for Recent Activity and usage tracking.
    File conversions are free and unlimited on every plan — this no
    longer enforces FREE_CONVERSIONS_LIMIT at all, by explicit
    decision; conversions_used keeps incrementing (still meaningful as
    a usage signal, and reset_at keeps rolling over the same way) but
    nothing here ever raises ConversionError over it anymore.
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


def _refund_sage_message(user_row):
    """Returns one Remy message to the user's monthly allowance, for a
    request that was counted but never got a reply. Never goes below zero
    (the allowance may have reset in between). Best-effort: a failure here
    is logged, never shown to the user on top of the original error."""
    conn = None
    try:
        conn = get_db()
        with conn.cursor() as cur:
            cur.execute(
                "UPDATE users SET sage_messages_used = GREATEST(sage_messages_used - 1, 0) WHERE id = %s",
                (user_row["id"],),
            )
        conn.commit()
    except Exception as e:
        logger.error("Couldn't refund a Remy message for user %s: %s", user_row.get("id"), e)
        if conn:
            conn.rollback()
    finally:
        if conn:
            conn.close()


def _log_sage_message_and_consume(user_row):
    """Enforce Sage's own monthly message cap and record the message —
    same row-locking pattern as _log_conversion_and_consume above, for
    the same reason (a concurrent pair of requests must not both slip
    past the cap), but with its own counter and its own limit that
    applies on every plan, not just free: SAGE_PAID_MESSAGES_LIMIT
    exists specifically because "paid" doesn't mean "an unbounded bill
    is fine," it means a much higher ceiling aimed at runaway or
    automated use rather than genuine tutoring conversations."""
    conn = get_db()
    conn.autocommit = False
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM users WHERE id = %s FOR UPDATE", (user_row["id"],))
            row = cur.fetchone()
            now = datetime.now(timezone.utc)
            used = row["sage_messages_used"]
            reset_at = row["sage_messages_reset_at"]
            if reset_at and now - reset_at > timedelta(days=SAGE_WINDOW_DAYS):
                used = 0
                reset_at = now

            limit = SAGE_FREE_MESSAGES_LIMIT if row["plan"] == "free" else SAGE_PAID_MESSAGES_LIMIT
            if used >= limit:
                conn.rollback()
                if row["plan"] == "free":
                    raise ConversionError(
                        f"You've used your {SAGE_FREE_MESSAGES_LIMIT} free messages with Remy this month. "
                        "Upgrade for a much higher limit."
                    )
                raise ConversionError(
                    f"You've reached Remy's monthly limit ({SAGE_PAID_MESSAGES_LIMIT} messages) for this "
                    "account. It resets next month."
                )

            cur.execute(
                "UPDATE users SET sage_messages_used = %s, sage_messages_reset_at = %s WHERE id = %s",
                (used + 1, reset_at, row["id"]),
            )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


# ---------- conversion + AI endpoints (unchanged logic, now behind auth) ----------


@app.route("/api/convert", methods=["POST", "OPTIONS"])
@limiter.limit(HEAVY_LIMIT, key_func=_account_or_ip_key)
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

    # Fix Slides' own font/format options — read and validated here,
    # right at the API boundary, rather than trusting raw form input
    # all the way into a file-generation pipeline. Every field is
    # optional; pptx_to_pptx already has sensible defaults (and a
    # template-derived style, when a template's uploaded) for anything
    # left unset. Only ever built for the pptx->pptx pair specifically.
    pptx_options = None
    if from_fmt == "pptx" and to_fmt == "pptx":
        pptx_options = {}
        for size_field, key in (("title_size", "title_size"), ("body_size", "body_size")):
            raw = request.form.get(size_field)
            if raw:
                try:
                    size_val = float(raw)
                except ValueError:
                    return jsonify({"error": f"'{size_field}' must be a number"}), 400
                if not (6 <= size_val <= 200):
                    return jsonify({"error": f"'{size_field}' must be between 6 and 200 points"}), 400
                pptx_options[key] = size_val
        for font_field in ("title_font", "body_font"):
            raw = (request.form.get(font_field) or "").strip()
            if raw:
                if len(raw) > 100:
                    return jsonify({"error": f"'{font_field}' is too long"}), 400
                pptx_options[font_field] = raw
        for color_field in ("title_color", "body_color"):
            raw = (request.form.get(color_field) or "").strip().lstrip("#")
            if raw:
                if not re.fullmatch(r"[0-9A-Fa-f]{6}", raw):
                    return jsonify({"error": f"'{color_field}' must be a 6-digit hex color"}), 400
                pptx_options[color_field] = raw
        title_align = (request.form.get("title_align") or "").strip().lower()
        if title_align:
            if title_align not in ("left", "center", "right"):
                return jsonify({"error": "'title_align' must be left, center, or right"}), 400
            pptx_options["title_align"] = title_align

    try:
        _log_conversion_and_consume(request.current_user, from_fmt, to_fmt)
    except ConversionError as e:
        return jsonify({"error": str(e)}), 402

    work_dir = tempfile.mkdtemp(prefix=f"mc_{uuid.uuid4().hex[:8]}_")
    try:
        safe_name, _ext = safe_upload_filename(file.filename)
        src_path = os.path.join(work_dir, safe_name)
        file.save(src_path)

        try:
            result_path = convert(src_path, from_fmt, to_fmt, work_dir, style=style, pptx_options=pptx_options)
        except ConversionError as e:
            return jsonify({"error": str(e)}), 422
        except FileNotFoundError as e:
            return jsonify({"error": f"Required conversion tool missing on server: {e}"}), 500

        base_name = safe_download_name(file.filename).rsplit(".", 1)[0]
        download_name = f"{base_name}.{to_fmt}"

        return send_file(
            result_path,
            mimetype=MIME_TYPES.get(to_fmt, "application/octet-stream"),
            as_attachment=True,
            download_name=download_name,
        )
    except Exception as e:
        logger.error(f"Conversion failed unexpectedly: {e}", exc_info=True)
        return jsonify({"error": "Conversion failed. Please try again."}), 500
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


@app.route("/api/preview-file", methods=["POST", "OPTIONS"])
@limiter.limit(HEAVY_LIMIT, key_func=_account_or_ip_key)
@auth_required
def preview_file_endpoint():
    """Takes a file the frontend already has in hand — one it just
    produced via a conversion or an AI export — and returns a PDF of
    it for the browser to render natively, so "preview" means the
    genuine converted or generated file's real content, not a mocked
    stand-in. Deliberately a separate, additive endpoint rather than a
    change to convert_endpoint or the AI export endpoints themselves:
    every one of those keeps returning exactly what it already did,
    so nothing already working is put at risk by adding this."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "Empty filename"}), 400

    ext = file.filename.rsplit(".", 1)[-1].lower() if "." in file.filename else ""
    if ext not in ("pdf", "docx", "pptx", "xlsx", "doc", "ppt", "xls"):
        return jsonify({"error": f"Preview isn't supported for .{ext} files."}), 400

    work_dir = tempfile.mkdtemp(prefix=f"mc_preview_{uuid.uuid4().hex[:8]}_")
    try:
        safe_name, _ext = safe_upload_filename(file.filename)
        src_path = os.path.join(work_dir, safe_name)
        file.save(src_path)

        try:
            result_path = preview_file(src_path, work_dir, ext)
        except ConversionError as e:
            return jsonify({"error": str(e)}), 422
        except FileNotFoundError as e:
            return jsonify({"error": f"Required preview tool missing on server: {e}"}), 500

        return send_file(result_path, mimetype="application/pdf", as_attachment=False)
    except Exception as e:
        logger.error(f"Preview generation failed unexpectedly: {e}", exc_info=True)
        return jsonify({"error": "Couldn't generate a preview. Please try again."}), 500
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


@app.route("/api/convert-batch", methods=["POST", "OPTIONS"])
@limiter.limit(BATCH_LIMIT, key_func=_account_or_ip_key)
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

            src_path = os.path.join(work_dir, safe_upload_filename(filename)[0])
            file.save(src_path)
            try:
                result_path = convert(src_path, from_fmt, to_fmt, work_dir, style=style)
                base_name = safe_download_name(filename).rsplit(".", 1)[0]
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
        logger.error(f"Batch conversion failed unexpectedly: {e}", exc_info=True)
        return jsonify({"error": "Batch conversion failed. Please try again."}), 500
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


@app.route("/api/pdf/edit", methods=["POST", "OPTIONS"])
@limiter.limit(HEAVY_LIMIT, key_func=_account_or_ip_key)
@auth_required
def pdf_edit_endpoint():
    """Runs a chain of PDF edit operations (merge, reorder, delete/
    extract pages, rotate, page numbers, watermark, password protect/
    remove, compress) in one pass. 'file' is the primary PDF;
    'operations' is a JSON-encoded list (see apply_pdf_operations'
    docstring in converters.py for the schema); any additional files
    a merge_append step references are uploaded under 'extra_files'
    (repeatable) and looked up by their own original filename, so the
    operations JSON can reference "files": ["worksheet2.pdf"] directly
    rather than needing a separately-assigned key."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "Empty filename"}), 400
    ext = file.filename.rsplit(".", 1)[-1].lower() if "." in file.filename else ""
    if ext != "pdf":
        return jsonify({"error": "The primary file must be a .pdf"}), 400

    try:
        operations = json.loads(request.form.get("operations") or "[]")
    except json.JSONDecodeError:
        return jsonify({"error": "'operations' must be valid JSON"}), 400
    if not isinstance(operations, list):
        return jsonify({"error": "'operations' must be a JSON list"}), 400

    try:
        _log_conversion_and_consume(request.current_user, "pdf", "pdf-edit")
    except ConversionError as e:
        return jsonify({"error": str(e)}), 402

    work_dir = tempfile.mkdtemp(prefix=f"mc_pdfedit_{uuid.uuid4().hex[:8]}_")
    try:
        safe_name, _ext = safe_upload_filename(file.filename)
        src_path = os.path.join(work_dir, safe_name)
        file.save(src_path)

        extra_files = {}
        for extra in request.files.getlist("extra_files"):
            if extra.filename:
                extra_safe_name, _ext = safe_upload_filename(extra.filename)
                extra_path = os.path.join(work_dir, extra_safe_name)
                extra.save(extra_path)
                # Keyed by the *original* filename on purpose — the
                # operations JSON's merge_append step references files
                # by that name, and a dict key is just an in-memory
                # string with no filesystem risk; only the disk write
                # location (the path above) needed to be randomized.
                extra_files[extra.filename] = extra_path

        try:
            result_path = apply_pdf_operations(src_path, operations, work_dir, extra_files)
        except ConversionError as e:
            return jsonify({"error": str(e)}), 422

        base_name = safe_download_name(file.filename).rsplit(".", 1)[0]
        return send_file(
            result_path, mimetype="application/pdf", as_attachment=True,
            download_name=f"{base_name}_edited.pdf",
        )
    except Exception as e:
        logger.error(f"PDF edit failed unexpectedly: {e}", exc_info=True)
        return jsonify({"error": "PDF edit failed. Please try again."}), 500
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


@app.route("/api/pdf/thumbnails", methods=["POST", "OPTIONS"])
@limiter.limit(HEAVY_LIMIT, key_func=_account_or_ip_key)
@auth_required
def pdf_thumbnails_endpoint():
    """Returns a base64-encoded JPEG thumbnail per page, for a
    page-picker UI (select pages to delete/extract, drag to reorder)
    — base64 JSON rather than a zip, since the frontend can drop each
    one straight into an <img src="data:..."> with no unzip step.
    Read-only preview, so this doesn't consume conversion quota the
    way an actual edit does."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "Empty filename"}), 400

    work_dir = tempfile.mkdtemp(prefix=f"mc_pdfthumb_{uuid.uuid4().hex[:8]}_")
    try:
        src_path = os.path.join(work_dir, safe_upload_filename(file.filename)[0])
        file.save(src_path)
        try:
            thumb_paths = pdf_get_page_thumbnails(src_path, work_dir)
        except ConversionError as e:
            return jsonify({"error": str(e)}), 422

        thumbnails = []
        for p in thumb_paths:
            with open(p, "rb") as f:
                thumbnails.append(base64.b64encode(f.read()).decode("ascii"))
        return jsonify({"page_count": len(thumbnails), "thumbnails": thumbnails})
    except Exception as e:
        logger.error(f"Thumbnail generation failed unexpectedly: {e}", exc_info=True)
        return jsonify({"error": "Thumbnail generation failed. Please try again."}), 500
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


@app.route("/api/pdf/split", methods=["POST", "OPTIONS"])
@limiter.limit(HEAVY_LIMIT, key_func=_account_or_ip_key)
@auth_required
def pdf_split_endpoint():
    """Splits a PDF and returns the parts as a zip. With no 'ranges'
    field, splits into one file per page; with 'ranges' (a JSON list
    of page specs, e.g. ["1-3","4-6","7"]), splits into that many
    files instead."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "Empty filename"}), 400

    ranges = None
    ranges_raw = request.form.get("ranges")
    if ranges_raw:
        try:
            ranges = json.loads(ranges_raw)
            if not isinstance(ranges, list):
                raise ValueError
        except (json.JSONDecodeError, ValueError):
            return jsonify({"error": "'ranges' must be a JSON list of page specs"}), 400

    try:
        _log_conversion_and_consume(request.current_user, "pdf", "pdf-split")
    except ConversionError as e:
        return jsonify({"error": str(e)}), 402

    work_dir = tempfile.mkdtemp(prefix=f"mc_pdfsplit_{uuid.uuid4().hex[:8]}_")
    try:
        src_path = os.path.join(work_dir, safe_upload_filename(file.filename)[0])
        file.save(src_path)
        try:
            part_paths = pdf_split(src_path, work_dir, ranges=ranges)
        except ConversionError as e:
            return jsonify({"error": str(e)}), 422

        zip_path = os.path.join(work_dir, "split_pages.zip")
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for p in part_paths:
                zf.write(p, arcname=os.path.basename(p))
        return send_file(zip_path, mimetype="application/zip", as_attachment=True, download_name="split_pages.zip")
    except Exception as e:
        logger.error(f"PDF split failed unexpectedly: {e}", exc_info=True)
        return jsonify({"error": "Split failed. Please try again."}), 500
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


@app.route("/api/pdf/images-to-pdf", methods=["POST", "OPTIONS"])
@limiter.limit(HEAVY_LIMIT, key_func=_account_or_ip_key)
@auth_required
def images_to_pdf_endpoint():
    """Combines multiple uploaded images (in the order given) into one
    PDF, one image per page."""
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "No images uploaded"}), 400

    try:
        _log_conversion_and_consume(request.current_user, "images", "pdf")
    except ConversionError as e:
        return jsonify({"error": str(e)}), 402

    work_dir = tempfile.mkdtemp(prefix=f"mc_img2pdf_{uuid.uuid4().hex[:8]}_")
    try:
        image_paths = []
        display_names = []
        for f in files:
            if not f.filename:
                continue
            p = os.path.join(work_dir, safe_upload_filename(f.filename)[0])
            f.save(p)
            image_paths.append(p)
            display_names.append(safe_download_name(f.filename))
        try:
            result_path = images_to_pdf(image_paths, work_dir, display_names=display_names)
        except ConversionError as e:
            return jsonify({"error": str(e)}), 422
        return send_file(result_path, mimetype="application/pdf", as_attachment=True, download_name="images.pdf")
    except Exception as e:
        logger.error(f"images-to-pdf failed unexpectedly: {e}", exc_info=True)
        return jsonify({"error": "Combining these images into a PDF failed. Please try again."}), 500
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
        logger.error(f"Text-to-presentation failed unexpectedly: {e}", exc_info=True)
        return jsonify({"error": "Couldn't build the presentation. Please try again."}), 500
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


@app.route("/api/extract-text", methods=["POST", "OPTIONS"])
@limiter.limit(HEAVY_LIMIT, key_func=_account_or_ip_key)   # reading photos (OCR) is heavy work
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
        src_path = os.path.join(work_dir, safe_upload_filename(file.filename)[0])
        file.save(src_path)
        text = extract_text(src_path, ext)
        if len(text) > 20000:
            text = text[:20000]
        return jsonify({"text": text, "filename": safe_download_name(file.filename)})
    except ConversionError as e:
        return jsonify({"error": str(e)}), 422
    except Exception as e:
        logger.error(f"Text extraction failed unexpectedly: {e}", exc_info=True)
        return jsonify({"error": "Couldn't read this file. Please try again."}), 500
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


SUMMARIZE_SYSTEM_PROMPT = (
    "You turn dense study material into a slide-by-slide presentation plan for a tutoring "
    "app — not a flat list of bullet points, but a genuine sequence of slides, each doing one "
    "job. Return ONLY a JSON array — no markdown code fences, no preamble, no commentary "
    "before or after it, nothing but the array itself, starting with [ and ending with ].\n\n"
    "Each element is one slide, an object with exactly these keys:\n"
    '  "header": an action-oriented or descriptive headline for this one slide (e.g. '
    '"Revenue Grew 15% in Q3", not a generic label like "Financials")\n'
    '  "bullets": an array of short strings for this slide — omit or leave empty when '
    'this slide instead uses "chart" below\n'
    '  "icon": one single emoji that visually represents this slide\'s content (e.g. '
    '\U0001F4C8 for growth, \U0001F4A1 for an idea, \u2705 for a takeaway/action) — pick '
    "one that actually fits the specific content, not the same one repeatedly\n"
    '  "chart": either null, or an object {"type": "bar"|"line"|"pie", "categories": '
    '[...], "series_name": "...", "values": [...]} when — and only when — the source '
    "material contains real numeric/comparative data suited to a chart\n"
    '  "is_cta": true only on the final slide, false everywhere else\n\n'
    "Follow every one of these — each is a specific, checkable property, not a vague style "
    "goal:\n\n"
    "1. RADICAL CONTENT CONDENSATION. The 6x6 rule: at most 6 bullets per slide, at most 6 "
    "words per bullet. One idea per slide — isolate a single core message per slide rather "
    "than crowding several concepts onto one; this means splitting the material across "
    "multiple slides is expected and correct, not a fallback. Every header is an "
    "action-oriented or descriptive headline stating the actual takeaway, never a generic "
    "category label.\n"
    "2. VISUAL HIERARCHY THROUGH BOLDING. Wrap key metrics, dates, names, and other "
    "load-bearing terms in **double asterisks** so a reader can scan the slide in under "
    "three seconds and immediately see what matters — but don't bold everything; bolding "
    "every word bolds nothing.\n"
    "3. DATA BECOMES CHARTS, NOT PROSE. When the source material contains real comparative "
    'or numeric data (figures across categories, a trend over time, a breakdown of parts), '
    'represent it with the "chart" field instead of describing it in bullets — a table or '
    "paragraph of numbers in the source should become a bar/line/pie chart on the slide, not "
    "restated as text. Only use a chart when the source genuinely has this kind of data; "
    "don't invent numbers or force a chart where none fits.\n"
    "4. PROGRESSIVE DISCLOSURE. Order the slides so they flow from a high-level big-picture "
    "opening slide, through supporting detail slides, to a final takeaway/call-to-action "
    "slide (is_cta: true) — never end on a random supporting detail.\n"
    "5. CONTEXT IS PRESERVED, NEVER STRIPPED. Names, acronyms, dates, and specific figures "
    "keep their exact original meaning — condensing to 6 words a bullet doesn't mean "
    "vague-ing out a specific number or a proper noun; cut filler words, never cut the "
    "load-bearing fact itself.\n\n"
    "If the material is too short or thin to support multiple genuinely distinct slides, "
    "return fewer slides rather than padding with repetitive or trivial ones — a real "
    "2-slide summary beats a padded 6-slide one."
)


def _parse_summary_json(raw_text):
    """Parses and validates the summarize AI's slide-plan JSON, the same
    defensive-parsing approach as _parse_quiz_json: an AI's raw JSON output
    is never trusted as-is, since a per-item structural mistake (a bad
    chart shape, a missing header) should downgrade that one slide
    gracefully rather than take down the whole presentation."""
    cleaned = raw_text.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
        cleaned = re.sub(r"\s*```$", "", cleaned)
    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError:
        raise ConversionError("Couldn't generate a presentation from that material — please try again.")
    if not isinstance(data, list):
        raise ConversionError("Couldn't generate a presentation from that material — please try again.")

    slides = []
    for item in data:
        if not isinstance(item, dict):
            continue
        header = str(item.get("header") or "").strip()
        if not header:
            continue
        bullets = [str(b).strip() for b in (item.get("bullets") or []) if str(b).strip()][:6]
        icon = str(item.get("icon") or "").strip()
        chart = item.get("chart")
        if isinstance(chart, dict):
            chart_type = chart.get("type") if chart.get("type") in ("bar", "line", "pie") else None
            categories = [str(c) for c in (chart.get("categories") or [])]
            values = chart.get("values") or []
            try:
                values = [float(v) for v in values]
            except (TypeError, ValueError):
                values = []
            # A chart needs a type and matching category/value pairs to mean
            # anything — anything less isn't a valid chart, just downgrade
            # to no chart for this slide rather than fail the whole request.
            if not (chart_type and categories and values and len(categories) == len(values)):
                chart = None
            else:
                chart = {
                    "type": chart_type,
                    "categories": categories,
                    "series_name": str(chart.get("series_name") or "Value"),
                    "values": values,
                }
        else:
            chart = None
        if not bullets and not chart:
            continue  # a slide with neither content type is empty; skip it
        slides.append({
            "header": header,
            "bullets": bullets,
            "icon": icon,
            "chart": chart,
            "is_cta": bool(item.get("is_cta")),
        })
    if not slides:
        raise ConversionError("Couldn't generate a presentation from that material — please try again.")
    return slides


SAGE_SYSTEM_PROMPT = (
    "You are Remy, Docente's built-in AI tutor — not a narrow, single-purpose tool "
    "like the app's other AI features, but a genuine, open-ended assistant a tutor or "
    "student can bring almost anything to. Docente also has three purpose-built AI "
    "tools you can point people toward when they'd genuinely help more than a chat "
    "answer would: Smart Summarize (condenses a document into a slide deck), Practice "
    "Quiz (generates a quiz with an answer key from source material), and Evidence-"
    "Based Writing (drafts academic writing with real, verified citations). Mention one "
    "only when it's a clearly better fit for what they're actually asking — e.g. "
    "someone pasting a long document and asking for a slide deck should be pointed to "
    "Smart Summarize rather than have you attempt it inline — never as a reflexive "
    "sign-off.\n\n"
    "What you're actually for: explaining a concept at whatever depth and level "
    "actually fits the person in front of you, working through a problem step by step "
    "rather than only handing over a final answer, answering questions about material "
    "someone has pasted or uploaded, helping plan out how to study or teach a topic, "
    "and anything else a tutor or student would reasonably bring to a knowledgeable, "
    "patient assistant. You have access to web search for anything that depends on "
    "current or fast-changing information — use it rather than guessing when it "
    "matters, and don't reach for it for timeless, well-established material you "
    "already know well.\n\n"
    "Be direct and genuinely useful over performatively warm: skip preamble, answer the "
    "actual question first, and match the length of your response to what the question "
    "actually needs — a quick factual question earns a few sentences, a request to "
    "work through a proof or a step-by-step problem earns the space that takes. When "
    "someone's question is ambiguous, make the most reasonable assumption and say so "
    "briefly rather than stalling on a clarifying question."
)


@app.route("/api/sage", methods=["POST", "OPTIONS"])
@auth_required
def sage_endpoint():
    data = request.get_json(silent=True) or {}
    message = (data.get("message") or "").strip()
    if not message:
        return jsonify({"error": "No message provided"}), 400
    if len(message) > 20000:
        return jsonify({"error": "That's too long for one message (20,000 character limit)"}), 400

    attachment = (data.get("attachment") or "").strip()
    if len(attachment) > 20000:
        return jsonify({"error": "That attached document is too long (20,000 character limit)"}), 400

    # Real multi-turn history, capped the same way /api/write already
    # caps its own — bounded so a very long-running conversation can't
    # grow the request without limit, not because 8 turns is some ideal
    # number, just a proven, already-working budget elsewhere in this
    # file.
    raw_history = data.get("history") or []
    turns = []
    if isinstance(raw_history, list):
        for turn in raw_history[-8:]:
            if not isinstance(turn, dict):
                continue
            role = turn.get("role")
            content = (turn.get("content") or "").strip()[:4000]
            if role in ("user", "assistant") and content:
                turns.append({"role": role, "content": content})

    user_content = message
    if attachment:
        user_content = f"[Attached document]\n{attachment}\n\n[Message]\n{message}"
    turns.append({"role": "user", "content": user_content})

    # Checked and consumed before the expensive call below, not after —
    # the whole point is to never spend the API cost on a request that
    # was already over its limit.
    try:
        _log_sage_message_and_consume(request.current_user)
    except ConversionError as e:
        return jsonify({"error": str(e)}), 429

    try:
        result = call_claude(
            system_prompt=SAGE_SYSTEM_PROMPT,
            messages=turns,
            max_tokens=4000,
            use_search=True,
            model=ANTHROPIC_MODEL_STRONG,
        )
    except Exception as e:
        # The message was counted above, before the AI call — but no reply
        # arrived, so give it back. Otherwise a timeout or AI outage quietly
        # spent one of a free user's 10 monthly messages, and tapping Retry
        # spent another.
        _refund_sage_message(request.current_user)
        if isinstance(e, ConversionError):
            return jsonify({"error": str(e)}), 502
        raise
    return jsonify({"reply": result})


@app.route("/api/summarize", methods=["POST", "OPTIONS"])
@limiter.limit(AI_LIMIT, key_func=_account_or_ip_key)
@auth_required
def summarize_endpoint():
    data = request.get_json(silent=True) or {}
    text = (data.get("text") or "").strip()
    if not text:
        return jsonify({"error": "No text provided"}), 400
    if len(text) > 20000:
        return jsonify({"error": "Text too long (20,000 character limit)"}), 400

    slide_count = data.get("slide_count")
    if slide_count is not None:
        try:
            slide_count = int(slide_count)
        except (TypeError, ValueError):
            slide_count = None
        else:
            slide_count = max(3, min(20, slide_count))

    # "auto" (slide_count is None) is a real, first-class default here,
    # not a missing value to fall back from: letting the AI decide slide
    # count from how much the material actually supports is frequently
    # the better choice than forcing a number, the same tension already
    # handled in the quiz generator's own "don't pad thin material"
    # instruction — this just makes that choice explicit and available
    # to the user rather than only ever inferred.
    if slide_count:
        user_message = (
            f"Aim for approximately {slide_count} slides — match this closely when the "
            f"material genuinely supports that many genuinely distinct slides, but return "
            f"fewer rather than padding with repetitive or trivial ones if it doesn't. "
            f"Material:\n\n{text}"
        )
    else:
        user_message = text

    # Scales with the requested count rather than a flat number, the
    # same fix already applied to the quiz generator's own token budget
    # and for the identical reason: a fixed budget sized for a small
    # request silently truncates a larger one into invalid, unparseable
    # JSON rather than a clean error, which is exactly the failure mode
    # a "give me 20 slides" option would otherwise hit. 300 tokens/slide
    # is deliberately generous — a slide's JSON includes its header,
    # up to 6 bullets, an icon, and sometimes chart data, and richer
    # bullets (this prompt's whole point) run longer than bare-minimum
    # ones would. The "auto" case (no explicit count) still gets more
    # than the old flat 3000, since a long, rich source document can
    # reasonably warrant many slides even without an explicit request.
    max_tokens = min(8000, max(4000, (slide_count or 12) * 300 + 500))

    try:
        result = call_claude(
            system_prompt=SUMMARIZE_SYSTEM_PROMPT,
            user_message=user_message,
            max_tokens=max_tokens,
        )
        slides = _parse_summary_json(result)
        return jsonify({"slides": slides})
    except ConversionError as e:
        return jsonify({"error": str(e)}), 502


@app.route("/api/summary-to-pptx", methods=["POST", "OPTIONS"])
@auth_required
def summary_to_pptx_endpoint():
    """Builds the actual presentation from an already-generated slide
    plan (the response of /api/summarize above), rather than the
    previous "Export to PowerPoint" behavior of silently re-running the
    original, un-summarized text through the unrelated text_to_pptx —
    confirmed directly that the AI's summary was, until now, never
    actually used for the export a user downloads, only shown inline
    in the app and then discarded the moment they clicked export."""
    data = request.get_json(silent=True) or {}
    slides = data.get("slides")
    if not isinstance(slides, list) or not slides:
        return jsonify({"error": "No slide plan provided."}), 400
    style = data.get("style") if data.get("style") in ("visual", "plain") else "visual"

    work_dir = tempfile.mkdtemp(prefix=f"mc_sum_{uuid.uuid4().hex[:8]}_")
    try:
        result_path = summary_slides_to_pptx(slides, work_dir, style=style)
        return send_file(
            result_path,
            mimetype=MIME_TYPES["pptx"],
            as_attachment=True,
            download_name="Presentation.pptx",
        )
    except ConversionError as e:
        return jsonify({"error": str(e)}), 422
    except Exception as e:
        logger.error(f"Summary-to-presentation failed unexpectedly: {e}", exc_info=True)
        return jsonify({"error": "Couldn't build the presentation. Please try again."}), 500
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


QUIZ_SYSTEM_PROMPT = (
    "You write practice quiz questions for a tutoring app, from whatever study material the "
    "user gives you. Return ONLY a JSON array — no markdown code fences, no preamble, no "
    "commentary before or after it, nothing but the array itself, starting with [ and ending "
    "with ].\n\n"
    "Each element is an object with exactly these keys:\n"
    '  "type": either "multiple_choice" or "short_answer"\n'
    '  "question": the question text, self-contained and answerable from the material given\n'
    '  "options": for multiple_choice ONLY — an array of exactly 4 answer choices, '
    "in any order, as plain strings with no letter prefixes (the document adds A/B/C/D itself); "
    "omit this key entirely for short_answer\n"
    '  "correct_answer": for multiple_choice, the exact text of the correct option, copied '
    "verbatim from the options array; for short_answer, a concise model answer\n"
    '  "explanation": one short sentence on why that answer is correct — this is for an '
    "answer key a tutor reviews with a student, not for the student's copy of the quiz\n\n"
    "Write a mix of multiple_choice and short_answer unless told otherwise. Base every "
    "question strictly on the material provided — never introduce facts the material doesn't "
    "support or contradict it. If the material is too short or thin to support the requested "
    "number of distinct questions, return as many genuinely distinct ones as it actually "
    "supports rather than padding with repetitive or trivial ones.\n\n"
    "The multiple-choice questions you write must avoid these specific, well-documented "
    "failure modes of AI-generated MCQs — each one is a concrete, checkable property, not a "
    "vague quality goal:\n\n"
    "1. NOT JUST RECALL. Don't make every question 'what is the definition of X' or 'which "
    "term means Y'. For most questions, require applying a concept to a new situation, "
    "comparing two ideas from the material, interpreting what a result or example implies, or "
    "identifying why a plausible-sounding claim is actually wrong. A good rule of thumb: if the "
    "question could be answered by finding one sentence in the material and matching its "
    "wording to an option, rewrite it so the student has to reason with the material instead of "
    "just locating it.\n"
    "2. DISTRACTORS MUST BE GENUINELY TEMPTING, NOT OBVIOUSLY WRONG. Never write a throwaway, "
    "silly, or extreme wrong option (nothing a student could eliminate without knowing the "
    "material at all). Each of the 3 incorrect options should be something a student who "
    "half-understood the material, or who holds a common misconception about it, would "
    "plausibly pick. Base wrong options on: a real misconception about the topic, a mixed-up "
    "adjacent fact from elsewhere in the material, a common calculation or reasoning error, or "
    "a statement that's true in general but wrong in this specific context.\n"
    "3. NO STRUCTURAL GIVEAWAYS. All 4 options must be similar to each other in length, "
    "grammatical form, and level of detail — never make the correct answer the longest, the "
    "most qualified/hedged, or the only one written as a complete sentence when the others are "
    "fragments. A student should not be able to spot the answer by how it's written.\n"
    "4. NO KEYWORD ECHOING. The correct option should not simply repeat a distinctive word or "
    "phrase straight from the question stem — that's a giveaway that requires no actual "
    "understanding. Paraphrase the correct answer using different wording than the stem where "
    "the material allows it, and don't compensate by stuffing that same keyword into the "
    "distractors either (vary all 4 options' wording naturally).\n"
    "5. WRITE WITH A PURPOSE, NOT TO FILL A QUOTA. Before writing each question, identify one "
    "specific fact, relationship, or skill from the material it's meant to test — don't "
    "generate generic filler questions just to hit the requested count. Prefer testing the "
    "ideas most central to the material, and points students commonly get confused about, over "
    "minor or incidental details.\n"
    "6. VARY DIFFICULTY ACROSS THE SET. Don't make every question the same difficulty — include "
    "a genuine range from straightforward to challenging so the quiz can distinguish partial "
    "understanding from mastery, not just pass/fail on the same level of question repeated.\n\n"
    "You don't need to track or balance which letter position (A/B/C/D) ends up correct — "
    "options are shuffled into random order after you write them, so putting the correct "
    "answer in a natural, sensible place in your own options array (rather than always first "
    "or always last) is all that's needed here.\n\n"
    "Most of this — testing real understanding rather than recall, having a clear reason for "
    "each question, and varying difficulty across the set — matters just as much for "
    "short_answer questions, even though the specific mechanics above (distractors, option "
    "length, letter position) are multiple-choice-only concepts that don't apply to them. Two "
    "more short_answer-specific points: don't phrase the question so its own wording already "
    "hands the student the answer (a question that echoes back most of the model answer's key "
    "term isn't testing anything); and write a model answer that actually matches what the "
    "question asks for — if it asks the student to explain why or how something happens, the "
    "model answer should give that reasoning, not just name the term, since a tutor grading "
    "against it needs to see what a complete answer actually looks like."
)


def _parse_quiz_json(raw_text):
    """Claude is instructed to return a bare JSON array, but models
    sometimes wrap it in a markdown code fence despite that instruction
    — stripped defensively here rather than trusting the instruction
    held. Raises ConversionError with a clear, non-technical message on
    anything that still doesn't parse, rather than letting a JSON
    error surface as a raw 500."""
    text = raw_text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```[a-zA-Z]*\n?", "", text)
        text = re.sub(r"\n?```$", "", text)
        text = text.strip()
    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        raise ConversionError("Couldn't generate a quiz from that material — please try again.")
    if not isinstance(data, list) or not data:
        raise ConversionError("Couldn't generate a quiz from that material — please try again.")

    questions = []
    for item in data:
        if not isinstance(item, dict):
            continue
        q_type = item.get("type") if item.get("type") in ("multiple_choice", "short_answer") else "short_answer"
        q_text = str(item.get("question") or "").strip()
        if not q_text:
            continue
        options = item.get("options") if isinstance(item.get("options"), list) else None
        if q_type == "multiple_choice" and (not options or len(options) < 2):
            q_type = "short_answer"  # a malformed MC question still becomes a usable question, not a dropped one
            options = None
        correct_answer = str(item.get("correct_answer") or "").strip()
        if q_type == "multiple_choice" and options and correct_answer:
            # Confirmed directly this is a real failure mode, not a
            # theoretical one: a model can name a correct_answer that
            # doesn't actually match any of its own listed options
            # (case/whitespace variation, or an outright inconsistency
            # like answering "Paris" to a question whose choices are
            # London/Berlin/Madrid/Rome). A student checking the answer
            # key against a quiz where the right answer isn't even one
            # of the choices is a worse experience than a short-answer
            # question, so this downgrades the same way too-few-options
            # already does above rather than shipping a quiz that
            # can't actually be answered as multiple-choice.
            if not any(correct_answer.lower() == str(o).strip().lower() for o in options):
                q_type = "short_answer"
                options = None
        final_options = [str(o) for o in options][:8] if options else None
        if final_options:
            # A model's own habits about *where* it places the correct
            # option (first, or last, after the distractors) would
            # otherwise carry straight through to the printed quiz —
            # exactly the "predictable letter position" weakness a
            # prompt instruction alone can't reliably fix, since an
            # LLM asked to "randomize" its own output order isn't
            # actually drawing from a uniform distribution. A real
            # shuffle here is the only way to guarantee it. correct_answer
            # is stored and matched by its text, not by position (the
            # answer key prints that text directly, never a letter), so
            # reordering this list can never desynchronize the two.
            random.shuffle(final_options)
        questions.append({
            "type": q_type,
            "question": q_text,
            "options": final_options,
            "correct_answer": correct_answer,
            "explanation": str(item.get("explanation") or "").strip(),
        })
    if not questions:
        raise ConversionError("Couldn't generate a quiz from that material — please try again.")
    return questions


@app.route("/api/quiz/generate", methods=["POST", "OPTIONS"])
@limiter.limit(AI_LIMIT, key_func=_account_or_ip_key)
@auth_required
def quiz_generate_endpoint():
    data = request.get_json(silent=True) or {}
    text = (data.get("text") or "").strip()
    title = (data.get("title") or "Practice Quiz").strip()[:150]
    if not text:
        return jsonify({"error": "No study material provided"}), 400
    if len(text) > 30000:
        return jsonify({"error": "Text too long (30,000 character limit)"}), 400

    num_questions = data.get("num_questions", 8)
    try:
        num_questions = int(num_questions)
    except (TypeError, ValueError):
        num_questions = 8
    num_questions = max(3, min(200, num_questions))

    question_style = data.get("question_style") or "a mix of multiple_choice and short_answer"
    if question_style == "multiple_choice":
        question_style = "only multiple_choice"
    elif question_style == "short_answer":
        question_style = "only short_answer"
    else:
        question_style = "a mix of multiple_choice and short_answer"

    user_message = (
        f"Generate exactly {num_questions} questions ({question_style}) from this material:\n\n{text}"
    )

    # Scales with num_questions rather than a flat value: confirmed
    # against the actual JSON shape (question + 4 options + answer +
    # explanation per item) that a fixed 4000-token budget — sized for
    # the old 20-question cap — would silently truncate a large
    # request into invalid, unparseable JSON well before reaching 200
    # questions. 300 tokens/question is a deliberately generous
    # per-question estimate, since the richer, more carefully-reasoned
    # distractors this prompt now asks for run longer than the
    # shorter, weaker ones a bare-minimum prompt would produce.
    # Capped at 64000 to stay within every current model's max-output
    # ceiling (verified as low as 64K tokens for some tiers) regardless
    # of which one handles this request.
    max_tokens = min(64000, max(4000, num_questions * 300 + 500))

    try:
        result = call_claude(
            system_prompt=QUIZ_SYSTEM_PROMPT,
            user_message=user_message,
            max_tokens=max_tokens,
        )
        questions = _parse_quiz_json(result)
        return jsonify({"questions": questions, "title": title})
    except ConversionError as e:
        return jsonify({"error": str(e)}), 502


@app.route("/api/quiz/export-docx", methods=["POST", "OPTIONS"])
@auth_required
def quiz_export_docx_endpoint():
    """Builds the actual .docx from an already-generated quiz (the
    response of /api/quiz/generate above) — the same split already
    proven for Smart Summarize (/api/summarize generates the JSON plan,
    /api/summary-to-pptx builds the file from it), applied here for the
    identical reason: generating and exporting were one inseparable
    step before, downloading a quiz the user had never actually seen,
    with no way to review it first."""
    data = request.get_json(silent=True) or {}
    questions = data.get("questions")
    title = (data.get("title") or "Practice Quiz").strip()[:150]
    if not isinstance(questions, list) or not questions:
        return jsonify({"error": "No quiz to export."}), 400

    work_dir = tempfile.mkdtemp(prefix=f"mc_quiz_{uuid.uuid4().hex[:8]}_")
    try:
        try:
            result_path = quiz_to_docx(title, questions, work_dir)
        except ConversionError as e:
            return jsonify({"error": str(e)}), 422
        safe_title = safe_download_name(title).replace(" ", "_") or "Practice_Quiz"
        return send_file(
            result_path, mimetype=MIME_TYPES["docx"], as_attachment=True,
            download_name=f"{safe_title}.docx",
        )
    except Exception as e:
        logger.error(f"Quiz export failed unexpectedly: {e}", exc_info=True)
        return jsonify({"error": "Couldn't build the quiz document. Please try again."}), 500
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


ROLE_DESCRIPTION = (
    "You're Docente's academic writing partner for a student or researcher — help with "
    "whatever the work actually needs: drafting a proposal, lit review, or full paper; "
    "tightening grammar and tone in something they've pasted in; formatting APA citations; "
    "talking through gaps in an argument or methodology; explaining a concept or a study in "
    "plain language; advising on study design, sampling, stats, or ethics; interpreting a "
    "results table; or building a questionnaire or outline. Use your own judgment for shape "
    "and length — flowing prose for a draft, a list for a questionnaire, a short critique for "
    "spotting gaps — rather than defaulting to one format. Talk to the student like you "
    "normally would: warm, direct, genuinely useful, not scripted. When you're drafting the "
    "actual academic content itself, though, write that in standard formal register — the "
    "warmth is in how you talk to them, not inside the paper you hand them."
)

CLARIFY_MARKER = "[CLARIFY]"
CLARIFY_INSTRUCTION = (
    "Interpret the request generously and do your best with it even if it is unclear, "
    "incomplete, oddly phrased, or missing details a student would ideally have given — use "
    "your own general knowledge and judgment to fill reasonable gaps rather than requiring a "
    "perfectly-worded request. If a specific missing detail would meaningfully change what "
    "you write, briefly state the assumption you're making in one sentence at the very start "
    "of your reply, then continue immediately with the full piece written on that assumption "
    "— never stop at just the assumption. Reserve asking a clarifying question INSTEAD of "
    "writing for the rare case where the request is so minimal, contentless, or contradictory "
    "that no reasonable topic can be identified at all — not merely because it's broad, "
    f"informal, or underspecified. In that rare case only, start your entire reply with the "
    f"exact marker {CLARIFY_MARKER} and nothing else before it."
)


def split_clarify(text):
    """Detects the model's clarify marker at the very start of its reply and strips
    it, so a genuinely ambiguous request surfaces as a short question instead of a
    full piece of writing being generated from a guess."""
    stripped = text.strip()
    if stripped.startswith(CLARIFY_MARKER):
        return True, stripped[len(CLARIFY_MARKER):].strip()
    return False, stripped


THINK_START = "[THINKING]"
THINK_END = "[/THINKING]"
THINKING_INSTRUCTION = (
    "Before your main reply, briefly think out loud about how you're approaching this "
    "request — your read on what's being asked, which sources or angle you're leaning on, "
    "and why. Keep it a few sentences, in your own natural voice. Wrap it in the exact "
    f"markers {THINK_START} and {THINK_END} with nothing else on those lines, then "
    "immediately after the closing marker write your actual answer in full — the thinking "
    "is shown to the student separately and is never a substitute for the complete answer."
)


def split_thinking(text):
    """Strips a leading [THINKING]...[/THINKING] block from the reply and returns
    (thinking_text_or_None, remaining_text) — keeps the model's own reasoning
    visually and functionally separate from the actual answer."""
    stripped = text.strip()
    if stripped.startswith(THINK_START):
        end_idx = stripped.find(THINK_END)
        if end_idx != -1:
            thinking = stripped[len(THINK_START):end_idx].strip()
            rest = stripped[end_idx + len(THINK_END):].strip()
            return (thinking or None), rest
    return None, text


MEMORY_START = "[MEMORY]"
MEMORY_END = "[/MEMORY]"
MEMORY_INSTRUCTION = (
    "If, from this exchange, you learn something durable and genuinely reusable about this "
    "student's ongoing work — their field of study, a specific project or thesis they're on, "
    "a recurring topic, a citation-style or formatting preference — add ONE short sentence "
    "noting it, wrapped in the exact markers "
    f"{MEMORY_START} and {MEMORY_END}, placed after your full answer and any references. "
    "Only include this when you've learned something genuinely new and worth remembering for "
    "next time — omit it entirely otherwise, and never repeat something you already noted "
    "before. This is saved privately for future context and never shown to the student."
)


def split_memory_note(text):
    """Strips a trailing [MEMORY]...[/MEMORY] note from the reply, wherever it
    appears, and returns (note_text_or_None, remaining_text)."""
    start = text.find(MEMORY_START)
    if start == -1:
        return None, text
    end = text.find(MEMORY_END, start)
    if end == -1:
        return None, text
    note = text[start + len(MEMORY_START):end].strip()
    remaining = (text[:start] + text[end + len(MEMORY_END):]).strip()
    return (note or None), remaining


@app.route("/api/write", methods=["POST", "OPTIONS"])
@limiter.limit(AI_LIMIT, key_func=_account_or_ip_key)
@auth_required
def write_endpoint():
    data = request.get_json(silent=True) or {}
    topic = (data.get("topic") or "").strip()
    if not topic:
        return jsonify({"error": "No topic provided"}), 400
    if len(topic) > 20000:
        return jsonify({"error": "That's too long for one message (20,000 character limit) — try splitting it into sections"}), 400

    # An uploaded document's text, if any — kept separate from the typed
    # command so it never pollutes the search query, and so the frontend can
    # display just the command in the chat transcript.
    attachment = (data.get("attachment") or "").strip()
    if len(attachment) > 20000:
        return jsonify({"error": "That attached document is too long (20,000 character limit)"}), 400

    # Optional conversation history for the chat UI — absent/empty means the
    # existing one-shot topic/outline behavior below, unchanged.
    raw_history = data.get("history") or []
    history = []
    if isinstance(raw_history, list):
        for turn in raw_history[-8:]:
            if not isinstance(turn, dict):
                continue
            role = turn.get("role")
            content = (turn.get("content") or "").strip()[:2000]
            if role in ("user", "assistant") and content:
                history.append({"role": role, "content": content})

    outline = parse_outline(topic)

    # Continuation of a truncated previous answer — a distinct, simpler path:
    # no clarify/thinking/outline logic, just pick the citation-bearing
    # writing back up exactly where it stopped.
    if data.get("continue_previous"):
        previous_text = (data.get("previous_text") or "").strip()
        if not previous_text:
            return jsonify({"error": "Nothing to continue"}), 400
        try:
            cont_sources, cont_diag = search_for_sources(topic)
            cont_sources = enrich_sources_for_apa(cont_sources)
            if cont_sources:
                cont_sources_block = build_sources_block(cont_sources)
                cont_system_prompt = (
                    f"{ROLE_DESCRIPTION} You are continuing your own previous response, which "
                    "was cut off partway through. Below is a list of real sources for this "
                    "topic, each labeled with its exact APA in-text citation, e.g. "
                    "(Smith, 2023), listed most-recent-first. Continue writing directly from "
                    "where the previous text left off — do not repeat, restate, or summarize "
                    "anything already written. Keep citing EVERY sentence that makes a claim "
                    "using the EXACT citation key next to its supporting source, copied "
                    "exactly, preferring the most recent source when more than one could "
                    "support a claim; never invent a citation. If a sentence can't be tied to "
                    "a source, leave that claim out rather than writing it uncited. Do not "
                    "write a references list yourself — it is generated separately.\n\n"
                    f"PREVIOUS TEXT (do not repeat any of this):\n{previous_text}\n\nSOURCES:\n{cont_sources_block}"
                )
            else:
                cont_system_prompt = (
                    f"{ROLE_DESCRIPTION} You are continuing your own previous response, which "
                    "was cut off partway through. Continue writing directly from where the "
                    "previous text left off — do not repeat, restate, or summarize anything "
                    "already written. No verified sources are available for this topic, so "
                    "continue from your own well-informed general knowledge, with no "
                    "citation markers or invented sources.\n\n"
                    f"PREVIOUS TEXT (do not repeat any of this):\n{previous_text}"
                )
            cont_result = call_claude(
                system_prompt=cont_system_prompt,
                user_message=topic,
                max_tokens=4000,
                use_search=False,
                return_meta=True,
            )
            cont_text = cont_result["text"].strip()
            if not cont_text:
                raise ConversionError("Couldn't continue — try again")
            return jsonify({
                "type": "result",
                "text": cont_text,
                "references": build_references_list(cont_sources),
                "truncated": cont_result["stop_reason"] == "max_tokens",
            })
        except ConversionError as e:
            return jsonify({"error": str(e)}), 502

    try:
        if len(outline) >= 2:
            # Outline mode: e.g. "5.1 Discussion / 5.1.1 General Awareness / 5.1.2 ..."
            # Cap section count so one request can't run unbounded.
            outline = outline[:14]
            combined_headings = "; ".join(s["heading"] for s in outline)

            # One shared search across the whole outline — a bigger, single pool
            # of real sources every section draws from, instead of one search per
            # section (slower and more likely to fragment/duplicate sources).
            sources, sources_diag = search_for_sources(combined_headings, max_results=8)
            sources = enrich_sources_for_apa(sources)

            if sources:
                sources_block = build_sources_block(sources)
                base_prompt = (
                    "You are writing one subsection of a longer academic piece on "
                    f"\"{combined_headings}\". Write ONLY the subsection given below as the "
                    "user message — 100-160 words, formal academic prose. Below is a list of "
                    "real sources relevant to the overall piece, each labeled with its exact "
                    "APA in-text citation, e.g. (Smith, 2023), listed most-recent-first. Write "
                    "the way you actually would — ground every factual or evidentiary claim in "
                    "one of these sources, using the EXACT key next to it, copied exactly — "
                    "never alter an author name or year, and never invent a citation not in "
                    "this list. Let the writing read naturally: a claim gets a citation, a "
                    "transition or synthesis sentence doesn't need one bolted on for its own "
                    "sake. When more than one source could support the same claim, prefer the "
                    "most recent one. If a claim can't be tied to one of the given sources, "
                    "leave it out rather than writing it uncited or attaching a fabricated "
                    "citation — keep the subsection only as long as what the sources genuinely "
                    "support. Never write a placeholder in place of a missing citation ([Author, "
                    "Year], [Citation needed], or similar) — if it's not covered, just leave the "
                    "claim out. No heading, no reference list — just the paragraph.\n\nSOURCES:\n"
                    f"{sources_block}"
                )
            else:
                base_prompt = (
                    "You are writing one subsection of a longer academic piece on "
                    f"\"{combined_headings}\". Write ONLY the subsection given below as the "
                    "user message — 100-160 words, formal academic prose. No verified sources "
                    "were found for this topic, so do not invent any citations, authors, or "
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
                "type": "result",
                "sections": [
                    {"number": s["number"], "heading": s["heading"], "level": s["level"], "text": s["text"]}
                    for s in outline
                ],
                "references": build_references_list(sources),
            })

        # Chat mode (history present) and simple mode (first message, no history
        # yet) share the same citation and clarify-guardrail logic below — they
        # differ only in whether prior turns are folded into the prompt.
        user_id = request.current_user["id"] if getattr(request, "current_user", None) else None
        memory_notes = get_user_memory(user_id)
        role_with_memory = ROLE_DESCRIPTION + (
            " Here's what you've picked up about this student from past work together — use "
            "it naturally where it's actually relevant, don't just recite it back: "
            + "; ".join(memory_notes)
            if memory_notes else ""
        )

        if history:
            history_block = "\n".join(
                ("Student: " if t["role"] == "user" else "You: ") + t["content"]
                for t in history
            )
            search_query = topic
            if len(topic) < 15:
                prior_user = next((t["content"] for t in reversed(history) if t["role"] == "user"), "")
                if prior_user:
                    search_query = f"{prior_user} {topic}"
            raw_sources, search_diag = search_for_sources(search_query)
            sources = enrich_sources_for_apa(raw_sources)

            if sources:
                sources_block = build_sources_block(sources)
                system_prompt = (
                    f"{role_with_memory} {THINKING_INSTRUCTION} {MEMORY_INSTRUCTION} {CLARIFY_INSTRUCTION} You're in an ongoing "
                    "conversation with the student — below is the conversation so far, then "
                    "a list of real sources found for their latest message, each labeled "
                    "with its exact APA in-text citation, e.g. (Smith, 2023), listed "
                    "most-recent-first. When your response is generating written academic "
                    "content (a draft, an explanation, a literature summary, an argument), "
                    "write the way you actually would — ground every factual or evidentiary "
                    "claim in one of these sources, using the EXACT citation key next to it, "
                    "copied exactly, never altering an author name or year, and never "
                    "inventing a citation not in this list. Let the writing read naturally: a "
                    "claim gets a citation, a transition or synthesis sentence doesn't need "
                    "one bolted on for its own sake. When more than one source could support "
                    "the same claim, prefer the most recent one. If a claim can't be tied to "
                    "one of the given sources, leave it out rather than writing it uncited or "
                    "attaching a fabricated citation — keep the response only as long as what "
                    "the sources genuinely support. Never write a placeholder in place of a "
                    "missing citation ([Author, Year], [Citation needed], or similar) — if "
                    "it's not covered, just leave the claim out. Citations don't apply to "
                    "tasks that aren't "
                    "generating cited content — skip them entirely for grammar editing, "
                    "building a questionnaire, or interpreting a table the student pasted "
                    "in. Write as much as the task genuinely needs — a quick fix might be a "
                    "sentence or two, a literature review draft might run several "
                    "paragraphs. Do not write a references list yourself — it is generated "
                    f"separately.\n\nCONVERSATION SO FAR:\n{history_block}\n\nSOURCES:\n{sources_block}"
                )
            else:
                system_prompt = (
                    f"{role_with_memory} {THINKING_INSTRUCTION} {MEMORY_INSTRUCTION} {CLARIFY_INSTRUCTION} Source search ran for "
                    "the student's latest message and came back with nothing usable — here's "
                    f"exactly why, verbatim: \"{search_diag}\". Write from your own "
                    "well-informed general knowledge instead — but do NOT insert any "
                    "placeholder in place of a citation, in any form: no [Author, Year], no "
                    "[Citation needed], no [source], no blank brackets, nothing standing in "
                    "for a reference that isn't there. If the topic genuinely calls for "
                    "citation-backed evidence you don't have, say so plainly to the student in "
                    "a sentence or two — and include the verbatim search-failure reason above "
                    "so they know exactly what to check next, rather than a vague 'no sources "
                    "found'. Write as much as the task genuinely needs.\n\n"
                    f"CONVERSATION SO FAR:\n{history_block}"

                )
        else:
            raw_sources2, search_diag2 = search_for_sources(topic)
            sources = enrich_sources_for_apa(raw_sources2)

            if sources:
                sources_block = build_sources_block(sources)
                system_prompt = (
                    f"{role_with_memory} {THINKING_INSTRUCTION} {MEMORY_INSTRUCTION} {CLARIFY_INSTRUCTION} Below is a list of real "
                    "sources found for this exact topic, each labeled with its exact APA "
                    "in-text citation, e.g. (Smith, 2023), listed most-recent-first. When "
                    "your response is generating written academic content (a draft, an "
                    "explanation, a literature summary, an argument), write the way you "
                    "actually would — ground every factual or evidentiary claim in one of "
                    "these sources, using the EXACT citation key next to it, copied exactly, "
                    "never altering an author name or year, and never inventing a citation "
                    "not in this list. Let the writing read naturally: a claim gets a "
                    "citation, a transition or synthesis sentence doesn't need one bolted on "
                    "for its own sake. When more than one source could support the same "
                    "claim, prefer the most recent one. If a claim can't be tied to one of "
                    "the given sources, leave it out rather than writing it uncited or "
                    "attaching a fabricated citation — keep the response only as long as "
                    "what the sources genuinely support. Never write a placeholder in place "
                    "of a missing citation ([Author, Year], [Citation needed], or similar) — "
                    "if it's not covered, just leave the claim out. Citations don't apply to "
                    "tasks that aren't generating cited content — skip them entirely for grammar "
                    "editing, building a questionnaire, or interpreting a table the student "
                    "pasted in. Write as much as the task genuinely needs. Do not write a "
                    f"references list yourself — it is generated separately.\n\nSOURCES:\n{sources_block}"
                )
            else:
                system_prompt = (
                    f"{role_with_memory} {THINKING_INSTRUCTION} {MEMORY_INSTRUCTION} {CLARIFY_INSTRUCTION} Source search ran for "
                    "this exact topic and came back with nothing usable — here's exactly why, "
                    f"verbatim: \"{search_diag2}\". Write from your own well-informed general "
                    "knowledge instead — but do NOT insert any placeholder in place of a "
                    "citation, in any form: no [Author, Year], no [Citation needed], no "
                    "[source], no blank brackets, nothing standing in for a reference that "
                    "isn't there. If the topic genuinely calls for citation-backed evidence "
                    "you don't have, say so plainly to the student in a sentence or two — and "
                    "include the verbatim search-failure reason above so they know exactly "
                    "what to check next, rather than a vague 'no sources found'. Write as "
                    "much as the task genuinely needs."
                )

        user_message = (
            f"[Attached document]\n{attachment}\n[End of attached document]\n\nStudent's message: {topic}"
            if attachment else topic
        )

        result = call_claude(
            system_prompt=system_prompt,
            user_message=user_message,
            max_tokens=4000,
            use_search=False,
            return_meta=True,
        )
        raw_text = result["text"].strip()
        truncated = result["stop_reason"] == "max_tokens"

        thinking, after_thinking = split_thinking(raw_text)
        is_clarify, text = split_clarify(after_thinking)
        memory_note, text = split_memory_note(text)
        if memory_note:
            add_user_memory(user_id, memory_note)

        if not text:
            raise ConversionError("Couldn't generate a reply — try rephrasing")
        if not is_clarify and len(text) < 15:
            raise ConversionError("Couldn't generate a reply — try rephrasing")

        if is_clarify:
            return jsonify({"type": "clarify", "text": text, "thinking": thinking})

        return jsonify({
            "type": "result",
            "text": text,
            "thinking": thinking,
            "references": build_references_list(sources),
            "truncated": truncated,
        })
    except ConversionError as e:
        return jsonify({"error": str(e)}), 502


@app.route("/api/write/export-docx", methods=["POST", "OPTIONS"])
@auth_required
def write_export_docx_endpoint():
    data = request.get_json(silent=True) or {}
    title = (data.get("title") or "Academic Response").strip()
    text = data.get("text")
    sections = data.get("sections")
    references = data.get("references") or []
    if not text and not sections:
        return jsonify({"error": "No content to export"}), 400

    work_dir = tempfile.mkdtemp(prefix=f"mc_docx_{uuid.uuid4().hex[:8]}_")
    try:
        result_path = academic_essay_to_docx(
            {"title": title, "text": text, "sections": sections, "references": references},
            work_dir,
        )
        return send_file(
            result_path,
            mimetype=MIME_TYPES["docx"],
            as_attachment=True,
            download_name="Academic_Response.docx",
        )
    except ConversionError as e:
        return jsonify({"error": str(e)}), 422
    except Exception as e:
        logger.error(f"Document export failed unexpectedly: {e}", exc_info=True)
        return jsonify({"error": "Export failed. Please try again."}), 500
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


# ------------------------------------------------------------- Slide Studio
# The in-app slide editor: the Student and Tutor templates, reading an
# existing PowerPoint into editable slides (Fix Slides), and building the
# edited slides into a finished .pptx.
@app.route("/api/slides/templates", methods=["GET", "OPTIONS"])
def slides_templates_endpoint():
    return jsonify({"templates": template_catalog()})


@app.route("/api/slides/import", methods=["POST", "OPTIONS"])
@limiter.limit(HEAVY_LIMIT, key_func=_account_or_ip_key)
@auth_required
def slides_import_endpoint():
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify({"error": "Choose a PowerPoint file first."}), 400
    ext = f.filename.rsplit(".", 1)[-1].lower() if "." in f.filename else ""
    if ext != "pptx":
        return jsonify({"error": "Choose a PowerPoint (.pptx) file — older .ppt files need saving as .pptx first."}), 400
    work_dir = tempfile.mkdtemp(prefix=f"mc_slides_{uuid.uuid4().hex[:8]}_")
    try:
        path = os.path.join(work_dir, "deck.pptx")
        f.save(path)
        deck = parse_pptx_to_deck(path)
        deck["filename"] = safe_download_name(f.filename)
        return jsonify({"deck": deck})
    except StudioError as e:
        return jsonify({"error": str(e)}), 422
    except Exception as e:
        logger.error(f"Slide import failed unexpectedly: {e}", exc_info=True)
        return jsonify({"error": "That file couldn't be read. Please try another."}), 500
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


@app.route("/api/slides/build", methods=["POST", "OPTIONS"])
@limiter.limit(HEAVY_LIMIT, key_func=_account_or_ip_key)
@auth_required
def slides_build_endpoint():
    deck = request.get_json(silent=True)
    if not isinstance(deck, dict):
        return jsonify({"error": "The slides couldn't be read — please try again."}), 400
    work_dir = tempfile.mkdtemp(prefix=f"mc_build_{uuid.uuid4().hex[:8]}_")
    try:
        path = build_template_deck(deck, work_dir, str(deck.get("filename") or "Presentation.pptx"))
        return send_file(path, mimetype=MIME_TYPES["pptx"], as_attachment=True,
                         download_name=os.path.basename(path))
    except StudioError as e:
        return jsonify({"error": str(e)}), 422
    except Exception as e:
        logger.error(f"Slide build failed unexpectedly: {e}", exc_info=True)
        return jsonify({"error": "The slides couldn't be built. Please try again."}), 500
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


init_db()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    app.run(host="0.0.0.0", port=port)
