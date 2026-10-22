import json
import logging
import base64
import hmac
import os
import hashlib
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
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
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
from converters import _soffice_convert   # the server's own LibreOffice runner
from converters import pdf_apply_plan, pdf_plan_needs_pro, pdf_to_images   # 3.1: the PDF editor
from converters import word_open, word_save   # 3.1: the Word editor
from slidestudio import build_template_deck, parse_pptx_to_deck, template_catalog, StudioError, parse_pptx_original, build_original_deck

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
def _client_ip():
    """The visitor's own address. Render's edge sets True-Client-IP to it (replacing any value a visitor sends);
    elsewhere (a computer, the tests) the connection's own address."""
    v = (request.headers.get("True-Client-IP") or request.headers.get("CF-Connecting-IP") or "").strip()
    if v and len(v) <= 45 and re.fullmatch(r"[0-9A-Fa-f:.]+", v):
        return v
    return get_remote_address()


limiter = Limiter(_client_ip, app=app, storage_uri="memory://")


def _account_or_ip_key():
    """Counts requests per signed-in account rather than per network, so a
    whole school computer lab sharing one internet connection isn't treated
    as one person. Falls back to the IP address when there's no valid login."""
    auth = request.headers.get("Authorization", "")
    if auth.startswith("Bearer "):
        uid = verify_token(auth[7:])
        if uid:
            return f"user:{uid}"
    return _client_ip()


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
logger = logging.getLogger("academic-home")
if not (os.environ.get("EMAIL_ADDRESS") and (os.environ.get("EMAIL_APP_PASSWORD") or os.environ.get("BREVO_API_KEY"))):
    logger.warning("Email: NOT set up — add EMAIL_ADDRESS and EMAIL_APP_PASSWORD on Render so new accounts can confirm their email")

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
    "https://getdocente.vercel.app",
    "https://docently.vercel.app",
    "https://masterconvert-tau.vercel.app",
}
# Additional origins can be added without a code change via this env var
# (comma-separated) — set on Render if a custom domain is added later.
ALLOWED_ORIGINS |= {o.strip() for o in os.environ.get("EXTRA_ALLOWED_ORIGINS", "").split(",") if o.strip()}
# The app's address, used in password-reset emails. getdocente.vercel.app
# is the current name; the old addresses still open the same app.
FRONTEND_URL = os.environ.get("FRONTEND_URL", "https://getdocente.vercel.app").rstrip("/")


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
# Optional instead of Gmail's password: a Brevo API key (brevo.com, free tier),
# which sends over HTTPS from the same EMAIL_ADDRESS once it's a verified sender.
BREVO_API_KEY = os.environ.get("BREVO_API_KEY", "")

# Email confirmation: every new account confirms its address with a 6-digit
# code before using the app. Accounts made before this existed count as
# confirmed. Only enforced once email sending is set up (above) — a code that
# can't be sent can't be asked for.
EMAIL_CODE_MINUTES = 15
EMAIL_CODE_MAX_TRIES = 5
EMAIL_RESEND_SECONDS = 45
EMAIL_SENDS_PER_HOUR = 6

# Free Pro for everyone until this moment (UTC = Ghana time). Set PROMO_PRO_UNTIL
# to another ISO date to move it, or to "off" to end it early.
PROMO_PRO_UNTIL = os.environ.get("PROMO_PRO_UNTIL", "2026-11-01T00:00:00+00:00")

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
    # The ids to register in App Store Connect and Play Console:
    #   Pay-per-Use — one full use of Pro: PAYPERUSE_MESSAGES Remy messages
    #   that never expire (Apple: Consumable; Google Play: one-time product,
    #   consumed after it's granted). Each store transaction is credited
    #   once only (store_purchases).
    #   Monthly / Yearly — auto-renewable subscriptions in one group
    #   ("Academic-Home Pro"); renewal and cancellation are handled by the stores.
    "com.docente.app.payperuse": "payperuse",
    "com.docente.app.pass30": "payperuse",
    "com.docente.app.pro.monthly": "monthly",
    "com.docente.app.pro.halfyear": "halfyear",   # 3.2: 6 months (both stores offer a 6-month period)
    "com.docente.app.pro.yearly": "yearly",
    # Earlier placeholder ids, still accepted.
    "com.docently.app.payperuse": "payperuse",
    "com.docently.app.monthly": "monthly",
    "com.docently.app.yearly": "yearly",
}
PASS_DAYS = 30   # kept for any account on the earlier 30-day pass
# 3.2: Pro is 1 month or 6 months ('halfyear'). 'yearly' is no longer sold, but one bought earlier still counts.
PRO_PLANS = ("monthly", "halfyear", "yearly")
_PRO_SQL = "('monthly', 'halfyear', 'yearly')"
# What Pro adds beyond Free (none of Free's own allowances change).
REMY_ATTACH_FREE, REMY_ATTACH_PRO = 20000, 60000     # characters of an attached file Remy reads (~8 / ~25 pages)
REMY_PAPERS_FREE, REMY_PAPERS_PRO = 6, 12            # research papers behind each answer
DECK_CONTENT_FREE, DECK_CONTENT_PRO = 14, 22         # content slides in a deck Remy drafts
PAYPERUSE_MESSAGES = 30   # one Pay-per-Use purchase: Remy messages that never expire

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
            # When a Pay-per-Use pass ends. NULL for Free and for store
            # subscriptions (the stores handle their renewals).
            cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS plan_expires_at TIMESTAMPTZ")
            # Pay-per-Use: Remy messages bought once, used after the monthly
            # allowance, never expiring — and every store transaction that
            # granted them, so the same purchase is never credited twice.
            cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS remy_credits INTEGER NOT NULL DEFAULT 0")
            cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS memory_enabled BOOLEAN NOT NULL DEFAULT TRUE")
            cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS plan_source TEXT")   # 'store', 'paystack' or 'bonus'
            # Accounts that already exist count as confirmed (DEFAULT TRUE); new
            # password sign-ups are created unconfirmed and confirm with a code.
            cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS email_verified BOOLEAN NOT NULL DEFAULT TRUE")
            # A new address waiting for its code (the account keeps its current one until then).
            cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS pending_email TEXT")
            # Bumped to sign the account out everywhere (password reset, an address reclaimed by its owner).
            cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS token_version INTEGER NOT NULL DEFAULT 0")
            # The address has really been shown to belong to the person: a code, Google, or a reset link.
            cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS email_proven BOOLEAN NOT NULL DEFAULT FALSE")
            # A friend's referral code given at sign-up (or at checkout), and whether
            # its bonus months have been given — once per account.
            cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS referred_code TEXT")
            cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS referral_rewarded BOOLEAN NOT NULL DEFAULT FALSE")
            # Small server-wide switches both workers read (e.g. "email is paused until …").
            cur.execute("CREATE TABLE IF NOT EXISTS app_flags (name TEXT PRIMARY KEY, until TIMESTAMPTZ)")
            cur.execute("""
                CREATE TABLE IF NOT EXISTS email_codes (
                    user_id INTEGER PRIMARY KEY REFERENCES users(id) ON DELETE CASCADE,
                    code_hash TEXT NOT NULL,
                    expires_at TIMESTAMPTZ NOT NULL,
                    attempts INTEGER NOT NULL DEFAULT 0,
                    sent_at TIMESTAMPTZ NOT NULL DEFAULT now(),
                    sends_hour INTEGER NOT NULL DEFAULT 1,
                    hour_started TIMESTAMPTZ NOT NULL DEFAULT now()
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS paystack_payments (
                    reference TEXT PRIMARY KEY,
                    user_id INTEGER REFERENCES users(id) ON DELETE SET NULL,
                    plan TEXT NOT NULL,
                    amount INTEGER NOT NULL,
                    currency TEXT NOT NULL DEFAULT 'GHS',
                    status TEXT NOT NULL DEFAULT 'pending',
                    created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
                    applied_at TIMESTAMPTZ
                )
                """
            )
            cur.execute("ALTER TABLE paystack_payments ADD COLUMN IF NOT EXISTS referral_code TEXT")
            cur.execute("ALTER TABLE paystack_payments ADD COLUMN IF NOT EXISTS referral_applied BOOLEAN NOT NULL DEFAULT FALSE")
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS store_purchases (
                    transaction_key TEXT PRIMARY KEY,
                    user_id INTEGER NOT NULL,
                    product TEXT NOT NULL,
                    created_at TIMESTAMPTZ NOT NULL DEFAULT now()
                )
            """)
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
            _meet_tables(cur)
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
                # No duplicates, and at most MEMORY_LIMIT notes per person.
                cur.execute("SELECT note FROM user_memory WHERE user_id = %s", (user_id,))
                existing = [r["note"].strip().lower() for r in cur.fetchall()]
                if note.lower() in existing or len(existing) >= MEMORY_LIMIT:
                    return
                cur.execute(
                    "INSERT INTO user_memory (user_id, note) VALUES (%s, %s)",
                    (user_id, note),
                )
        finally:
            conn.close()
    except Exception:
        pass  # memory is a nice-to-have — never let it break the actual response


MEMORY_LIMIT = 50


def _memory_on(user_row):
    """Whether this person lets Remy remember (Profile -> Remy's memory)."""
    try:
        return user_row.get("memory_enabled", True) is not False
    except Exception:
        return True


def _owns_store_plan(user_id):
    """Whether this account holds a verified App Store / Google Play
    subscription (any length of Pro)."""
    try:
        conn = get_db()
        try:
            with conn.cursor() as cur:
                cur.execute("SELECT 1 AS ok FROM store_purchases WHERE user_id = %s AND product IN " + _PRO_SQL + " LIMIT 1", (user_id,))
                return cur.fetchone() is not None
        finally:
            conn.close()
    except Exception:
        return False


def _unpaid_plan(row):
    """True for a Monthly or Yearly plan that has no end date (so it isn't a
    Paystack pass or a dated store plan) and no verified store purchase on
    record — Pro that was never paid for."""
    if row.get("plan") not in PRO_PLANS or row.get("plan_expires_at"):
        return False
    return not _owns_store_plan(row["id"])


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
# New addresses (sign-up, profile, a corrected typo) must be ones mail can really reach.
_NEW_EMAIL_RE = re.compile(r"^[A-Za-z0-9!#$%&'*+/=?^_`{|}~.-]{1,64}@(?:[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?\.)+[A-Za-z]{2,63}$")


def _json_body():
    d = request.get_json(silent=True)
    return d if isinstance(d, dict) else {}


def _text(v):
    """A text field from a request, or "" when it's missing or not text."""
    return v if isinstance(v, str) else ""


def _release_unconfirmed(cur, email, keep_id):
    """Someone has just proved this address is theirs: an unconfirmed account holding it lets it go
    (and is signed out), so it can never block the real owner."""
    cur.execute("UPDATE users SET email = 'released-' || id || '@invalid.invalid', pending_email = NULL, "
                "token_version = token_version + 1 WHERE lower(email) = lower(%s) AND email_verified = FALSE AND id <> %s",
                (email, keep_id))


def _valid_new_email(email):
    if len(email) > 254 or not _NEW_EMAIL_RE.match(email):
        return False
    local = email.split("@", 1)[0]
    return not (local.startswith(".") or local.endswith(".") or ".." in local)


def random_suffix(n=4):
    alphabet = string.ascii_uppercase + string.digits
    return "".join(secrets.choice(alphabet) for _ in range(n))


def generate_referral_code(name):
    base = re.sub(r"[^A-Za-z]", "", name)[:8].upper() or "USER"
    return f"{base}{random_suffix()}"


def generate_promo_code():
    return "".join(secrets.choice(string.ascii_uppercase + string.digits) for _ in range(10))


def make_token(user_id, version=0):
    return _serializer.dumps({"user_id": user_id, "v": int(version or 0)})


def _token_for(row):
    return make_token(row["id"], row.get("token_version") if hasattr(row, "get") else 0)


def _token_version(token):
    try:
        return int((_serializer.loads(token, max_age=TOKEN_MAX_AGE) or {}).get("v") or 0)
    except Exception:
        return 0


def verify_token(token):
    if not _serializer:
        return None
    try:
        data = _serializer.loads(token, max_age=TOKEN_MAX_AGE)
        return data.get("user_id")
    except (BadSignature, SignatureExpired):
        return None


def make_reset_token(user_id, email=None):
    return _serializer.dumps({"reset_user_id": user_id, "email": (email or "").lower()}, salt="password-reset")


def verify_reset_token(token):
    if not _serializer:
        return None
    try:
        data = _serializer.loads(token, max_age=RESET_TOKEN_MAX_AGE, salt="password-reset")
        return data.get("reset_user_id")
    except (BadSignature, SignatureExpired):
        return None


def _reset_token_email(token):
    """The address a reset link was sent to (older links don't say)."""
    try:
        return (_serializer.loads(token, max_age=RESET_TOKEN_MAX_AGE, salt="password-reset") or {}).get("email") or None
    except (BadSignature, SignatureExpired):
        return None


# Whether email can be sent right now. A refused password, or mail that keeps
# failing to get out, must never leave new people stuck behind a code that can't
# arrive — so confirmation pauses (and the log says why). The pause is kept in the
# database so both server workers agree; each reads it at most every 20 seconds.
_PAUSE_CACHE = {"at": 0.0, "until": None}
_EMAIL_FAILS = []          # (time, recipient) of recent failed sends in this worker
_PAUSE_FOREVER_MINUTES = 60 * 24 * 365 * 50


def _email_configured():
    return bool(EMAIL_ADDRESS and (EMAIL_APP_PASSWORD or BREVO_API_KEY))


def _email_paused():
    now = time.time()
    if now - _PAUSE_CACHE["at"] > 20 and DATABASE_URL:
        try:
            conn = get_db()
            try:
                with conn.cursor() as cur:
                    cur.execute("SELECT until FROM app_flags WHERE name = 'email_paused'")
                    r = cur.fetchone()
            finally:
                conn.close()
            _PAUSE_CACHE["until"] = r["until"].timestamp() if r and r["until"] else None
            _PAUSE_CACHE["at"] = now
        except Exception:
            pass   # keep the last known answer
    u = _PAUSE_CACHE["until"]
    return u is not None and now < u


def _email_ready():
    """Email can be sent: Gmail (address + app password) or Brevo (address + API key), and it isn't paused."""
    return _email_configured() and not _email_paused()


def _set_email_pause(minutes):
    """minutes > 0 pauses confirmation; 0 lifts the pause."""
    until = time.time() + minutes * 60 if minutes else None
    _PAUSE_CACHE.update(until=until, at=time.time())
    if not DATABASE_URL:
        return
    try:
        conn = get_db()
        try:
            with conn.cursor() as cur:
                if minutes:
                    cur.execute("INSERT INTO app_flags (name, until) VALUES ('email_paused', now() + make_interval(mins => %s)) "
                                "ON CONFLICT (name) DO UPDATE SET until = EXCLUDED.until", (int(minutes),))
                else:
                    cur.execute("DELETE FROM app_flags WHERE name = 'email_paused'")
        finally:
            conn.close()
    except Exception as e:
        logger.warning(f"Couldn't save the email pause: {e}")


def _email_refused(why, retry_minutes=0):
    """A refused password pauses confirmation for 30 minutes (then email is tried again);
    mail that keeps failing pauses it for 15."""
    _set_email_pause(retry_minutes or 30)
    if retry_minutes:
        logger.warning("Email: emails keep failing to send (%s). New accounts aren't asked for a code for the next %d minutes.", why, retry_minutes)
    else:
        logger.warning("Email: the email account REFUSED to sign in (%s). Check EMAIL_APP_PASSWORD (a Gmail App Password, "
                       "not the normal password) or BREVO_API_KEY on Render. New accounts aren't asked for a code for the next 30 minutes.", why)


def _send_failure_kind(err):
    """'auth' (wrong password/key), 'address' (this address can never get mail) or 'general' (mail isn't getting out)."""
    text = str(err)
    if isinstance(err, smtplib.SMTPAuthenticationError) or "answered 401" in text or "answered 403" in text:
        return "auth"
    if isinstance(err, UnicodeError):
        return "address"
    if isinstance(err, smtplib.SMTPRecipientsRefused):
        codes = [v[0] for v in (getattr(err, "recipients", None) or {}).values() if isinstance(v, tuple) and v]
        return "address" if codes and all(500 <= int(c) < 600 for c in codes) else "general"
    if "answered 400" in text and re.search(r"\bin to\b|\bto\[", text, re.I):
        return "address"
    return "general"


def _email_send_result(ok, err=None, to=None):
    """Three sends that fail for a general reason within 10 minutes pause confirmation for 15
    minutes. An address that can never receive mail doesn't count."""
    now = time.time()
    if ok:
        _EMAIL_FAILS.clear()
        if _PAUSE_CACHE["until"]:
            _set_email_pause(0)
        return
    kind = _send_failure_kind(err)
    if kind == "auth":
        return _email_refused(str(err)[:120])
    if kind == "address":
        return
    _EMAIL_FAILS[:] = [f for f in _EMAIL_FAILS if now - f[0] < 600] + [(now, (to or "").lower())]
    if len(_EMAIL_FAILS) >= 3:
        _EMAIL_FAILS.clear()
        _email_refused(f"{err.__class__.__name__}: {str(err)[:100]}", retry_minutes=15)


def _email_unreachable(err):
    """The email server can't be reached at all: rather than leave new students waiting for a code
    that can't arrive, confirmation pauses for 15 minutes and the check runs again then."""
    import threading
    _set_email_pause(15)
    logger.warning("Email: can't reach the email server (%s). New accounts aren't asked for a code for the next 15 minutes "
                   "(they confirm later); checking again then. If this keeps happening with Gmail, add BREVO_API_KEY on Render "
                   "— Brevo sends over the web instead.", err.__class__.__name__)
    t = threading.Timer(15 * 60, _check_email_login)
    t.daemon = True
    t.start()


def _claim_email_check():
    """Only one server worker signs in at start-up — two at the same moment made Gmail hang up on both."""
    if not DATABASE_URL:
        return True
    try:
        conn = get_db()
        try:
            with conn.cursor() as cur:
                cur.execute("INSERT INTO app_flags (name, until) VALUES ('email_checked', now()) "
                            "ON CONFLICT (name) DO UPDATE SET until = now() WHERE app_flags.until IS NULL OR app_flags.until < now() - interval '2 minutes' "
                            "RETURNING name")
                return cur.fetchone() is not None
        finally:
            conn.close()
    except Exception:
        return True


_GMAIL_PORT = {"port": 465}   # the Gmail door that last worked from this server


class _GmailSession:
    """A signed-in connection to Gmail: port 465 (SSL) or, if that door hangs up, port 587 (STARTTLS)."""
    def __enter__(self):
        first = _GMAIL_PORT["port"]
        last = None
        self.s = None
        for port in [first] + [p for p in (465, 587) if p != first]:
            try:
                if port == 465:
                    s = smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=15)
                else:
                    s = smtplib.SMTP("smtp.gmail.com", 587, timeout=15)
                    s.starttls()
                self.s = s
                if port != first:
                    logger.warning(f"Email: Gmail answered on port {port} (port {first} hung up) — using {port} from now on")
                _GMAIL_PORT["port"] = port
                break
            except (smtplib.SMTPServerDisconnected, smtplib.SMTPConnectError, ConnectionError, TimeoutError) as e:
                last = e
        if self.s is None:
            raise last
        try:
            self.s.login(EMAIL_ADDRESS, EMAIL_APP_PASSWORD)
        except Exception:
            self._close()
            raise
        return self.s

    def _close(self):
        try:
            self.s.quit()
        except Exception:
            try:
                self.s.close()
            except Exception:
                pass

    def __exit__(self, *a):
        self._close()
        return False


def _check_email_login(attempt=1):
    """Signs in once at start-up (sends nothing) so a wrong password shows up in the log straight away."""
    if not _email_configured():
        return
    if attempt == 1:
        time.sleep(random.uniform(0.5, 3.0))
        if not _claim_email_check():
            return
        logger.warning("Email: set up — checking that it can sign in…")
    try:
        if BREVO_API_KEY:
            r = requests.get("https://api.brevo.com/v3/account", timeout=15,
                             headers={"api-key": BREVO_API_KEY, "accept": "application/json"})
            if r.status_code in (401, 403):
                return _email_refused(f"Brevo answered {r.status_code}")
            if r.status_code >= 300:
                return
        else:
            with _GmailSession():
                pass
        _set_email_pause(0)
        logger.warning("Email: signed in OK — new accounts confirm their email with a code")
    except smtplib.SMTPAuthenticationError as e:
        _email_refused(f"Gmail said: {getattr(e, 'smtp_code', '')} wrong address or app password")
    except Exception as e:
        if attempt == 1:
            time.sleep(20)
            return _check_email_login(attempt=2)
        _email_unreachable(e)


if os.environ.get("EMAIL_LOGIN_CHECK", "on") != "off":
    import threading
    threading.Thread(target=_check_email_login, daemon=True).start()


def send_email(to_email, subject, body, html=None):
    if not _email_ready():
        raise ConversionError("Email sending isn't configured on the server yet")
    if BREVO_API_KEY:
        payload = {"sender": {"name": "Academic-Home", "email": EMAIL_ADDRESS}, "to": [{"email": to_email}],
                   "subject": subject, "textContent": body}
        if html:
            payload["htmlContent"] = html
        for attempt in (1, 2):
            try:
                r = requests.post("https://api.brevo.com/v3/smtp/email", json=payload, timeout=15,
                                  headers={"api-key": BREVO_API_KEY, "accept": "application/json", "content-type": "application/json"})
                break
            except (requests.exceptions.ConnectionError, requests.exceptions.Timeout) as e:
                if attempt == 2:
                    raise
                logger.warning(f"Email: the connection dropped ({e.__class__.__name__}) — trying once more")
                time.sleep(1.5)
        if r.status_code >= 300:
            raise ConversionError(f"The email service answered {r.status_code}: {r.text[:300]}")
        return
    if html:
        msg = MIMEMultipart("alternative")
        msg.attach(MIMEText(body, "plain", "utf-8"))
        msg.attach(MIMEText(html, "html", "utf-8"))
    else:
        msg = MIMEText(body, "plain", "utf-8")
    msg["Subject"] = subject
    msg["From"] = formataddr(("Academic-Home", EMAIL_ADDRESS))
    msg["To"] = to_email
    for attempt in (1, 2):
        try:
            with _GmailSession() as server:
                server.sendmail(EMAIL_ADDRESS, to_email, msg.as_string())
            return
        except (smtplib.SMTPServerDisconnected, smtplib.SMTPConnectError, ConnectionError, TimeoutError) as e:
            if attempt == 2:
                raise
            logger.warning(f"Email: the connection dropped ({e.__class__.__name__}) — trying once more")
            time.sleep(1.5)


# ---------- email confirmation codes ----------
def _needs_verification(row):
    """This account still has to confirm its email before using the app."""
    return bool(row) and (row.get("email_verified") if hasattr(row, "get") else True) is False and _email_ready()


def _code_hash(user_id, code, email):
    """A code only confirms the address it was sent to."""
    return hmac.new((SECRET_KEY or "docente").encode(), f"verify:{user_id}:{(email or '').strip().lower()}:{code}".encode(),
                    hashlib.sha256).hexdigest()


def _cancel_code(cur, user_id):
    """The address changed: the waiting code stops working (the hourly count is kept)."""
    cur.execute("UPDATE email_codes SET code_hash = '', expires_at = now() WHERE user_id = %s", (user_id,))


def _code_email(name, code):
    first = (str(name or "").strip().split() or ["there"])[0]
    text = (f"Hi {first},\n\nYour Academic-Home confirmation code is:\n\n    {code}\n\n"
            f"It expires in {EMAIL_CODE_MINUTES} minutes. Enter it in the app to confirm this is your email address.\n\n"
            "If you didn't create an Academic-Home account, you can ignore this email.\n\n— Academic-Home")
    spaced = " ".join(code)
    html = (
        '<div style="background:#f2f3f7;padding:32px 12px;font-family:Segoe UI,Helvetica,Arial,sans-serif;color:#14161c">'
        '<div style="max-width:440px;margin:0 auto;background:#ffffff;border-radius:18px;padding:30px 26px;border:1px solid #e3e5ec">'
        '<div style="font-size:20px;font-weight:800;letter-spacing:-.01em">Academic-Home</div>'
        f'<p style="font-size:16px;line-height:1.5;margin:18px 0 6px">Hi {first},</p>'
        '<p style="font-size:16px;line-height:1.5;margin:0 0 18px">Here is your confirmation code:</p>'
        f'<div style="font-size:34px;font-weight:800;letter-spacing:.18em;text-align:center;background:#fff7e6;border:1px solid #f1d38a;'
        f'border-radius:14px;padding:16px 8px;color:#14161c">{spaced}</div>'
        f'<p style="font-size:14px;line-height:1.5;color:#4b5060;margin:18px 0 0">It expires in {EMAIL_CODE_MINUTES} minutes. '
        "If you didn't create an Academic-Home account, you can ignore this email.</p>"
        '</div></div>')
    return text, html


def _send_verification_code(row, force=False, to=None):
    """Makes a fresh 6-digit code (replacing any earlier one) and emails it.
    Returns (sent, seconds_to_wait, error_message). force skips the short
    resend wait (a corrected email address), never the hourly cap."""
    if not _email_ready():
        return False, 0, "Email isn't set up on the server yet."
    now = datetime.now(timezone.utc)
    uid = row["id"]
    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT sent_at, sends_hour, hour_started FROM email_codes WHERE user_id = %s", (uid,))
            prev = cur.fetchone()
            sends, hour_started = 1, now
            if prev:
                since = (now - prev["sent_at"]).total_seconds()
                if not force and since < EMAIL_RESEND_SECONDS:
                    wait = int(EMAIL_RESEND_SECONDS - since) + 1
                    return False, wait, f"A code is on its way — you can ask for another in {wait} seconds."
                if now - prev["hour_started"] < timedelta(hours=1):
                    sends, hour_started = prev["sends_hour"] + 1, prev["hour_started"]
                    if sends > EMAIL_SENDS_PER_HOUR:
                        wait = int((prev["hour_started"] + timedelta(hours=1) - now).total_seconds()) + 1
                        return False, wait, "That's a lot of codes in an hour — please wait a little, then try again."
            code = f"{secrets.randbelow(1000000):06d}"
            cur.execute(
                "INSERT INTO email_codes (user_id, code_hash, expires_at, attempts, sent_at, sends_hour, hour_started) "
                "VALUES (%s, %s, %s, 0, %s, %s, %s) ON CONFLICT (user_id) DO UPDATE SET code_hash = EXCLUDED.code_hash, "
                "expires_at = EXCLUDED.expires_at, attempts = 0, sent_at = EXCLUDED.sent_at, sends_hour = EXCLUDED.sends_hour, "
                "hour_started = EXCLUDED.hour_started",
                (uid, _code_hash(uid, code, to or row["email"]), now + timedelta(minutes=EMAIL_CODE_MINUTES), now, sends, hour_started))
    finally:
        conn.close()
    text, html = _code_email(row.get("name"), code)
    try:
        send_email(to or row["email"], f"{code} is your Academic-Home code", text, html)
    except Exception as e:
        logger.warning(f"Couldn't send a confirmation code to user {uid}: {e}")
        _email_send_result(False, e, to or row["email"])
        # Let them try again straight away, and don't count the failed send against the hourly allowance.
        conn = get_db()
        try:
            with conn.cursor() as cur:
                cur.execute("UPDATE email_codes SET sent_at = now() - make_interval(secs => %s), sends_hour = GREATEST(sends_hour - 1, 0) "
                            "WHERE user_id = %s", (EMAIL_RESEND_SECONDS, uid))
        finally:
            conn.close()
        if _send_failure_kind(e) == "address":
            return False, 0, "That address can't receive our email — check it for typos (tap “Wrong email address?”)."
        return False, 0, "We couldn't send the email just now — please try “Send a new code” in a moment."
    _email_send_result(True)
    return True, EMAIL_RESEND_SECONDS, None


def _active_code(user_id):
    """There's an unexpired, unused code waiting for this account."""
    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT 1 FROM email_codes WHERE user_id = %s AND expires_at > now() AND attempts < %s", (user_id, EMAIL_CODE_MAX_TRIES))
            return bool(cur.fetchone())
    finally:
        conn.close()


def _verify_info(row, sent=None):
    """What the app needs to show the 'confirm your email' screen."""
    need = _needs_verification(row)
    info = {"required": need, "email": row["email"] if row else None}
    if sent is not None:
        info["sent"] = sent
    return info


# ---------- free Pro for everyone (October 2026) ----------
def _promo_until():
    if not PROMO_PRO_UNTIL or PROMO_PRO_UNTIL.strip().lower() in ("off", "no", "false", "0"):
        return None
    try:
        d = datetime.fromisoformat(PROMO_PRO_UNTIL.strip().replace("Z", "+00:00"))
    except ValueError:
        return None
    return d if d.tzinfo else d.replace(tzinfo=timezone.utc)


def _promo_active():
    until = _promo_until()
    return bool(until and datetime.now(timezone.utc) < until)


def _with_promo(row):
    """During the promotion a Free account works as Pro (Monthly), until the
    promotion ends — computed, never written, so nothing needs undoing after."""
    if not row or not _promo_active() or effective_plan(row) != "free":
        return row
    r = dict(row)
    r["plan"], r["plan_source"], r["plan_expires_at"] = "monthly", "promo", _promo_until()
    return r


# ---------- referral codes ----------
def _clean_code(code):
    return re.sub(r"[^A-Z0-9]", "", str(code or "").upper())[:40]


def _find_referrer(cur, code, user_id=None):
    """The account a referral code belongs to: someone's personal code, or a
    14-day code they made. Never the person's own code."""
    code = _clean_code(code)
    if not code:
        return None
    cur.execute("SELECT id, name FROM users WHERE referral_code = %s", (code,))
    ref = cur.fetchone()
    if not ref:
        cur.execute("SELECT u.id, u.name FROM promo_codes p JOIN users u ON u.id = p.created_by "
                    "WHERE p.code = %s AND p.expires_at > now()", (code,))
        ref = cur.fetchone()
    if ref and user_id and ref["id"] == user_id:
        return None
    return ref


def _reward_referral(cur, user_id, code=None):
    """A friend's code: both accounts get a bonus month, once per account — the
    first time it buys Monthly or Yearly with a code (typed at checkout, or the
    one saved at sign-up). Returns True when the bonus was given."""
    cur.execute("SELECT id, referred_code, referral_rewarded FROM users WHERE id = %s FOR UPDATE", (user_id,))
    me = cur.fetchone()
    if not me or me["referral_rewarded"]:
        return False
    code = _clean_code(code) or _clean_code(me.get("referred_code"))
    ref = _find_referrer(cur, code, user_id) if code else None
    if not ref:
        return False
    cur.execute("UPDATE users SET bonus_credit_months = bonus_credit_months + 1 WHERE id = %s", (ref["id"],))
    cur.execute("UPDATE users SET bonus_credit_months = bonus_credit_months + 1, referral_rewarded = TRUE, referred_code = %s WHERE id = %s",
                (code, user_id))
    logger.info(f"Referral bonus: user {user_id} with code from user {ref['id']}")
    return True


def _pass_expired(row):
    exp = row.get("plan_expires_at") if hasattr(row, "get") else None
    return bool(exp and row.get("plan") != "free" and exp <= datetime.now(timezone.utc))


def effective_plan(row):
    """The plan an account really has now: a Pay-per-Use pass whose 30 days
    have run out counts as Free, even before the database is corrected."""
    return "free" if _pass_expired(row) else row["plan"]


def user_row_to_dict(row):
    row = _with_promo(row)
    exp = row.get("plan_expires_at") if hasattr(row, "get") else None
    msgs_used, msgs_limit, msgs_reset = _sage_status(row)
    return {
        "id": row["id"],
        "name": row["name"],
        "email": row["email"],
        "plan": effective_plan(row),
        "plan_expires_at": exp.isoformat() if exp and not _pass_expired(row) else None,
        "remy_credits": (row.get("remy_credits") or 0) if hasattr(row, "get") else 0,
        "memory_enabled": _memory_on(row),
        # Where Pro came from — the app describes it truthfully: a store plan
        # renews, a Paystack pass or a bonus month runs to its date.
        "plan_source": (row.get("plan_source") if hasattr(row, "get") else None) if (row.get("plan") if hasattr(row, "get") else None) != "free" else None,
        # The Remy message meter: this window's allowance, used so far, and
        # when it renews (30 days after it began).
        "remy_messages_used": msgs_used,
        "remy_messages_limit": msgs_limit,
        "remy_messages_resets_at": (msgs_reset + timedelta(days=SAGE_WINDOW_DAYS)).isoformat() if msgs_reset else None,
        "referral_code": row["referral_code"],
        "bonus_credit_months": row.get("bonus_credit_months", 0) or 0,
        # Email confirmation: the app shows the code screen while this is True.
        "email_verified": bool(row.get("email_verified", True)) if hasattr(row, "get") else True,
        "verify_email_required": _needs_verification(row),
        # A new address waiting for its code (the account still uses "email" until then).
        "pending_email": (row.get("pending_email") if hasattr(row, "get") else None) or None,
        # A friend's referral code saved at sign-up, waiting for a first Pro purchase.
        "referral_pending": bool(row.get("referred_code") and not row.get("referral_rewarded")) if hasattr(row, "get") else False,
        "promo_until": _promo_until().isoformat() if _promo_active() else None,
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
        if _token_version(token) != (row.get("token_version") or 0):
            return jsonify({"error": "You were signed out — please log in again."}), 401
        if _pass_expired(row):
            # A Pay-per-Use pass has run out: back to Free, in the database
            # and for this request (allowances included).
            conn = get_db()
            try:
                with conn.cursor() as cur:
                    cur.execute("UPDATE users SET plan = 'free', plan_expires_at = NULL WHERE id = %s AND plan_expires_at <= now()", (user_id,))
                conn.commit()
            except Exception as e:
                logger.warning(f"Couldn't end an expired pass for user {user_id}: {e}")
            finally:
                conn.close()
            row = dict(row)
            row["plan"], row["plan_expires_at"] = "free", None
        elif _unpaid_plan(row):
            # Monthly/Yearly with no end date and no verified store purchase
            # behind it was never paid for: back to Free.
            conn = get_db()
            try:
                with conn.cursor() as cur:
                    cur.execute("UPDATE users SET plan = 'free', plan_expires_at = NULL WHERE id = %s AND plan IN " + _PRO_SQL + " AND plan_expires_at IS NULL", (user_id,))
                conn.commit()
                logger.warning(f"Removed a Monthly/Yearly plan with no payment behind it from user {user_id}")
            except Exception as e:
                logger.warning(f"Couldn't remove an unpaid plan from user {user_id}: {e}")
            finally:
                conn.close()
            row = dict(row)
            row["plan"], row["plan_expires_at"] = "free", None
        if _needs_verification(row) and fn.__name__ not in _UNCONFIRMED_OK:
            return jsonify({"error": f"Please confirm your email first — enter the 6-digit code we sent to {row['email']}.",
                            "verify_email": True}), 403
        if row.get("plan") == "free" and (row.get("bonus_credit_months") or 0) > 0 and not _promo_active():
            # A referral bonus month: kept in the bank while paid Pro runs (so
            # it never overlaps paid time) and started whenever the account is
            # on Free — straight away, or the moment paid time ends. The WHERE
            # clause lets each bonus month start exactly once.
            conn = get_db()
            try:
                with conn.cursor() as cur:
                    cur.execute("UPDATE users SET plan = 'monthly', plan_expires_at = now() + interval '30 days', plan_source = 'bonus', "
                                "bonus_credit_months = bonus_credit_months - 1 WHERE id = %s AND plan = 'free' AND bonus_credit_months > 0 RETURNING *", (user_id,))
                    started = cur.fetchone()
                conn.commit()
                if started:
                    row = started
                    logger.info(f"Bonus month started for user {user_id}")
            except Exception as e:
                logger.warning(f"Couldn't start a bonus month for user {user_id}: {e}")
            finally:
                conn.close()
        # Free Pro during the promotion (bonus months stay banked until it ends).
        request.current_user = _with_promo(row)
        return fn(*args, **kwargs)

    return wrapper


# What an account that hasn't confirmed its email yet may still do.
_UNCONFIRMED_OK = {"me", "verify_email", "resend_code", "change_email", "delete_account_endpoint", "update_profile"}


# 3.2: what each AI call used, logged as counts with an estimated cost (USD per million tokens; web search per
# search) so the real cost of a Remy answer can be checked — never anyone's words or account.
_AI_RATES = {"claude-opus-5": (5.0, 25.0), "claude-sonnet-5": (2.0, 10.0)}
_AI_SEARCH_COST = 0.01


def _log_ai_use(model, usage, feature):
    try:
        u = usage or {}
        cin, cout = _AI_RATES.get(model, (5.0, 25.0))
        tin, tout = int(u.get("input_tokens") or 0), int(u.get("output_tokens") or 0)
        cw, cr = int(u.get("cache_creation_input_tokens") or 0), int(u.get("cache_read_input_tokens") or 0)
        searches = int((u.get("server_tool_use") or {}).get("web_search_requests") or 0)
        cost = (tin * cin + cw * cin * 1.25 + cr * cin * 0.1 + tout * cout) / 1e6 + searches * _AI_SEARCH_COST
        logger.info(f"AI use: {feature or '-'} {model} in={tin} cache_w={cw} cache_r={cr} out={tout} searches={searches} cost~${cost:.4f}")
    except Exception:
        pass


# 3.3.1: what people see when the AI can't answer. Never the AI service's own words — they can name this account's
# credit balance or keys. The real reason goes to the server log for the owner (see _ai_failure).
AI_MSG_UNAVAILABLE = "Remy isn\u2019t available right now. Please try again a little later."
AI_MSG_BUSY = "Remy is busy right now. Please try again in a minute."
AI_MSG_SLOW = "Remy took too long to answer. Please try again."
AI_MSG_TOO_BIG = "That\u2019s more than Remy can read at once. Please try a shorter message or file."
AI_MSG_FAILED = "Remy couldn\u2019t answer just now. Please try again."
AI_MSG_NOT_COUNTED = " This didn\u2019t use any of your messages."


def _ai_failure(status, body, feature=None):
    """The AI service answered with an error: log the real reason for the owner, and return the plain words the
    person sees instead."""
    err = {}
    try:
        err = (json.loads(body or "{}") or {}).get("error") or {}
    except (ValueError, AttributeError):
        err = {}
    if not isinstance(err, dict):
        err = {}
    kind = str(err.get("type") or "")
    detail = str(err.get("message") or "") or str(body or "")[:300]
    low = detail.lower()
    where = feature or "-"
    if "credit balance" in low or "purchase credits" in low or kind == "billing_error" or status == 402:
        logger.error(f"AI is OFF for everyone: the Anthropic account has run out of credit. Add credit in the Anthropic "
                     f"Console (Plans & Billing); people see \"{AI_MSG_UNAVAILABLE}\" until then. ({where})")
        return AI_MSG_UNAVAILABLE
    if "usage limit" in low or "spend limit" in low:
        logger.error(f"AI is OFF for everyone: the Anthropic account reached its spending limit. Raise it in the Anthropic "
                     f"Console (Limits). ({where}): {detail[:200]}")
        return AI_MSG_UNAVAILABLE
    if status in (401, 403) or kind in ("authentication_error", "permission_error"):
        logger.error(f"AI is OFF for everyone: the AI service refused the key ({status} {kind}). Check ANTHROPIC_API_KEY "
                     f"on Render. ({where})")
        return AI_MSG_UNAVAILABLE
    if status == 429 or kind == "rate_limit_error":
        logger.warning(f"AI: rate limit reached ({where}): {detail[:200]}")
        return AI_MSG_BUSY
    if status == 413 or kind == "request_too_large" or "too long" in low:
        logger.warning(f"AI: the request was too large ({where}): {detail[:200]}")
        return AI_MSG_TOO_BIG
    if status >= 500 or kind in ("overloaded_error", "api_error"):
        logger.warning(f"AI: the AI service is having trouble ({status} {kind}) ({where})")
        return AI_MSG_BUSY
    logger.error(f"AI request failed ({status} {kind}) ({where}): {detail[:300]}")
    return AI_MSG_FAILED


def call_claude(system_prompt, user_message=None, max_tokens=600, use_search=False, return_meta=False, model=None, messages=None,
                timeout=None, feature=None):
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
        logger.error("AI is OFF for everyone: ANTHROPIC_API_KEY isn't set on Render.")
        raise ConversionError(AI_MSG_UNAVAILABLE)
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
            timeout=timeout or (140 if use_search else 110),
        )
    except requests.exceptions.Timeout:
        logger.warning(f"AI: no answer in time ({feature or '-'})")
        raise ConversionError(AI_MSG_SLOW)
    except requests.exceptions.RequestException as e:
        logger.warning(f"AI: couldn't reach the AI service ({feature or '-'}): {e}")
        raise ConversionError(AI_MSG_FAILED)
    if resp.status_code != 200:
        raise ConversionError(_ai_failure(resp.status_code, resp.text, feature))
    try:
        data = resp.json()
    except ValueError:
        logger.warning(f"AI: the answer wasn't readable ({feature or '-'})")
        raise ConversionError(AI_MSG_FAILED)
    _log_ai_use(payload["model"], data.get("usage"), feature)
    text = "".join(b.get("text", "") for b in data.get("content", []) if b.get("type") == "text")
    if return_meta:
        sources, queries = _search_trail(data.get("content") or [])
        return {"text": text, "stop_reason": data.get("stop_reason"), "sources": sources, "queries": queries,
                "content": data.get("content") or []}
    return text


def _search_trail(content):
    """What a web-searching answer looked at: the searches it ran, and the pages —
    the ones it actually cited first, then the others it found."""
    found, order, queries = {}, [], []
    for b in content:
        if not isinstance(b, dict):
            continue
        kind = b.get("type")
        if kind == "server_tool_use" and b.get("name") == "web_search":
            q = str((b.get("input") or {}).get("query") or "").strip()
            if q and q not in queries:
                queries.append(q[:200])
        elif kind == "web_search_tool_result" and isinstance(b.get("content"), list):
            for item in b["content"]:
                if isinstance(item, dict) and item.get("type") == "web_search_result" and item.get("url"):
                    url = str(item["url"])[:600]
                    if url not in found:
                        found[url] = {"title": str(item.get("title") or url).strip()[:200], "url": url, "cited": False}
                        order.append(url)
        elif kind == "text":
            for c in b.get("citations") or []:
                url = str((c or {}).get("url") or "")[:600]
                if not url.startswith("http"):
                    continue
                if url not in found:
                    found[url] = {"title": str(c.get("title") or url).strip()[:200], "url": url, "cited": True}
                    order.append(url)
                found[url]["cited"] = True
    ranked = [found[u] for u in order if found[u]["cited"]] + [found[u] for u in order if not found[u]["cited"]]
    return ranked, queries


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


# Semantic Scholar's key allows one request a second — counted across every worker process of this
# server through a small lock file holding the time of the last request.
_S2_GAP = 1.05
_S2_LOCK = os.path.join(tempfile.gettempdir(), "docente-s2.lock")
_S2_BUSY = "Semantic Scholar was busy (one request a second), so the other sources were used"


def _s2_slot(max_wait=1.2):
    """Waits for this server's next turn at Semantic Scholar. False when no turn comes within
    max_wait seconds — the caller then goes on without it."""
    import fcntl
    deadline = time.time() + max_wait
    while True:
        try:
            fd = os.open(_S2_LOCK, os.O_RDWR | os.O_CREAT, 0o600)
        except OSError:
            return True                      # no shared file: never block a search over it
        try:
            fcntl.flock(fd, fcntl.LOCK_EX)
            try:
                last = float(os.pread(fd, 40, 0).decode("ascii", "ignore").strip() or 0)
            except ValueError:
                last = 0.0
            now = time.time()
            if last > now + 60:              # a clock that went back: start again
                last = 0.0
            wait = last + _S2_GAP - now
            if wait <= 0:
                os.ftruncate(fd, 0)
                os.pwrite(fd, repr(now).encode("ascii"), 0)
                return True
            if now + wait > deadline:
                return False
        finally:
            try:
                fcntl.flock(fd, fcntl.LOCK_UN)
            finally:
                os.close(fd)
        time.sleep(min(wait, 0.25))           # wait outside the lock, then look again


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
        if not _s2_slot():
            return [], _S2_BUSY
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


# ------------------------------------------------ searching by topic
# Real use (server logs) showed students' whole instructions being searched —
# "Write a paragraph on vital signs" instead of "vital signs" — and a plain
# "Hi" running the full source search, spending the shared academic-search
# allowance for nothing. These turn a message into its topic, skip the search
# for small talk, and put any search failure into plain words for the student
# (the technical detail stays in the server log).
_INSTRUCTION_PREFIX = re.compile(
    r"^(?:(?:please|kindly|can you|could you|help me|i need|i want)\s+(?:to\s+)?)*"
    r"(?:(?:write|draft|give|generate|prepare|compose|create|produce|do|make)\s+(?:me\s+)?"
    r"(?:(?:a|an|the|one|two|three|\d+)\s+)?(?:(?:short|brief|detailed|long|full|simple|good|one[- ]page|\d+[- ]words?)\s+)*"
    r"(?:literature review|page|paragraph|essay|report|summary|note|review|introduction|conclusion|discussion|abstract|"
    r"assignment|article|section|write-?up|piece|answer|explanation)s?\s+"
    r"(?:of\s+\d+\s+words\s+)?(?:on|about|regarding|of|for|explaining)\s+"
    r"|(?:explain|describe|discuss|define|summari[sz]e|outline|tell me about|what (?:is|are)|write (?:on|about))\s+)",
    re.IGNORECASE)
_SMALL_TALK = re.compile(
    r"^(?:hi+|hello|hey|good (?:morning|afternoon|evening|night)|thanks?|thank you|ok(?:ay)?|yes|no|bye|goodbye|"
    r"how are you|good|great|cool|nice)\b[\s!.?,]*(?:remy|docente|academic[\s-]?home)?[\s!.?]*$", re.IGNORECASE)


def _search_topic(text):
    original = re.sub(r"\s+", " ", str(text or "")).strip()
    topic = original
    for _ in range(2):
        shorter = _INSTRUCTION_PREFIX.sub("", topic, count=1).strip()
        if shorter == topic:
            break
        topic = shorter
    topic = re.sub(r"^(?:the|a|an)\s+", "", topic, flags=re.IGNORECASE).strip(" .?!:;,")
    return (topic or original)[:200]


def _worth_searching(query):
    q = (query or "").strip()
    return len(q) >= 3 and not _SMALL_TALK.match(q) and bool(re.search(r"[A-Za-z]{2,}", q))


def _topic_search(text, max_results=8):
    """search_for_sources for a student's message: by its topic, and not at all
    for small talk. Returns (sources, diagnostic) like search_for_sources."""
    query = _search_topic(text)
    if not _worth_searching(query):
        return [], "not a topic"
    return search_for_sources(query, max_results=max_results)


# ------------------------------------------------ real papers anyone can download (3.0)
# Evidence comes from real papers with a free PDF: open-access works from OpenAlex, Semantic
# Scholar and Europe PMC, searched together. Nothing without a PDF link is kept; DOIs and
# abstracts come along so each citation can be traced to the paper. Any index that fails is
# skipped — the others still answer.
import concurrent.futures as _futures
import math as _math

OPENALEX_API_KEY = os.environ.get("OPENALEX_API_KEY", "").strip()
_PAPER_TIMEOUT = 9
_PAPER_UA = {"User-Agent": "Academic-Home/3.4 (study app; https://getdocente.vercel.app)"}
_PAPER_STOP = set((
    "a an and are as at be been being but by can could do does for from had has have how i if in into is it its "
    "me my of on or our so than that the their them then there these they this those to was we were what when "
    "where which while who whom why will with would you your about explain describe discuss define write essay "
    "paragraph review literature please give tell make short brief detailed note notes some any more most role "
    "effect effects use using used between among").split())


def _paper_terms(q):
    """The words that carry a query's meaning (no small words, plural -s dropped)."""
    out = []
    for w in re.findall(r"[a-z][a-z\-]{2,}", str(q or "").lower()):
        if w in _PAPER_STOP:
            continue
        w = w[:-1] if len(w) > 4 and w.endswith("s") and not w.endswith("ss") else w
        if w not in out:
            out.append(w)
    return out[:14]


def _oa_abstract(inv):
    """OpenAlex keeps an abstract as {word: [positions]}: put the words back in order."""
    if not isinstance(inv, dict) or not inv:
        return None
    pos = []
    for word, places in inv.items():
        for p in places or []:
            if isinstance(p, int):
                pos.append((p, str(word)))
    if not pos:
        return None
    pos.sort()
    text = " ".join(w for _, w in pos[:700]).strip()
    return text or None


def _norm_doi(d):
    d = re.sub(r"^(?:https?://(?:dx\.)?doi\.org/|doi:\s*)", "", str(d or "").strip(), flags=re.I).strip()
    return d.lower() if d.startswith("10.") else None


def _paper(title, authors, year, venue, doi, pdf, landing, abstract, found_in, cites=0, biblio=None):
    title = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", str(title or ""))).strip().rstrip(".")
    if not title or not pdf or not str(pdf).startswith("http"):
        return None
    try:
        year = int(year) if year else None
    except (TypeError, ValueError):
        year = None
    authors = [re.sub(r"\s+", " ", str(a)).strip() for a in (authors or []) if str(a or "").strip()][:20]
    abstract = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", str(abstract or ""))).strip() or None
    doi = _norm_doi(doi)
    return {"title": title[:400], "authors": authors or None, "year": year, "venue": (str(venue).strip()[:200] or None) if venue else None,
            "doi": doi, "pdf": str(pdf)[:600], "landing": (str(landing)[:600] if landing else None),
            "url": f"https://doi.org/{doi}" if doi else str(pdf)[:600],
            "abstract": abstract[:2000] if abstract else None, "found_in": found_in, "cites": int(cites or 0), "biblio": biblio or {}}


def _openalex_papers(query, n, since=None):
    flt = "is_oa:true" + (f",publication_year:>{since - 1}" if since else "")
    params = {"search": query[:300], "filter": flt, "per-page": max(n * 3, 15),
              "select": "id,doi,display_name,publication_year,authorships,primary_location,best_oa_location,locations,"
                        "abstract_inverted_index,biblio,cited_by_count,type"}
    if OPENALEX_API_KEY:
        params["api_key"] = OPENALEX_API_KEY
    resp = requests.get("https://api.openalex.org/works", params=params, headers=_PAPER_UA, timeout=_PAPER_TIMEOUT)
    if resp.status_code == 400:      # a field name the index no longer knows: ask without choosing fields
        params.pop("select", None)
        resp = requests.get("https://api.openalex.org/works", params=params, headers=_PAPER_UA, timeout=_PAPER_TIMEOUT)
    if resp.status_code != 200:
        return [], f"OpenAlex HTTP {resp.status_code}"
    out = []
    for w in (resp.json() or {}).get("results") or []:
        if not isinstance(w, dict):
            continue
        best = w.get("best_oa_location") or {}
        locs = [l for l in (w.get("locations") or []) if isinstance(l, dict)]
        pdf = best.get("pdf_url") or next((l.get("pdf_url") for l in locs if l.get("pdf_url")), None)
        prim = w.get("primary_location") or {}
        venue = ((prim.get("source") or {}).get("display_name")) or ((best.get("source") or {}).get("display_name"))
        authors = [((a or {}).get("author") or {}).get("display_name") for a in (w.get("authorships") or [])]
        b = w.get("biblio") or {}
        p = _paper(w.get("display_name") or w.get("title"), authors, w.get("publication_year"), venue, w.get("doi"), pdf,
                   best.get("landing_page_url") or prim.get("landing_page_url"), _oa_abstract(w.get("abstract_inverted_index")),
                   "OpenAlex", w.get("cited_by_count"),
                   {k: b.get(k) for k in ("volume", "issue", "first_page", "last_page") if b.get(k)})
        if p:
            out.append(p)
    return out, None if out else "OpenAlex found no open papers with a PDF"


def _s2_papers(query, n, since=None):
    headers = dict(_PAPER_UA)
    if SEMANTIC_SCHOLAR_API_KEY:
        headers["x-api-key"] = SEMANTIC_SCHOLAR_API_KEY
    params = {"query": query[:300], "limit": max(n * 3, 15), "openAccessPdf": "",
              "fields": "title,url,year,authors,abstract,venue,openAccessPdf,externalIds,citationCount,journal"}
    if since:
        params["year"] = f"{since}-"
    if not _s2_slot():
        return [], _S2_BUSY
    resp = requests.get("https://api.semanticscholar.org/graph/v1/paper/search", params=params, headers=headers, timeout=_PAPER_TIMEOUT)
    if resp.status_code == 400:      # the "has a PDF" filter refused: ask without it (each paper is still checked below)
        params.pop("openAccessPdf", None)
        if not _s2_slot():
            return [], _S2_BUSY
        resp = requests.get("https://api.semanticscholar.org/graph/v1/paper/search", params=params, headers=headers, timeout=_PAPER_TIMEOUT)
    if resp.status_code != 200:
        return [], f"Semantic Scholar HTTP {resp.status_code}"
    out = []
    for x in (resp.json() or {}).get("data") or []:
        if not isinstance(x, dict):
            continue
        pdf = ((x.get("openAccessPdf") or {}).get("url") or "").strip()
        j = x.get("journal") or {}
        pages = str(j.get("pages") or "").strip()
        first, _, last = pages.partition("-")
        biblio = {k: v for k, v in (("volume", str(j.get("volume") or "").strip()), ("first_page", first.strip()), ("last_page", last.strip())) if v}
        p = _paper(x.get("title"), [(a or {}).get("name") for a in (x.get("authors") or [])], x.get("year"),
                   x.get("venue") or j.get("name"), (x.get("externalIds") or {}).get("DOI"), pdf, x.get("url"), x.get("abstract"),
                   "Semantic Scholar", x.get("citationCount"), biblio)
        if p:
            out.append(p)
    return out, None if out else "Semantic Scholar found no open papers with a PDF"


def _epmc_papers(query, n, since=None):
    q = f"({query[:250]}) AND OPEN_ACCESS:y AND HAS_PDF:y" + (f" AND PUB_YEAR:[{since} TO 3000]" if since else "")
    params = {"query": q, "format": "json", "resultType": "core", "pageSize": max(n * 2, 12)}
    resp = requests.get("https://www.ebi.ac.uk/europepmc/webservices/rest/search", params=params, headers=_PAPER_UA, timeout=_PAPER_TIMEOUT)
    if resp.status_code != 200:
        return [], f"Europe PMC HTTP {resp.status_code}"
    out = []
    for x in ((resp.json() or {}).get("resultList") or {}).get("result") or []:
        if not isinstance(x, dict):
            continue
        urls = ((x.get("fullTextUrlList") or {}).get("fullTextUrl")) or []
        pdf = next((u.get("url") for u in urls if isinstance(u, dict) and str(u.get("documentStyle")).lower() == "pdf"
                    and str(u.get("availability", "")).lower() in ("open access", "free") and u.get("url")), None)
        if not pdf and x.get("pmcid"):
            pdf = f"https://europepmc.org/articles/{x['pmcid']}?pdf=render"
        names = []
        for a in ((x.get("authorList") or {}).get("author")) or []:
            if isinstance(a, dict):
                full = " ".join(v for v in (a.get("firstName"), a.get("lastName")) if v) or a.get("fullName")
                names.append(full)
        journal = (((x.get("journalInfo") or {}).get("journal")) or {}).get("title")
        pages = str(x.get("pageInfo") or "")
        first, _, last = pages.partition("-")
        biblio = {k: v for k, v in (("volume", str((x.get("journalInfo") or {}).get("volume") or "")), ("issue", str((x.get("journalInfo") or {}).get("issue") or "")),
                                     ("first_page", first.strip()), ("last_page", last.strip())) if v}
        landing = f"https://europepmc.org/article/{x.get('source', 'MED')}/{x.get('id')}" if x.get("id") else None
        p = _paper(x.get("title"), names, x.get("pubYear"), journal, x.get("doi"), pdf, landing, x.get("abstractText"),
                   "Europe PMC", x.get("citedByCount"), biblio)
        if p:
            out.append(p)
    return out, None if out else "Europe PMC found no open papers with a PDF"


def _title_key(t):
    return re.sub(r"[^a-z0-9]+", "", str(t or "").lower())[:120]


def _rank_papers(papers, query):
    """Most relevant first: how many of the query's words the title and abstract share, then
    having an abstract (something to ground a sentence in), then being recent and well cited."""
    terms = _paper_terms(query)
    year_now = datetime.now(timezone.utc).year

    def score(p):
        rank = p.get("_rank", 0)
        title = (p["title"] or "").lower()
        abstract = (p.get("abstract") or "").lower()
        hit_t = sum(1 for t in terms if t in title)
        hit_a = sum(1 for t in terms if t in abstract)
        cover = (hit_t + 0.5 * hit_a) / max(1, len(terms))
        recent = 1.0 if (p.get("year") or 0) >= year_now - 5 else 0.5 if (p.get("year") or 0) >= year_now - 10 else 0
        return cover * 6 + (1.2 if p.get("abstract") else 0) + recent + min(_math.log10((p.get("cites") or 0) + 1), 3) * 0.3 - rank * 0.04

    return sorted(papers, key=score, reverse=True)


def _pdf_alive(url):
    """False only for a link that is plainly dead: 404/410, or no such site."""
    try:
        r = requests.head(url, allow_redirects=True, timeout=4, headers=_PAPER_UA)
        return r.status_code not in (404, 410)
    except requests.exceptions.SSLError:
        return True
    except requests.exceptions.Timeout:
        return True
    except requests.exceptions.RequestException:
        return False
    except Exception:
        return True


def search_papers(query, max_results=8, since=None, check_links=True):
    """Real, downloadable papers for a topic: [{title, authors, year, venue, doi, pdf, url,
    abstract, ...}] most relevant first, and a reason when there are none."""
    query = re.sub(r"\s+", " ", str(query or "")).strip()
    if not query:
        return [], "no topic"
    order = (("OpenAlex", _openalex_papers), ("Semantic Scholar", _s2_papers), ("Europe PMC", _epmc_papers))
    got, why = {}, []
    with _futures.ThreadPoolExecutor(max_workers=3) as pool:
        jobs = {name: pool.submit(fn, query, max_results, since) for name, fn in order}
        for name, _ in order:
            try:
                papers, diag = jobs[name].result(timeout=_PAPER_TIMEOUT * 2 + 3)
            except Exception as e:
                papers, diag = [], f"{name} failed: {e}"
            for i, p in enumerate(papers):
                p["_rank"] = i          # its place in that index's own relevance order
            got[name] = papers
            if diag:
                why.append(diag)
    found = [p for name, _ in order for p in got.get(name, [])]
    seen_doi, seen_title, merged = set(), set(), []
    for p in found:
        tk = _title_key(p["title"])
        if (p["doi"] and p["doi"] in seen_doi) or tk in seen_title:
            # the same paper from another index: keep whatever this copy adds
            twin = next((m for m in merged if (p["doi"] and m["doi"] == p["doi"]) or _title_key(m["title"]) == tk), None)
            if twin is not None:
                for k in ("abstract", "venue", "doi", "authors", "year"):
                    if not twin.get(k) and p.get(k):
                        twin[k] = p[k]
                if not twin.get("biblio") and p.get("biblio"):
                    twin["biblio"] = p["biblio"]
                if twin["doi"]:
                    twin["url"] = f"https://doi.org/{twin['doi']}"
            continue
        if p["doi"]:
            seen_doi.add(p["doi"])
        seen_title.add(tk)
        merged.append(p)
    ranked = _rank_papers(merged, query)[: max_results * 2]
    if check_links and ranked:
        alive = {}
        with _futures.ThreadPoolExecutor(max_workers=8) as pool:
            jobs = {pool.submit(_pdf_alive, p["pdf"]): i for i, p in enumerate(ranked)}
            try:
                for job in _futures.as_completed(jobs, timeout=6):
                    try:
                        alive[jobs[job]] = job.result()
                    except Exception:
                        alive[jobs[job]] = True
            except _futures.TimeoutError:
                pass
        ranked = [p for i, p in enumerate(ranked) if alive.get(i, True)]
    papers = ranked[:max_results]
    for p in papers:
        p.pop("_rank", None)
    logger.info("search_papers: %r -> %d paper(s) (%s)", query[:80], len(papers), "; ".join(why) or "ok")
    return papers, (None if papers else ("; ".join(why) or "no open papers with a PDF were found"))


def paper_topic_search(text, max_results=8, since=None):
    """search_papers for what a person wrote: by its topic, and not at all for small talk."""
    query = _search_topic(text)
    if not _worth_searching(query):
        return [], "not a topic"
    return search_papers(query, max_results=max_results, since=since)


def _no_sources_note(diag, subject):
    """What the writing model is told when no sources are available — in plain
    words, never the technical diagnostic (which named HTTP codes, service
    messages and "API key is invalid" for students to read)."""
    if diag == "not a topic":
        return ("No source search was needed: the student's latest message isn't a writing topic (a greeting or "
                "a quick remark). Reply naturally and briefly — if they haven't said what they'd like written yet, "
                "ask for the topic, how long it should be and which course it's for.")
    d = (diag or "").lower()
    if "429" in d or "rate-limit" in d or "too many requests" in d:
        reason = "the academic database is busy right now, so trying again in a minute or two usually works"
    elif "401" in d or "403" in d or "authentication" in d:
        reason = "the source search isn't available at the moment"
    elif "timeout" in d or "timed out" in d or "connection" in d:
        reason = "the source search didn't respond in time, so trying again usually works"
    else:
        reason = "no matching academic sources turned up for this topic, so a more specific topic may help"
    return (f"Source search ran for {subject} and came back with nothing usable — the reason, in plain words: "
            f"\"{reason}\". Write from your own well-informed general knowledge instead — but do NOT insert any "
            "placeholder in place of a citation, in any form: no [Author, Year], no [Citation needed], no [source], "
            "no blank brackets, nothing standing in for a reference that isn't there. If the topic genuinely calls "
            "for citation-backed evidence you don't have, say so plainly to the student in a sentence or two, giving "
            "that plain-words reason so they know what to do next — never technical details such as error codes, "
            "service names or error messages. Write as much as the task genuinely needs.")


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
        title = e["title"].rstrip(".")
        vol, iss, pages = _biblio_parts(e)
        tail = ""
        if e.get("venue"):      # APA 7: *Journal*, *volume*(issue), pages.
            tail = f" *{e['venue']}*" + (f", *{vol}*" if vol else "") + (f"({iss})" if vol and iss else "") + (f", {pages}" if pages else "") + "."
        e["ref_text"] = f"{e['author_refs']} ({e['year_display']}). {title}." + tail

    return enriched


def _biblio_parts(e):
    """(volume, issue, pages) as the index gave them; pages as "100\u2013110"."""
    b = e.get("biblio") or {}
    vol, iss = str(b.get("volume") or "").strip(), str(b.get("issue") or "").strip()
    first, last = str(b.get("first_page") or "").strip(), str(b.get("last_page") or "").strip()
    pages = f"{first}\u2013{last}" if first and last and first != last else first
    return vol[:20], iss[:20], pages[:30]


# Citation styles for Evidence-Based Writing. Every style is built from the
# same real metadata (authors, year, title, venue, link) as APA — nothing is
# invented — and keeps the same a/b suffixes when one author has two works
# in the same year, so the in-text keys and the references always agree.
CITE_STYLES = {
    "apa": {"name": "APA 7", "heading": "References", "example": "(Smith, 2023)"},
    "harvard": {"name": "Harvard", "heading": "Reference list", "example": "(Smith 2023)"},
    "mla": {"name": "MLA 9", "heading": "Works Cited", "example": "(Smith)"},
    "chicago": {"name": "Chicago (author-date)", "heading": "References", "example": "(Smith 2023)"},
}


def enrich_sources(sources, style="apa"):
    enriched = enrich_sources_for_apa(sources)
    if style not in CITE_STYLES or style == "apa":
        return enriched
    for e in enriched:
        names = [a.split() for a in (e.get("authors") or []) if a.split()]
        surn = [n[-1] for n in names]
        org = e["author_intext"]
        yd = e["year_display"]
        title = e["title"].rstrip(".")
        venue = e.get("venue")
        if not surn:
            au = org
        elif len(surn) == 1:
            au = surn[0]
        elif len(surn) == 2:
            au = f"{surn[0]} and {surn[1]}"
        elif style == "chicago" and len(surn) == 3:
            au = f"{surn[0]}, {surn[1]}, and {surn[2]}"
        elif style == "harvard" and len(surn) == 3:
            au = f"{surn[0]}, {surn[1]} and {surn[2]}"
        else:
            au = f"{surn[0]} et al."
        e["citation_key"] = f"({au})" if style == "mla" else f"({au} {yd})"
        inverted = (f"{names[0][-1]}, {' '.join(names[0][:-1])}".strip().rstrip(",")) if names else org
        if style == "harvard":
            if names:
                parts = [(f"{n[-1]}, " + " ".join(p[0] + "." for p in n[:-1])).strip().rstrip(",") for n in names]
                auth = parts[0] + " et al." if len(parts) > 3 else parts[0] if len(parts) == 1 else ", ".join(parts[:-1]) + " and " + parts[-1]
            else:
                auth = org
            vol, iss, pages = _biblio_parts(e)
            e["ref_text"] = f"{auth} ({yd}) " + (f"\u2018{title}\u2019, *{venue}*" + (f", {vol}" if vol else "") + (f"({iss})" if vol and iss else "")
                                                  + (f", pp. {pages}" if pages else "") + "." if venue else f"*{title}*.")
        elif style == "chicago":
            full = [" ".join(n) for n in names]
            if not names:
                auth = org
            elif len(full) == 1:
                auth = inverted
            elif len(full) > 10:
                auth = inverted + ", " + ", ".join(full[1:7]) + ", et al."
            elif len(full) == 2:
                auth = f"{inverted}, and {full[1]}"
            else:
                auth = inverted + ", " + ", ".join(full[1:-1]) + ", and " + full[-1]
            vol, iss, pages = _biblio_parts(e)
            e["ref_text"] = f"{auth}. {yd}. \u201c{title}.\u201d" + (f" *{venue}*" + (f" {vol}" if vol else "") + (f" ({iss})" if vol and iss else "")
                                                                  + (f": {pages}" if pages else "") + "." if venue else "")
        else:   # MLA 9
            year = str(e.get("year") or "n.d.")
            if names:
                auth = inverted if len(names) == 1 else f"{inverted}, and {' '.join(names[1])}" if len(names) == 2 else f"{inverted}, et al"
                vol, iss, pages = _biblio_parts(e)
                e["ref_text"] = (f"{auth}. \u201c{title}.\u201d " + (f"*{venue}*, " if venue else "") + (f"vol. {vol}, " if vol else "")
                                 + (f"no. {iss}, " if iss else "") + f"{year}" + (f", pp. {pages}" if pages else "") + ".")
            else:
                e["ref_text"] = f"\u201c{title}.\u201d *{venue or org}*, {year}."
    if style == "mla":     # two works by one author: MLA tells them apart by a short title
        counts = Counter(e["citation_key"] for e in enriched)
        for e in enriched:
            if counts[e["citation_key"]] > 1:
                short = " ".join(e["title"].split()[:3]).rstrip(".,:;")
                e["citation_key"] = e["citation_key"][:-1] + f", \u201c{short}\u201d)"
    return enriched


def build_sources_block(sources):
    """Source list for the writing prompt, keyed by the exact citation (in the chosen style)
    the model must copy verbatim. Sources with an abstract (from Semantic
    Scholar) include a short excerpt, so the model can cite what the paper
    actually found instead of guessing from the title alone."""
    lines = []
    for s in sources:
        line = f"{s['citation_key']} {s['title']}"
        if s.get("venue") or s.get("year"):
            line += f" \u2014 {s.get('venue') or 'paper'}, {s.get('year') or 'n.d.'}"
        if s.get("abstract"):
            line += f"\n    Abstract: {s['abstract'][:700]}"
        lines.append(line)
    return "\n".join(lines)


def build_references_list(sources, cited=None):
    """The reference list: only the sources the text cites (all of them when cited is None),
    alphabetical as APA requires. Each entry: its text, its link (DOI, or the PDF), the PDF
    itself, the DOI, and its in-text key — so a download can link every citation to it."""
    picked = [s for i, s in enumerate(sources) if cited is None or i in cited]
    ordered = sorted(picked, key=lambda s: (s["author_intext"].lower(), s.get("year_display") or ""))
    return [{"text": s["ref_text"], "url": s["url"], "pdf": s.get("pdf") or None, "doi": s.get("doi") or None,
             "key": s.get("citation_key") or ""} for s in ordered]


# An in-text citation's shape: a capitalised name, then the year (or n.d.), maybe a page.
_CITE_SHAPE = re.compile(r"^[A-Z\u00C0-\u024F][^;()]{0,120}?,?\s+(?:(?:1[5-9]|20)\d{2}[a-z]?|n\.\s?d\.)"
                         r"(?:\s*,\s*(?:p|pp|para|chap|ch)\.?\s*[\w\u2013\-]+)?$")


def _cite_norm(s):
    s = str(s or "").replace("&", " and ").replace("\u2019", "'").replace("\u2018", "'").lower()
    s = re.sub(r"\s+", " ", s)
    return s.strip(" ,.;")


def keep_known_citations(text, sources, style="apa"):
    """Keeps the in-text citations that match a real source and takes out any other, so every
    citation left in the text leads to a real paper. Returns (text, cited): the indices of the
    sources the text cites (in parentheses, or as "Mensah and Boateng (2022)")."""
    if not text:
        return text, set()
    inners = {}
    for i, s in enumerate(sources or []):
        key = str(s.get("citation_key") or "")
        if key.startswith("(") and key.endswith(")"):
            inners[_cite_norm(key[1:-1])] = i
    cited = set()

    def known(part):
        n = _cite_norm(part)
        if n in inners:
            return inners[n]
        for k, i in inners.items():
            if n.startswith(k + ",") or n.startswith(k + " p") or re.search(r"(?:^|\W)" + re.escape(k) + r"(?:$|[\s,])", n):
                return i
        return None

    def fix(m):
        parts = [p.strip() for p in m.group(1).split(";")]
        kept, dropped = [], False
        for p in parts:
            i = known(p)
            if i is not None:
                cited.add(i)
                kept.append(p)
            elif style != "mla" and _CITE_SHAPE.match(p):
                dropped = True          # looks like a citation, but no such source: out
            else:
                kept.append(p)
        if not dropped:
            return m.group(0)
        return "(" + "; ".join(kept) + ")" if kept else ""

    out = re.sub(r"\(([^()]{2,400})\)", fix, text)
    out = re.sub(r"[ \t]+([.,;:])", r"\1", out)
    out = re.sub(r"[ \t]{2,}", " ", out)
    out = re.sub(r"[ \t]+(?=\n|$)", "", out)
    # narrative citations: "Mensah and Boateng (2022)", "Asante et al. (2019)"
    low = _cite_norm(out)
    for k, i in inners.items():
        name, _, year = k.rpartition(",") if "," in k else k.rpartition(" ")
        name, year = name.strip(), year.strip()
        if name and year and re.search(re.escape(name) + r"\s*\(" + re.escape(year) + r"\)", low):
            cited.add(i)
    return out, cited


@app.after_request
def add_cors_headers(resp):
    origin = request.headers.get("Origin", "")
    if _origin_is_allowed(origin):
        resp.headers["Access-Control-Allow-Origin"] = origin
        resp.headers["Vary"] = "Origin"
    resp.headers["Access-Control-Allow-Methods"] = "POST, GET, DELETE, OPTIONS"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization"
    resp.headers["Access-Control-Expose-Headers"] = "X-Batch-Summary, X-Docente-Note"
    # 3.4: how a browser may treat this reply
    resp.headers.setdefault("X-Content-Type-Options", "nosniff")
    resp.headers.setdefault("X-Frame-Options", "DENY")
    resp.headers.setdefault("Referrer-Policy", "no-referrer")
    resp.headers.setdefault("Strict-Transport-Security", "max-age=31536000")
    if resp.mimetype == "application/json":
        resp.headers.setdefault("Content-Security-Policy", "default-src 'none'; frame-ancestors 'none'")
    if request.headers.get("Authorization") and "Cache-Control" not in resp.headers:
        resp.headers["Cache-Control"] = "no-store"
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
    return jsonify({"status": "ok", "email": _email_ready(),
                    "promo_until": _promo_until().isoformat() if _promo_active() else None})


@app.route("/api/promo", methods=["GET"])
def promo_status():
    """For the app's banner — no account needed."""
    return jsonify({"active": _promo_active(), "until": _promo_until().isoformat() if _promo_active() else None})


# ---------- auth ----------


@app.route("/api/auth/signup", methods=["POST", "OPTIONS"])
@limiter.limit("10 per hour")
def signup():
    if request.method == "OPTIONS":
        return "", 204
    if not DATABASE_URL or not SECRET_KEY:
        return jsonify({"error": "Accounts aren't configured on the server yet"}), 500

    data = _json_body()
    name = _text(data.get("name")).strip()
    email = _text(data.get("email")).strip().lower()
    password = _text(data.get("password"))

    if not name or not email or not password:
        return jsonify({"error": "Name, email, and password are all required"}), 400
    if len(name) > 200:
        return jsonify({"error": "That name is too long (200 characters max)"}), 400
    if len(email) > 254:  # RFC 5321's own limit on a valid email address
        return jsonify({"error": "That email is too long"}), 400
    if not _valid_new_email(email):
        return jsonify({"error": "That doesn't look like a valid email"}), 400
    if len(password) < 8:
        return jsonify({"error": "Password must be at least 8 characters"}), 400
    if len(password) > 200:
        return jsonify({"error": "That password is too long (200 characters max)"}), 400
    friend_code = _clean_code(data.get("referral_code"))

    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT id FROM users WHERE email = %s", (email,))
            if cur.fetchone():
                return jsonify({"error": "An account with that email already exists — log in instead, or reset your password."}), 409
            if friend_code and not _find_referrer(cur, friend_code):
                return jsonify({"error": "That referral code wasn't found — check it, or leave the box empty.", "field": "referral"}), 400

            referral_code = generate_referral_code(name)
            try:
                cur.execute(
                    """
                    INSERT INTO users (name, email, password_hash, referral_code, email_verified, referred_code)
                    VALUES (%s, %s, %s, %s, %s, %s)
                    RETURNING *
                    """,
                    # Accounts made while email isn't set up count as confirmed, like the older ones.
                    (name, email, generate_password_hash(password), referral_code, not _email_configured(), friend_code or None),
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
                return jsonify({"error": "An account with that email already exists — log in instead, or reset your password."}), 409
    finally:
        conn.close()

    sent = None
    if _needs_verification(row):
        sent, _, _ = _send_verification_code(row)
    return jsonify({"token": _token_for(row), "user": user_row_to_dict(row), "verify": _verify_info(row, sent)})


@app.route("/api/auth/login", methods=["POST", "OPTIONS"])
@limiter.limit("10 per minute; 30 per hour")
def login():
    if request.method == "OPTIONS":
        return "", 204
    if not DATABASE_URL or not SECRET_KEY:
        return jsonify({"error": "Accounts aren't configured on the server yet"}), 500

    data = _json_body()
    email = _text(data.get("email")).strip().lower()
    password = _text(data.get("password"))

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

    sent = None
    if _needs_verification(row) and not _active_code(row["id"]):
        sent, _, _ = _send_verification_code(row)
    return jsonify({"token": _token_for(row), "user": user_row_to_dict(row), "verify": _verify_info(row, sent)})


@app.route("/api/auth/forgot-password", methods=["POST", "OPTIONS"])
@limiter.limit("5 per hour")
def forgot_password():
    if request.method == "OPTIONS":
        return "", 204
    if not DATABASE_URL or not SECRET_KEY:
        return jsonify({"error": "Accounts aren't configured on the server yet"}), 500
    if not _email_ready():
        return jsonify({"error": "We can't send password-reset emails right now — please try again later, or write to support.docente@gmail.com."}), 503

    data = _json_body()
    email = _text(data.get("email")).strip().lower()
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
        token = make_reset_token(row["id"], email)
        reset_link = f"{FRONTEND_URL}/?reset_token={token}"
        try:
            send_email(
                email,
                "Reset your Academic-Home password",
                f"Hi {row['name'] or ''},\n\n"
                "We received a request to reset your Academic-Home password. "
                "This link expires in 1 hour:\n\n"
                f"{reset_link}\n\n"
                "If you didn't request this, you can safely ignore this email — "
                "your password won't change unless you click the link above.",
            )
        except Exception as e:
            # Never a 500, and never a hint about whether the address has an account.
            logger.warning(f"Couldn't send a password reset email: {e}")
            _email_send_result(False, e, email)

    return jsonify({"message": "If that email has an account, a reset link has been sent."})


@app.route("/api/auth/reset-password", methods=["POST", "OPTIONS"])
@limiter.limit("20 per hour")
def reset_password():
    if request.method == "OPTIONS":
        return "", 204
    if not DATABASE_URL or not SECRET_KEY:
        return jsonify({"error": "Accounts aren't configured on the server yet"}), 500

    data = _json_body()
    token = _text(data.get("token"))
    new_password = _text(data.get("password"))

    if len(new_password) < 8:
        return jsonify({"error": "Password must be at least 8 characters"}), 400

    user_id = verify_reset_token(token)
    if not user_id:
        return jsonify({"error": "This reset link is invalid or has expired — request a new one"}), 400
    link_email = _reset_token_email(token)

    conn = get_db()
    try:
        with conn.cursor() as cur:
            # Receiving the link proves the address it was sent to: if that's still the account's
            # address, it counts as confirmed (and a Google link made before it was confirmed is
            # dropped). Every other signed-in device is signed out.
            cur.execute(
                """UPDATE users SET password_hash = %s, token_version = token_version + 1,
                       google_sub = CASE WHEN email_verified OR lower(email) <> %s THEN google_sub ELSE NULL END,
                       email_verified = email_verified OR lower(email) = %s,
                       email_proven = email_proven OR lower(email) = %s
                   WHERE id = %s""",
                (generate_password_hash(new_password), link_email or "", link_email or "", link_email or "", user_id),
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
@limiter.limit("20 per minute; 60 per hour")
def google_signin():
    if request.method == "OPTIONS":
        return "", 204
    if not DATABASE_URL or not SECRET_KEY:
        return jsonify({"error": "Accounts aren't configured on the server yet"}), 500
    if not GOOGLE_CLIENT_ID:
        return jsonify({"error": "Google sign-in isn't configured on the server yet"}), 500

    from google.oauth2 import id_token as google_id_token
    from google.auth.transport import requests as google_requests

    data = _json_body()
    credential = _text(data.get("credential"))
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
    google_checked = payload.get("email_verified") is True or str(payload.get("email_verified")).lower() == "true"
    name = payload.get("name") or email.split("@")[0]
    if not email:
        return jsonify({"error": "Google didn't share an email address for this account."}), 400

    notice = None
    conn = get_db()
    try:
        with conn.cursor() as cur:
            # The Google account itself first; the email address only if this Google account is new here.
            cur.execute("SELECT * FROM users WHERE google_sub = %s", (google_sub,))
            row = cur.fetchone()
            if row:
                if google_checked and (row["email"] or "").lower() == email and not (row.get("email_verified") and row.get("email_proven")):
                    # Google has confirmed this very address.
                    cur.execute("UPDATE users SET email_verified = TRUE, email_proven = TRUE WHERE id = %s RETURNING *", (row["id"],))
                    row = cur.fetchone()
            else:
                cur.execute("SELECT * FROM users WHERE email = %s", (email,))
                row = cur.fetchone()
                if row and not google_checked:
                    return jsonify({"error": "This email already has an Academic-Home account — log in with your password."}), 409
                keep_ref = None
                if row and not row.get("email_proven"):
                    # An account holding this address that never proved it was theirs — made with a
                    # password, or by another Google account that switched its address to this one.
                    # Google has now proved the address belongs to this person: the account becomes
                    # theirs and confirmed, any other Google account and the unproven password are
                    # unlinked, and every other sign-in it handed out ends. Nothing is deleted.
                    notice = "password_off" if row.get("password_hash") else None
                    try:
                        cur.execute("UPDATE users SET google_sub = %s, email_verified = TRUE, email_proven = TRUE, password_hash = NULL, "
                                    "token_version = token_version + 1 WHERE id = %s RETURNING *", (google_sub, row["id"]))
                        row = cur.fetchone()
                        logger.warning(f"Account {row['id']} reclaimed on Google sign-in by the address's owner (old password and links switched off)")
                    except psycopg2.errors.UniqueViolation:
                        conn.rollback()
                        cur.execute("SELECT * FROM users WHERE google_sub = %s", (google_sub,))
                        row = cur.fetchone()
                        if not row:
                            return jsonify({"error": "Google sign-in failed. Please try again."}), 500
                elif row and not row.get("google_sub"):
                    try:
                        cur.execute("UPDATE users SET google_sub = %s WHERE id = %s RETURNING *", (google_sub, row["id"]))
                        row = cur.fetchone()
                    except psycopg2.errors.UniqueViolation:
                        conn.rollback()
                        cur.execute("SELECT * FROM users WHERE google_sub = %s", (google_sub,))
                        row = cur.fetchone()
                        if not row:
                            return jsonify({"error": "Google sign-in failed. Please try again."}), 500
                elif not row:
                    referral_code = generate_referral_code(name)
                    try:
                        cur.execute(
                            """
                            INSERT INTO users (name, email, google_sub, referral_code, email_verified, email_proven, referred_code)
                            VALUES (%s, %s, %s, %s, %s, %s, %s)
                            RETURNING *
                            """,
                            (name, email, google_sub, referral_code, bool(google_checked), bool(google_checked), keep_ref),
                        )
                        row = cur.fetchone()
                    except psycopg2.errors.UniqueViolation:
                        # The same narrow race as signup's — two sign-in attempts for the same
                        # brand-new Google account at nearly the same moment: the other request
                        # already created it, so carry on with that one.
                        conn.rollback()
                        cur.execute("SELECT * FROM users WHERE google_sub = %s", (google_sub,))
                        row = cur.fetchone()
                        if not row:
                            return jsonify({"error": "Google sign-in failed. Please try again."}), 500
    finally:
        conn.close()

    sent = None
    if _needs_verification(row) and not _active_code(row["id"]):
        sent, _, _ = _send_verification_code(row)
    out = {"token": _token_for(row), "user": user_row_to_dict(row), "verify": _verify_info(row, sent)}
    if notice:
        out["notice"] = notice
    return jsonify(out)


@app.route("/api/auth/verify-email", methods=["POST", "OPTIONS"])
@limiter.limit("30 per hour", key_func=_account_or_ip_key)
@auth_required
def verify_email():
    """Confirms the sign-up address, or (for a confirmed account) a new address waiting in pending_email."""
    row = request.current_user
    if row.get("email_verified", True) is not False and not row.get("pending_email"):
        return jsonify({"verified": True, "user": user_row_to_dict(row)})
    raw_code = _json_body().get("code")
    code = re.sub(r"\D", "", raw_code if isinstance(raw_code, str) else (str(raw_code) if isinstance(raw_code, int) else ""))[:6]
    if len(code) != 6:
        return jsonify({"error": "Enter the 6-digit code from the email."}), 400
    conn = get_db()
    conn.autocommit = False
    changed = False
    try:
        with conn.cursor() as cur:
            # The address is read and locked here, so a change made at the same moment can't slip through.
            cur.execute("SELECT email, email_verified, pending_email FROM users WHERE id = %s FOR UPDATE", (row["id"],))
            now_row = cur.fetchone()
            if not now_row:
                conn.rollback()
                return jsonify({"error": "This account no longer exists."}), 404
            signup = now_row["email_verified"] is False
            target = now_row["email"] if signup else now_row["pending_email"]
            if not target:
                conn.rollback()
                return jsonify({"verified": True, "user": user_row_to_dict(row)})
            cur.execute("SELECT * FROM email_codes WHERE user_id = %s FOR UPDATE", (row["id"],))
            ec = cur.fetchone()
            if not ec or ec["expires_at"] <= datetime.now(timezone.utc):
                conn.rollback()
                return jsonify({"error": "That code has expired — tap “Send a new code”.", "expired": True}), 400
            if ec["attempts"] >= EMAIL_CODE_MAX_TRIES:
                conn.rollback()
                return jsonify({"error": "Too many tries with that code — tap “Send a new code”.", "expired": True}), 429
            if not ec["code_hash"] or not hmac.compare_digest(ec["code_hash"], _code_hash(row["id"], code, target)):
                cur.execute("UPDATE email_codes SET attempts = attempts + 1 WHERE user_id = %s RETURNING attempts", (row["id"],))
                left = max(0, EMAIL_CODE_MAX_TRIES - cur.fetchone()["attempts"])
                conn.commit()
                more = f" {left} {'try' if left == 1 else 'tries'} left." if left else " Tap “Send a new code”."
                return jsonify({"error": "That code isn't right — check the email and try again." + more, "tries_left": left}), 400
            if signup:
                cur.execute("UPDATE users SET email_verified = TRUE, email_proven = TRUE WHERE id = %s RETURNING *", (row["id"],))
            else:
                try:
                    cur.execute("SAVEPOINT move_email")
                    _release_unconfirmed(cur, target, row["id"])
                    cur.execute("UPDATE users SET email = pending_email, pending_email = NULL, email_verified = TRUE, email_proven = TRUE "
                                "WHERE id = %s RETURNING *", (row["id"],))
                    changed = True
                except psycopg2.errors.UniqueViolation:
                    cur.execute("ROLLBACK TO SAVEPOINT move_email")
                    cur.execute("UPDATE users SET pending_email = NULL WHERE id = %s", (row["id"],))
                    cur.execute("DELETE FROM email_codes WHERE user_id = %s", (row["id"],))
                    conn.commit()
                    return jsonify({"error": "That address now belongs to another Academic-Home account, so your email wasn't changed.",
                                    "taken": True}), 409
            user = cur.fetchone()
            cur.execute("DELETE FROM email_codes WHERE user_id = %s", (row["id"],))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
    return jsonify({"verified": True, "changed": changed, "user": user_row_to_dict(user)})


@app.route("/api/auth/resend-code", methods=["POST", "OPTIONS"])
@limiter.limit("12 per hour", key_func=_account_or_ip_key)
@auth_required
def resend_code():
    row = request.current_user
    if row.get("email_verified", True) is not False:
        pending = row.get("pending_email")
        if not pending:
            return jsonify({"verified": True, "sent": False})
        if not _email_ready():
            return jsonify({"error": "We can't send emails right now — please try again in a little while."}), 503
        sent, wait, err = _send_verification_code(row, to=pending)
        if not sent:
            return jsonify({"error": err, "retry_after": wait}), (429 if wait else 503)
        return jsonify({"sent": True, "retry_after": wait, "email": pending, "pending": True})
    if not _email_ready():
        # Email stopped working on the server: don't keep anyone waiting for a code.
        return jsonify({"verified": True, "sent": False, "user": user_row_to_dict(row)})
    sent, wait, err = _send_verification_code(row)
    if not sent:
        return jsonify({"error": err, "retry_after": wait}), (429 if wait else 503)
    return jsonify({"sent": True, "retry_after": wait, "email": row["email"]})


@app.route("/api/auth/cancel-email-change", methods=["POST", "OPTIONS"])
@auth_required
def cancel_email_change():
    """Keep the current address: the waiting new one and its code are dropped."""
    row = request.current_user
    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("UPDATE users SET pending_email = NULL WHERE id = %s RETURNING *", (row["id"],))
            new_row = cur.fetchone()
            if row.get("email_verified", True) is not False:
                _cancel_code(cur, row["id"])
    finally:
        conn.close()
    return jsonify({"user": user_row_to_dict(new_row)})


@app.route("/api/auth/change-email", methods=["POST", "OPTIONS"])
@limiter.limit("6 per hour", key_func=_account_or_ip_key)
@auth_required
def change_email():
    """A typo in the sign-up address: fix it before confirming, and get a code there."""
    row = request.current_user
    if row.get("email_verified", True) is not False:
        return jsonify({"error": "Your email is already confirmed — change it from your profile."}), 400
    email = _text(_json_body().get("email")).strip().lower()
    if not _valid_new_email(email):
        return jsonify({"error": "That doesn't look like a valid email"}), 400
    conn = get_db()
    try:
        with conn.cursor() as cur:
            try:
                # Only while the account is still unconfirmed — decided by the database at this moment.
                cur.execute("UPDATE users SET email = %s, pending_email = NULL WHERE id = %s AND email_verified = FALSE RETURNING *",
                            (email, row["id"]))
                new_row = cur.fetchone()
                if not new_row:
                    return jsonify({"error": "Your email is already confirmed — change it from your profile."}), 400
                _cancel_code(cur, row["id"])
            except psycopg2.errors.UniqueViolation:
                conn.rollback()
                return jsonify({"error": "An account with that email already exists."}), 409
    finally:
        conn.close()
    if not _email_ready():
        return jsonify({"user": user_row_to_dict(new_row), "sent": False, "verified": not _needs_verification(new_row)})
    sent, wait, err = _send_verification_code(new_row, force=True)
    return jsonify({"user": user_row_to_dict(new_row), "sent": sent, "error": err, "retry_after": wait})


@app.route("/api/auth/me", methods=["GET", "OPTIONS"])
@auth_required
def me():
    row = request.current_user
    if _needs_verification(row) and not _active_code(row["id"]):
        sent, _, _ = _send_verification_code(row)
        return jsonify({"user": user_row_to_dict(row), "verify": _verify_info(row, sent)})
    return jsonify({"user": user_row_to_dict(row)})


@app.route("/api/auth/update-profile", methods=["POST", "OPTIONS"])
@limiter.limit("20 per hour", key_func=_account_or_ip_key)
@auth_required
def update_profile():
    data = _json_body()
    name = _text(data.get("name")).strip()
    email = _text(data.get("email")).strip().lower()
    me_row = request.current_user
    if not name or not email:
        return jsonify({"error": "Name and email are required"}), 400
    if len(name) > 200:
        return jsonify({"error": "That name is too long (200 characters max)"}), 400
    if len(email) > 254:
        return jsonify({"error": "That email is too long"}), 400
    current = (me_row.get("email") or "").strip().lower()
    changed = email != current
    if not EMAIL_RE.match(email) or (changed and not _valid_new_email(email)):
        return jsonify({"error": "That doesn't look like a valid email"}), 400

    unconfirmed = me_row.get("email_verified", True) is False
    # A confirmed account's new address waits in pending_email until its code is entered —
    # the current address keeps working meanwhile, and nobody else's address can be held.
    pending = changed and not unconfirmed and _email_configured()
    conn = get_db()
    try:
        with conn.cursor() as cur:
            if changed:
                cur.execute("SELECT id FROM users WHERE email = %s AND id != %s AND email_verified IS NOT FALSE", (email, me_row["id"]))
                if cur.fetchone():
                    return jsonify({"error": "That email is already in use"}), 409
            try:
                if pending:
                    cur.execute("UPDATE users SET name = %s, pending_email = %s WHERE id = %s RETURNING *", (name, email, me_row["id"]))
                    row = cur.fetchone()
                    _cancel_code(cur, me_row["id"])
                elif changed and unconfirmed:
                    # Still confirming the sign-up address: the same as fixing a typo on the code screen.
                    cur.execute("UPDATE users SET name = %s, email = %s, pending_email = NULL WHERE id = %s AND email_verified = FALSE RETURNING *",
                                (name, email, me_row["id"]))
                    row = cur.fetchone()
                    if not row:
                        return jsonify({"error": "Please try again."}), 409
                    _cancel_code(cur, me_row["id"])
                elif changed:
                    # Email isn't set up on the server, so there's no way to confirm: change it straight away.
                    cur.execute("UPDATE users SET name = %s, email = %s, pending_email = NULL, email_proven = FALSE WHERE id = %s RETURNING *",
                                (name, email, me_row["id"]))
                    row = cur.fetchone()
                else:
                    cur.execute("UPDATE users SET name = %s WHERE id = %s RETURNING *", (name, me_row["id"]))
                    row = cur.fetchone()
            except psycopg2.errors.UniqueViolation:
                # Another account claimed this exact email between the check above and this update.
                conn.rollback()
                return jsonify({"error": "That email is already in use"}), 409
    finally:
        conn.close()

    if pending:
        if not _email_ready():
            return jsonify({"user": user_row_to_dict(row), "verify": {"email": email, "pending": True, "sent": False,
                            "error": "We can't send emails right now — try “Send a new code” in a little while."}})
        sent, wait, err = _send_verification_code(row, force=True, to=email)
        return jsonify({"user": user_row_to_dict(row), "verify": {"email": email, "pending": True, "sent": sent, "error": err, "retry_after": wait}})
    if changed and unconfirmed and _email_ready():
        sent, wait, err = _send_verification_code(row, force=True)
        return jsonify({"user": user_row_to_dict(row), "verify": {"email": row["email"], "sent": sent, "error": err, "retry_after": wait}})
    return jsonify({"user": user_row_to_dict(row)})


# ---------- plans / promo ----------


class PurchaseVerificationError(Exception):
    """Raised for any reason a submitted purchase can't be trusted —
    wrong app, expired, refunded, or the platform's own API rejected
    it — so plan_upgrade has exactly one place to catch and turn into
    a clean error response, rather than three different failure shapes
    from two different platforms' verification functions."""
    pass


class _VerifiedPlan(str):
    """A verified plan name that also carries what the store said about it:
    when the paid period ends, and (Apple) the subscription's original
    transaction id, which stays the same across renewals."""
    expires_at = None
    original_id = None


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
    verified = _VerifiedPlan(STORE_PRODUCT_TO_PLAN[product_id])
    verified.expires_at = datetime.fromtimestamp(expires_ms / 1000, timezone.utc) if expires_ms else None
    verified.original_id = str(payload.get("originalTransactionId") or "") or None
    return verified


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
    verified = _VerifiedPlan(STORE_PRODUCT_TO_PLAN[product_id])
    if is_subscription:
        ends = []
        for item in (data.get("lineItems") or []):
            try:
                ends.append(datetime.fromisoformat(str(item.get("expiryTime", "")).replace("Z", "+00:00")))
            except ValueError:
                pass
        verified.expires_at = max(ends) if ends else None
    return verified


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

    transaction_key = None
    if platform == "free":
        plan = "free"
    elif platform == "ios":
        transaction_id = (data.get("transaction_id") or "").strip()
        if not transaction_id:
            return jsonify({"error": "Missing transaction_id."}), 400
        try:
            plan = _apple_verify_transaction(transaction_id)
            # Keyed by the original transaction id, which every renewal shares —
            # so a renewal can't be claimed again as if it were a new purchase.
            transaction_key = "ios:" + (getattr(plan, "original_id", None) or transaction_id)
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
            transaction_key = "android:" + purchase_token
        except PurchaseVerificationError as e:
            return jsonify({"error": str(e)}), 402
    else:
        return jsonify({"error": "platform must be 'free', 'ios', or 'android'."}), 400

    # Store plans end when the store's paid period ends (plus a day's grace);
    # renewals move the date on (notices or a fresh restore), so cancelling or
    # a refund never leaves Pro running for free.
    store_end = getattr(plan, "expires_at", None)
    plan_until = (store_end + timedelta(days=1)) if store_end else None
    plan = str(plan)
    bonus_applied = False
    conn = get_db()
    try:
        with conn.cursor() as cur:
            if transaction_key and plan in PRO_PLANS:
                # One purchase, one account: if another account held this
                # subscription, it moves here and leaves that account.
                cur.execute("SELECT user_id, product FROM store_purchases WHERE transaction_key = %s", (transaction_key,))
                holder = cur.fetchone()
                if holder and holder["user_id"] and holder["user_id"] != request.current_user["id"]:
                    cur.execute("UPDATE users SET plan = 'free', plan_expires_at = NULL WHERE id = %s AND plan = %s",
                                (holder["user_id"], holder["product"]))
                    logger.info(f"Store purchase {transaction_key} moved from user {holder['user_id']} to {request.current_user['id']}")
            if plan == "payperuse":
                # One full use of Pro: Remy messages that never expire, on top
                # of the monthly allowance; the plan itself doesn't change.
                # A transaction already credited (a retried or replayed
                # request) adds nothing the second time.
                cur.execute(
                    "INSERT INTO store_purchases (transaction_key, user_id, product) VALUES (%s, %s, %s) "
                    "ON CONFLICT (transaction_key) DO NOTHING RETURNING transaction_key",
                    (transaction_key, request.current_user["id"], "payperuse"),
                )
                if cur.fetchone():
                    cur.execute("UPDATE users SET remy_credits = COALESCE(remy_credits, 0) + %s WHERE id = %s RETURNING *",
                                (PAYPERUSE_MESSAGES, request.current_user["id"]))
                else:
                    cur.execute("SELECT * FROM users WHERE id = %s", (request.current_user["id"],))
            else:
                cur.execute("UPDATE users SET plan = %s, plan_expires_at = %s, plan_source = 'store' WHERE id = %s RETURNING *",
                            (plan, plan_until, request.current_user["id"]))
            row = cur.fetchone()
            if transaction_key and plan in PRO_PLANS:
                # Which account owns this subscription: the stores' renewal,
                # cancellation and refund notices find the account by it.
                cur.execute(
                    "INSERT INTO store_purchases (transaction_key, user_id, product) VALUES (%s, %s, %s) "
                    "ON CONFLICT (transaction_key) DO UPDATE SET user_id = EXCLUDED.user_id, product = EXCLUDED.product",
                    (transaction_key, request.current_user["id"], plan),
                )

            if plan in PRO_PLANS and _reward_referral(cur, row["id"], referral_code):
                cur.execute("SELECT * FROM users WHERE id = %s", (row["id"],))
                row = cur.fetchone()
                bonus_applied = True
    finally:
        conn.close()

    return jsonify({"user": user_row_to_dict(row), "result": {"referral_bonus_applied": bonus_applied}})


# ------------------------------------------------------------ Paystack (web only)
# On the web (getdocente.vercel.app) Pro is bought with Paystack — card or
# mobile money, charged in Ghana cedis — as a one-time pass: 30 or 365 days of
# Pro, or 30 Remy messages. The App Store and Google Play apps must use the
# stores' own purchases (their rules), so requests from the store apps are
# refused here and the app hides Paystack there too. The secret key is read
# only from the server's environment — never from code or the app.
PAYSTACK_SECRET_KEY = os.environ.get("PAYSTACK_SECRET_KEY", "").strip()
PAYSTACK_PUBLIC_KEY = os.environ.get("PAYSTACK_PUBLIC_KEY", "").strip()
PAYSTACK_API = "https://api.paystack.co"
PAYSTACK_PLANS = {
    # plan: (price in pesewas, days of Pro, Remy messages added)
    "monthly": (10000, 30, 0),
    "halfyear": (40000, 182, 0),   # 3.2: 6 months — a whole semester, for about the price of 4 months
    "payperuse": (3000, 0, 30),    # the Remy Pack
    "yearly": (30000, 365, 0),     # no longer sold (see PAYSTACK_ON_SALE); kept so an older payment still applies
}
PAYSTACK_ON_SALE = ("halfyear", "monthly", "payperuse")


def _paystack_on():
    return bool(PAYSTACK_SECRET_KEY and PAYSTACK_PUBLIC_KEY)


def _from_store_app():
    ua = request.headers.get("User-Agent", "")
    return "DocenteApp" in ua or request.headers.get("X-Docente-Store-App") == "1"


def _paystack_call(method, path, payload=None, timeout=20):
    resp = requests.request(method, PAYSTACK_API + path, json=payload, timeout=timeout,
                            headers={"Authorization": f"Bearer {PAYSTACK_SECRET_KEY}", "Content-Type": "application/json"})
    try:
        body = resp.json()
    except ValueError:
        body = {}
    if resp.status_code >= 400 or not body.get("status"):
        raise ConversionError(body.get("message") or f"Paystack answered {resp.status_code}")
    return body.get("data") or {}


def _paystack_apply(reference, paid_amount, currency):
    """Give the person what they paid for — exactly once per payment, however
    many times it's confirmed (the browser coming back, and the webhook)."""
    conn = get_db()
    conn.autocommit = False   # hold the row lock until the commit
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM paystack_payments WHERE reference = %s FOR UPDATE", (reference,))
            pay = cur.fetchone()
            if not pay:
                conn.rollback()
                return None, "unknown"
            if pay["applied_at"]:
                conn.rollback()
                return pay["user_id"], "already"
            if str(currency).upper() != "GHS" or int(paid_amount or 0) < int(pay["amount"]):
                cur.execute("UPDATE paystack_payments SET status = 'mismatch' WHERE reference = %s", (reference,))
                conn.commit()
                logger.warning(f"Paystack {reference}: paid {paid_amount} {currency}, expected {pay['amount']} GHS — not applied")
                return pay["user_id"], "mismatch"
            _, days, credits = PAYSTACK_PLANS[pay["plan"]]
            if days:
                # Bought during the free-Pro promotion: the paid days start when it ends.
                start = _promo_until() if _promo_active() else datetime.now(timezone.utc)
                cur.execute(
                    "UPDATE users SET plan = %s, plan_expires_at = GREATEST(COALESCE(plan_expires_at, now()), now(), %s::timestamptz) + make_interval(days => %s), plan_source = 'paystack' WHERE id = %s",
                    (pay["plan"], start, days, pay["user_id"]))
                if pay["plan"] in PRO_PLANS and _reward_referral(cur, pay["user_id"], pay.get("referral_code")):
                    cur.execute("UPDATE paystack_payments SET referral_applied = TRUE WHERE reference = %s", (reference,))
            if credits:
                cur.execute("UPDATE users SET remy_credits = COALESCE(remy_credits, 0) + %s WHERE id = %s", (credits, pay["user_id"]))
            cur.execute("UPDATE paystack_payments SET status = 'success', applied_at = now() WHERE reference = %s", (reference,))
        conn.commit()
        return pay["user_id"], "applied"
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


@app.route("/api/paystack/config", methods=["GET", "OPTIONS"])
def paystack_config():
    return jsonify({
        "enabled": _paystack_on(), "public_key": PAYSTACK_PUBLIC_KEY if _paystack_on() else None, "currency": "GHS",
        # live or test keys — never the key itself — so the setup can be checked at a glance
        "mode": ("live" if PAYSTACK_SECRET_KEY.startswith("sk_live") else "test") if _paystack_on() else None,
        "plans": {k: {"amount_ghs": v[0] / 100, "days": v[1], "credits": v[2]} for k, v in PAYSTACK_PLANS.items() if k in PAYSTACK_ON_SALE},
    })


@app.route("/api/paystack/initialize", methods=["POST", "OPTIONS"])
@limiter.limit("20 per hour", key_func=_account_or_ip_key)
@auth_required
def paystack_initialize():
    if not _paystack_on():
        return jsonify({"error": "Payments aren't set up yet — please try again later."}), 503
    if _from_store_app():
        return jsonify({"error": "In the Academic-Home app, Pro is bought through the App Store or Google Play."}), 403
    plan = str((request.get_json(silent=True) or {}).get("plan") or "")
    if plan not in PAYSTACK_ON_SALE:
        return jsonify({"error": "Choose 6 months or 1 month of Pro, or a Remy Pack."}), 400
    user = request.current_user
    if plan in PRO_PLANS and user.get("plan") in PRO_PLANS and _owns_store_plan(user["id"]):
        return jsonify({"error": "You already have Pro through the App Store or Google Play — manage it there."}), 409
    friend_code = _clean_code((request.get_json(silent=True) or {}).get("referral_code"))
    amount = PAYSTACK_PLANS[plan][0]
    reference = f"dct_{user['id']}_{secrets.token_hex(6)}"
    conn = get_db()
    try:
        with conn.cursor() as cur:
            if friend_code and plan in PRO_PLANS:
                if not _find_referrer(cur, friend_code, user["id"]):
                    return jsonify({"error": "That referral code wasn't found (or it's your own) — check it, or leave it empty.", "field": "referral"}), 400
            cur.execute("INSERT INTO paystack_payments (reference, user_id, plan, amount, referral_code) VALUES (%s, %s, %s, %s, %s)",
                        (reference, user["id"], plan, amount, friend_code or None))
    finally:
        conn.close()
    try:
        data = _paystack_call("POST", "/transaction/initialize", {
            "email": user["email"], "amount": amount, "currency": "GHS", "reference": reference,
            "callback_url": f"{FRONTEND_URL.rstrip('/')}/?paystack=return",
            "metadata": {"user_id": user["id"], "plan": plan,
                         "custom_fields": [{"display_name": "Academic-Home plan", "variable_name": "plan", "value": plan}]},
        })
    except Exception as e:
        logger.error(f"Paystack couldn't start payment {reference}: {e} \u2014 if this repeats, check PAYSTACK_SECRET_KEY "
                     "and the Paystack account's settings (GHS payments switched on).")
        return jsonify({"error": "Paystack couldn't start the payment just now. Please try again in a moment."}), 502
    return jsonify({"authorization_url": data.get("authorization_url"), "access_code": data.get("access_code"), "reference": reference})


@app.route("/api/paystack/verify", methods=["POST", "OPTIONS"])
@limiter.limit("60 per hour", key_func=_account_or_ip_key)
@auth_required
def paystack_verify():
    if not _paystack_on():
        return jsonify({"error": "Payments aren't set up yet."}), 503
    reference = str((request.get_json(silent=True) or {}).get("reference") or "").strip()[:80]
    user = request.current_user
    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT user_id FROM paystack_payments WHERE reference = %s", (reference,))
            pay = cur.fetchone()
    finally:
        conn.close()
    if not pay or pay["user_id"] != user["id"]:
        return jsonify({"error": "That payment wasn't found on your account."}), 404
    try:
        data = _paystack_call("GET", f"/transaction/verify/{reference}")
    except Exception as e:
        logger.warning(f"Paystack verify failed for {reference}: {e}")
        return jsonify({"error": "Couldn't check the payment with Paystack — please try again in a moment."}), 502
    status = data.get("status")
    if status != "success" or data.get("reference") != reference:
        # Paystack's own reason ("Declined", "Insufficient funds", …) and the
        # method tried — shown to the student and logged (no personal details).
        reason = str(data.get("gateway_response") or "")[:140]
        if status in ("failed", "abandoned", "reversed"):
            logger.warning(f"Paystack: {reference} {status} | channel: {data.get('channel') or '-'} | Paystack says: {reason or '-'}")
        return jsonify({"status": status or "unknown", "reason": reason, "channel": data.get("channel")})
    _, outcome = _paystack_apply(reference, data.get("amount"), data.get("currency"))
    if outcome == "mismatch":
        return jsonify({"status": "mismatch", "error": "The amount paid didn't match the plan — contact support.docente@gmail.com and we'll sort it out."}), 409
    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM users WHERE id = %s", (user["id"],))
            row = cur.fetchone()
            cur.execute("SELECT referral_applied FROM paystack_payments WHERE reference = %s", (reference,))
            bonus = bool((cur.fetchone() or {}).get("referral_applied"))
    finally:
        conn.close()
    return jsonify({"status": "success", "user": user_row_to_dict(row), "referral_bonus_applied": bonus})


@app.route("/api/paystack/webhook", methods=["POST"])
def paystack_webhook():
    """Paystack's own confirmation, signed with the secret key — it covers
    anyone who pays and closes the page before coming back."""
    if not _paystack_on():
        return jsonify({"ok": False}), 503
    raw = request.get_data() or b""
    expected = hmac.new(PAYSTACK_SECRET_KEY.encode(), raw, hashlib.sha512).hexdigest()
    if not hmac.compare_digest(expected, request.headers.get("x-paystack-signature", "")):
        return jsonify({"error": "bad signature"}), 401
    event = request.get_json(silent=True) or {}
    if event.get("event") == "charge.success":
        data = event.get("data") or {}
        try:
            _paystack_apply(str(data.get("reference") or ""), data.get("amount"), data.get("currency"))
        except Exception as e:
            logger.warning(f"Paystack webhook apply failed: {e}")
            return jsonify({"ok": False}), 500
    return jsonify({"ok": True})


@app.route("/api/paystack/payments", methods=["GET", "OPTIONS"])
@auth_required
def paystack_payments():
    """The person's last ten web payments, for the Payment history sheet."""
    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT reference, plan, amount, status, created_at FROM paystack_payments WHERE user_id = %s ORDER BY created_at DESC LIMIT 10",
                        (request.current_user["id"],))
            rows = cur.fetchall()
    finally:
        conn.close()
    return jsonify({"payments": [{
        "reference": r["reference"], "plan": r["plan"], "amount_ghs": r["amount"] / 100, "status": r["status"],
        "created_at": r["created_at"].isoformat() if r["created_at"] else None} for r in rows]})


@app.route("/api/paystack/sync", methods=["POST", "OPTIONS"])
@limiter.limit("30 per hour", key_func=_account_or_ip_key)
@auth_required
def paystack_sync():
    """Checks this person's unfinished payments (from the last 3 days) with
    Paystack and gives them what they paid for — so a payment still counts
    when they closed the page before coming back, even before the webhook is
    set up. Each payment is applied at most once."""
    user = request.current_user
    applied = 0
    if _paystack_on():
        conn = get_db()
        try:
            with conn.cursor() as cur:
                cur.execute("SELECT reference FROM paystack_payments WHERE user_id = %s AND status = 'pending' AND created_at > now() - interval '3 days' ORDER BY created_at DESC LIMIT 3",
                            (user["id"],))
                refs = [r["reference"] for r in cur.fetchall()]
        finally:
            conn.close()
        for ref in refs:
            try:
                data = _paystack_call("GET", f"/transaction/verify/{ref}", timeout=8)
            except Exception as e:
                logger.warning(f"Paystack sync couldn't check {ref}: {e}")
                continue
            if data.get("status") == "success" and data.get("reference") == ref:
                _, outcome = _paystack_apply(ref, data.get("amount"), data.get("currency"))
                if outcome == "applied":
                    applied += 1
    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM users WHERE id = %s", (user["id"],))
            row = cur.fetchone()
            cur.execute("SELECT COUNT(*) AS n FROM paystack_payments WHERE user_id = %s AND status = 'pending' AND created_at > now() - interval '3 days'", (user["id"],))
            pending = (cur.fetchone() or {}).get("n", 0)
    finally:
        conn.close()
    return jsonify({"applied": applied, "pending": int(pending or 0), "user": user_row_to_dict(row)})


# ---- Paystack: the server confirms payments itself ----
# A payment used to be confirmed only when the student's browser came back
# from Paystack (or by Paystack's webhook, if set up). If that return got lost,
# a payment could be taken and never credited. So the server also checks
# unfinished payments with Paystack every few minutes and credits any that
# were paid (exactly once — _paystack_apply guards that). Payments Paystack
# reports as abandoned or failed are marked so, once they're 6 hours old.
def _paystack_reconcile_once():
    if not _paystack_on():
        return 0
    credited = 0
    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT pg_try_advisory_lock(735001) AS ok")      # one server copy at a time
            if not (cur.fetchone() or {}).get("ok"):
                return 0
            try:
                cur.execute("SELECT reference, created_at FROM paystack_payments WHERE status = 'pending' "
                            "AND created_at > now() - interval '3 days' AND created_at < now() - interval '90 seconds' "
                            "ORDER BY created_at DESC LIMIT 20")
                rows = cur.fetchall()
                for r in rows:
                    try:
                        data = _paystack_call("GET", f"/transaction/verify/{r['reference']}", timeout=8)
                    except Exception as e:
                        logger.warning(f"Paystack check couldn't reach {r['reference']}: {e}")
                        continue
                    st = data.get("status")
                    if st == "success" and data.get("reference") == r["reference"]:
                        _, outcome = _paystack_apply(r["reference"], data.get("amount"), data.get("currency"))
                        logger.info(f"Paystack check: {r['reference']} was paid -> {outcome}")
                        if outcome == "applied":
                            credited += 1
                    elif st in ("abandoned", "failed", "reversed") and r["created_at"] < datetime.now(timezone.utc) - timedelta(hours=6):
                        logger.warning(f"Paystack check: {r['reference']} {st} | channel: {data.get('channel') or '-'} | Paystack says: {data.get('gateway_response') or '-'}")
                        cur.execute("UPDATE paystack_payments SET status = %s WHERE reference = %s AND status = 'pending'",
                                    ("abandoned" if st == "abandoned" else "failed", r["reference"]))
                        conn.commit()
            finally:
                cur.execute("SELECT pg_advisory_unlock(735001)")
    finally:
        conn.close()
    return credited


def _paystack_explain_recent():
    """Once at start-up: log why recent payments didn't go through — the
    method tried and Paystack's own reason, never any personal details.
    The lock is held for the WHOLE explanation (and briefly after), so the
    other server worker, starting at the same moment, can't repeat it.
    (Releasing it straight after reading the list let both workers log
    every reason twice.)"""
    if not _paystack_on():
        return
    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT pg_try_advisory_lock(735002) AS ok")
            if not (cur.fetchone() or {}).get("ok"):
                return
            cur.execute("SELECT reference, plan FROM paystack_payments WHERE status IN ('failed', 'abandoned', 'pending') "
                        "AND created_at > now() - interval '3 days' ORDER BY created_at DESC LIMIT 12")
            rows = cur.fetchall()
            for r in rows:
                try:
                    d = _paystack_call("GET", f"/transaction/verify/{r['reference']}", timeout=8)
                except Exception as e:
                    logger.warning(f"Paystack reason: {r['reference']} ({r['plan']}) -> Paystack couldn't say: {e}")
                    continue
                logger.warning(f"Paystack reason: {r['reference']} ({r['plan']}) -> {d.get('status')} | channel: {d.get('channel') or '-'} | Paystack says: {d.get('gateway_response') or '-'}")
            time.sleep(20)       # keep the lock a little longer, past the other worker's start
    finally:
        conn.close()             # closing the connection releases the lock


def _paystack_reconcile_loop():
    time.sleep(45)
    try:
        _paystack_explain_recent()
    except Exception as e:
        logger.warning(f"Paystack reason check failed: {e}")
    while True:
        try:
            _paystack_reconcile_once()
        except Exception as e:
            logger.warning(f"Paystack check pass failed: {e}")
        time.sleep(300)


if _paystack_on() and "fake" not in os.environ.get("DATABASE_URL", "fake"):
    _threading.Thread(target=_paystack_reconcile_loop, daemon=True, name="paystack-check").start()


# ------------------------------------------------------------ Remy's memory
# What Remy remembers about a person is theirs to see and control: list it,
# add a note, edit or delete one, clear everything, or switch memory off
# (then nothing is read or saved). Every change is limited to the person's
# own notes.
def _memory_state(cur, user_id, enabled):
    cur.execute("SELECT id, note, created_at FROM user_memory WHERE user_id = %s ORDER BY created_at DESC, id DESC LIMIT %s",
                (user_id, MEMORY_LIMIT))
    notes = [{"id": r["id"], "note": r["note"], "created_at": r["created_at"].isoformat() if r["created_at"] else None} for r in cur.fetchall()]
    return {"enabled": enabled, "notes": notes, "limit": MEMORY_LIMIT}


@app.route("/api/memory", methods=["GET", "POST", "DELETE", "OPTIONS"])
@auth_required
def memory_endpoint():
    user = request.current_user
    conn = get_db()
    try:
        with conn.cursor() as cur:
            if request.method == "POST":
                note = re.sub(r"\s+", " ", str((request.get_json(silent=True) or {}).get("note") or "")).strip()[:300]
                if not note:
                    return jsonify({"error": "Write something for Remy to remember."}), 400
                cur.execute("SELECT note FROM user_memory WHERE user_id = %s", (user["id"],))
                existing = [r["note"].strip().lower() for r in cur.fetchall()]
                if len(existing) >= MEMORY_LIMIT:
                    return jsonify({"error": f"Remy can remember up to {MEMORY_LIMIT} notes — delete one first."}), 400
                if note.lower() not in existing:
                    cur.execute("INSERT INTO user_memory (user_id, note) VALUES (%s, %s)", (user["id"], note))
            elif request.method == "DELETE":
                cur.execute("DELETE FROM user_memory WHERE user_id = %s", (user["id"],))
            return jsonify(_memory_state(cur, user["id"], _memory_on(user)))
    finally:
        conn.close()


@app.route("/api/memory/<int:note_id>", methods=["PATCH", "DELETE", "OPTIONS"])
@auth_required
def memory_note_endpoint(note_id):
    user = request.current_user
    conn = get_db()
    try:
        with conn.cursor() as cur:
            if request.method == "PATCH":
                note = re.sub(r"\s+", " ", str((request.get_json(silent=True) or {}).get("note") or "")).strip()[:300]
                if not note:
                    return jsonify({"error": "A note can't be empty — delete it instead."}), 400
                cur.execute("UPDATE user_memory SET note = %s WHERE id = %s AND user_id = %s RETURNING id", (note, note_id, user["id"]))
            else:
                cur.execute("DELETE FROM user_memory WHERE id = %s AND user_id = %s RETURNING id", (note_id, user["id"]))
            if not cur.fetchone():
                return jsonify({"error": "That note wasn't found."}), 404
            return jsonify(_memory_state(cur, user["id"], _memory_on(user)))
    finally:
        conn.close()


@app.route("/api/memory/settings", methods=["PUT", "OPTIONS"])
@auth_required
def memory_settings_endpoint():
    enabled = bool((request.get_json(silent=True) or {}).get("enabled"))
    user = request.current_user
    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("UPDATE users SET memory_enabled = %s WHERE id = %s", (enabled, user["id"]))
            return jsonify(_memory_state(cur, user["id"], enabled))
    finally:
        conn.close()


# ------------------------------------------------------------ referral codes
# A person can make one referral code a month; each lasts 14 days. When a
# friend subscribes with it, both get a bonus month — the same reward as a
# personal referral code. Kept in promo_codes (created_by = the referrer).
# The old /api/promo/generate address does the same, for older app versions.
REFERRAL_CODE_DAYS = 14
REFERRAL_CODE_EVERY_DAYS = 30


def _referral_state(last):
    now = datetime.now(timezone.utc)
    if not last:
        return {"code": None, "expires_at": None, "next_allowed_at": None, "can_generate": True}
    next_at = last["created_at"] + timedelta(days=REFERRAL_CODE_EVERY_DAYS)
    active = last["expires_at"] > now
    return {
        "code": last["code"] if active else None,
        "expires_at": last["expires_at"].isoformat() if active else None,
        "next_allowed_at": next_at.isoformat() if next_at > now else None,
        "can_generate": next_at <= now,
    }


def _latest_referral(cur, user_id):
    cur.execute("SELECT code, expires_at, created_at FROM promo_codes WHERE created_by = %s ORDER BY created_at DESC LIMIT 1", (user_id,))
    return cur.fetchone()


@app.route("/api/referral", methods=["GET", "OPTIONS"])
@auth_required
def referral_current():
    conn = get_db()
    try:
        with conn.cursor() as cur:
            return jsonify(_referral_state(_latest_referral(cur, request.current_user["id"])))
    finally:
        conn.close()


@app.route("/api/referral/check", methods=["POST", "OPTIONS"])
@limiter.limit("40 per hour")
def referral_check():
    """Is this a real referral code? Used as it's typed (sign-up and checkout)."""
    if request.method == "OPTIONS":
        return "", 204
    code = _clean_code((request.get_json(silent=True) or {}).get("code"))
    if len(code) < 4:
        return jsonify({"valid": False, "error": "Enter the whole code."})
    header = request.headers.get("Authorization", "")
    uid = verify_token(header.split(" ", 1)[1]) if header.startswith("Bearer ") else None
    conn = get_db()
    try:
        with conn.cursor() as cur:
            ref = _find_referrer(cur, code)
            if not ref:
                return jsonify({"valid": False, "error": "That code wasn't found — check the spelling, or ask for a fresh code."})
            if uid and ref["id"] == uid:
                return jsonify({"valid": False, "error": "That's your own code — share it with friends instead."})
            if uid:
                cur.execute("SELECT referral_rewarded FROM users WHERE id = %s", (uid,))
                me = cur.fetchone()
                if me and me["referral_rewarded"]:
                    return jsonify({"valid": False, "error": "You've already had a referral bonus — it's once per account."})
    finally:
        conn.close()
    first = (str(ref.get("name") or "").strip().split() or ["a friend"])[0][:30]
    return jsonify({"valid": True, "from": first})


@app.route("/api/referral/generate", methods=["POST", "OPTIONS"])
@app.route("/api/promo/generate", methods=["POST", "OPTIONS"], endpoint="promo_generate_legacy")
@limiter.limit("10 per hour", key_func=_account_or_ip_key)
@auth_required
def referral_generate():
    user_id = request.current_user["id"]
    conn = get_db()
    # One transaction, so the row lock below is held until the commit —
    # get_db() connections commit every statement (and release the lock)
    # immediately otherwise.
    conn.autocommit = False
    try:
        with conn.cursor() as cur:
            # Lock the person's row so two quick taps can't both make a code.
            cur.execute("SELECT id FROM users WHERE id = %s FOR UPDATE", (user_id,))
            state = _referral_state(_latest_referral(cur, user_id))
            if not state["can_generate"]:
                conn.rollback()
                when = datetime.fromisoformat(state["next_allowed_at"]).strftime("%d %B %Y")
                return jsonify(dict(state, error=f"You can make one referral code a month — your next one is ready on {when}.")), 429
            now = datetime.now(timezone.utc)
            code, expires_at = generate_promo_code(), now + timedelta(days=REFERRAL_CODE_DAYS)
            cur.execute(
                "INSERT INTO promo_codes (code, created_by, discount_desc, expires_at) VALUES (%s, %s, %s, %s)",
                (code, user_id, "Referral", expires_at),
            )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
    return jsonify({"code": code, "expires_at": expires_at.isoformat(),
                    "next_allowed_at": (now + timedelta(days=REFERRAL_CODE_EVERY_DAYS)).isoformat(), "can_generate": False})


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


def _refund_sage_message(user_row, source="allowance"):
    """Returns one Remy message to the user's monthly allowance, for a
    request that was counted but never got a reply. Never goes below zero
    (the allowance may have reset in between). Best-effort: a failure here
    is logged, never shown to the user on top of the original error."""
    conn = None
    try:
        conn = get_db()
        with conn.cursor() as cur:
            if source == "credit":
                # It came from a Pay-per-Use purchase: back to the purchase.
                cur.execute("UPDATE users SET remy_credits = remy_credits + 1 WHERE id = %s", (user_row["id"],))
            else:
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


def _sage_status(row):
    """(messages used in the current window, the plan's monthly limit, when
    the window started) for a user row — computed, never written. The one
    rule shared by _log_sage_message_and_consume (which enforces the limit)
    and the meter shown to the person (which displays it), so the number on
    screen is always exactly what the next real message would find. A window
    that has run out counts as fresh here even before the database row is
    reset by the next message."""
    now = datetime.now(timezone.utc)
    used = row["sage_messages_used"]
    reset_at = row["sage_messages_reset_at"]
    if reset_at and now - reset_at > timedelta(days=SAGE_WINDOW_DAYS):
        used, reset_at = 0, now
    limit = SAGE_FREE_MESSAGES_LIMIT if row["plan"] == "free" else SAGE_PAID_MESSAGES_LIMIT
    return used, limit, reset_at


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
            row = _with_promo(cur.fetchone())
            used, limit, reset_at = _sage_status(row)
            source = "allowance"
            if used >= limit:
                if (row.get("remy_credits") or 0) > 0:
                    # The month's allowance is used up: a Pay-per-Use message.
                    source = "credit"
                    cur.execute("UPDATE users SET remy_credits = remy_credits - 1 WHERE id = %s", (row["id"],))
                else:
                    conn.rollback()
                    if row["plan"] == "free":
                        raise ConversionError(
                            f"You've used your {SAGE_FREE_MESSAGES_LIMIT} free messages with Remy this month. "
                            f"Get Pro for {SAGE_PAID_MESSAGES_LIMIT} a month, or a Remy Pack for {PAYPERUSE_MESSAGES} more."
                        )
                    raise ConversionError(
                        f"You've reached Remy's monthly limit ({SAGE_PAID_MESSAGES_LIMIT} messages) for this "
                        f"account. It resets next month — or a Remy Pack adds {PAYPERUSE_MESSAGES} more now."
                    )
            if source == "allowance":
                cur.execute(
                    "UPDATE users SET sage_messages_used = %s, sage_messages_reset_at = %s WHERE id = %s",
                    (used + 1, reset_at, row["id"]),
                )
        conn.commit()
        return source
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

    # Quick automatic fix into one of Academic-Home's templates: the deck is read
    # and rebuilt by the same engine as Slide Studio (slides, pictures, logo
    # and notes kept). Without this field the original tidy-up runs as before.
    studio_template = None
    if from_fmt == "pptx" and to_fmt == "pptx":
        studio_template = (request.form.get("studio_template") or "").strip() or None
        if studio_template and studio_template not in {t["id"] for t in template_catalog()}:
            return jsonify({"error": "Unknown template — please choose one from the list."}), 400

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
            if studio_template:
                deck = parse_pptx_to_deck(src_path)
                deck["template"] = studio_template
                result_path = build_template_deck(deck, work_dir, "fixed.pptx")
            else:
                result_path = convert(src_path, from_fmt, to_fmt, work_dir, style=style, pptx_options=pptx_options)
        except (ConversionError, StudioError) as e:
            return jsonify({"error": str(e)}), 422
        except FileNotFoundError as e:
            return jsonify({"error": f"Required conversion tool missing on server: {e}"}), 500

        if to_fmt == "pptx":
            try:
                _colour_not_bold(result_path)
            except Exception as e:   # a finished deck is never lost to this tidy-up
                logger.warning(f"Colour-not-bold pass skipped: {e}")

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


def _pptx_fix_form(form):
    """Fix slides' choices, read and checked once for one file or a batch:
    Keep my design (sizes, fonts, colours, heading alignment) or an Academic-Home
    template. Returns (pptx_options, studio_template, error)."""
    pptx_options = {}
    for size_field in ("title_size", "body_size"):
        raw = form.get(size_field)
        if raw:
            try:
                size_val = float(raw)
            except ValueError:
                return None, None, f"'{size_field}' must be a number"
            if not (6 <= size_val <= 200):
                return None, None, f"'{size_field}' must be between 6 and 200 points"
            pptx_options[size_field] = size_val
    for font_field in ("title_font", "body_font"):
        raw = (form.get(font_field) or "").strip()
        if raw:
            if len(raw) > 100:
                return None, None, f"'{font_field}' is too long"
            pptx_options[font_field] = raw
    for color_field in ("title_color", "body_color"):
        raw = (form.get(color_field) or "").strip().lstrip("#")
        if raw:
            if not re.fullmatch(r"[0-9A-Fa-f]{6}", raw):
                return None, None, f"'{color_field}' must be a 6-digit hex color"
            pptx_options[color_field] = raw
    title_align = (form.get("title_align") or "").strip().lower()
    if title_align:
        if title_align not in ("left", "center", "right"):
            return None, None, "'title_align' must be left, center, or right"
        pptx_options["title_align"] = title_align
    studio_template = (form.get("studio_template") or "").strip() or None
    if studio_template and studio_template not in {t["id"] for t in template_catalog()}:
        return None, None, "Unknown template — please choose one from the list."
    return pptx_options, studio_template, None


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
    pptx_options, studio_template = None, None
    if from_fmt == "pptx" and to_fmt == "pptx":
        pptx_options, studio_template, err = _pptx_fix_form(request.form)
        if err:
            return jsonify({"error": err}), 400

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

            # each file in its own folder: results with the same name never overwrite each other
            file_dir = tempfile.mkdtemp(dir=work_dir)
            src_path = os.path.join(file_dir, safe_upload_filename(filename)[0])
            file.save(src_path)
            try:
                if studio_template:
                    deck = parse_pptx_to_deck(src_path)
                    deck["template"] = studio_template
                    result_path = build_template_deck(deck, file_dir, "fixed.pptx")
                else:
                    result_path = convert(src_path, from_fmt, to_fmt, file_dir, style=style, pptx_options=pptx_options)
                if to_fmt == "pptx":
                    try:
                        _colour_not_bold(result_path)
                    except Exception as e:   # a finished deck is never lost to this tidy-up
                        logger.warning(f"Colour-not-bold pass skipped: {e}")
                base_name = safe_download_name(filename).rsplit(".", 1)[0]
                converted.append((result_path, f"{base_name}.{to_fmt}"))
                details.append({"filename": filename, "success": True})
            except (ConversionError, StudioError) as e:
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


def _is_pro(user):
    return (user or {}).get("plan") in PRO_PLANS


@app.route("/api/pdf/apply", methods=["POST", "OPTIONS"])
@limiter.limit(HEAVY_LIMIT, key_func=_account_or_ip_key)
@auth_required
def pdf_apply_endpoint():
    """3.1 — the PDF editor's Save: 'file' (the original PDF), 'plan' (JSON: pages, items, form,
    output — see pdf_apply_plan), 'assets' (pictures, signatures and added PDFs, each named by its
    own filename in the plan) and, for a locked PDF, 'password'. Returns the edited PDF — or, when
    the plan splits it (output.split), a zip of the parts."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    file = request.files["file"]
    ext = file.filename.rsplit(".", 1)[-1].lower() if "." in (file.filename or "") else ""
    if ext != "pdf":
        return jsonify({"error": "The file must be a PDF."}), 400
    try:
        plan = json.loads(request.form.get("plan") or "{}")
    except json.JSONDecodeError:
        return jsonify({"error": "The changes couldn't be read — please try again."}), 400
    if not isinstance(plan, dict):
        return jsonify({"error": "The changes couldn't be read — please try again."}), 400
    if pdf_plan_needs_pro(plan) and not _is_pro(request.current_user):
        return jsonify({"error": "Redacting and protecting PDFs are part of Academic-Home Pro.", "pro": True}), 403
    try:
        _log_conversion_and_consume(request.current_user, "pdf", "pdf-edit")
    except ConversionError as e:
        return jsonify({"error": str(e)}), 402
    work_dir = tempfile.mkdtemp(prefix=f"mc_pdfapply_{uuid.uuid4().hex[:8]}_")
    try:
        src_path = os.path.join(work_dir, "source.pdf")
        file.save(src_path)
        assets = {}
        for n, a in enumerate(request.files.getlist("assets")[:60]):
            if not a.filename:
                continue
            safe, _e = safe_upload_filename(a.filename)
            p = os.path.join(work_dir, f"asset{n}_{safe}")
            a.save(p)
            assets[a.filename] = p
        try:
            result, notes = pdf_apply_plan(src_path, plan, work_dir, assets, password=request.form.get("password") or None)
        except ConversionError as e:
            return jsonify({"error": str(e)}), 422
        base = safe_download_name(file.filename).rsplit(".", 1)[0]
        if isinstance(result, list):
            # split into files: they come back together in one zip, each named after the PDF
            zip_path = os.path.join(work_dir, "parts.zip")
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                for p in result:
                    zf.write(p, arcname=f"{base} - " + os.path.basename(p)[len("split-"):])
            resp = send_file(zip_path, mimetype="application/zip", as_attachment=True, download_name=f"{base} (split).zip")
        else:
            resp = send_file(result, mimetype="application/pdf", as_attachment=True, download_name=f"{base} (edited).pdf")
        if notes:
            resp.headers["X-Docente-Note"] = " ".join(notes)[:300]
        return resp
    except Exception as e:
        logger.error(f"PDF editor save failed unexpectedly: {e}", exc_info=True)
        return jsonify({"error": "Saving the PDF failed. Please try again."}), 500
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


@app.route("/api/pdf/to-images", methods=["POST", "OPTIONS"])
@limiter.limit(HEAVY_LIMIT, key_func=_account_or_ip_key)
@auth_required
def pdf_to_images_endpoint():
    """3.1 — PDF → pictures: one JPG per page (a single page comes back as the JPG itself, more as a zip)."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    file = request.files["file"]
    try:
        _log_conversion_and_consume(request.current_user, "pdf", "jpg")
    except ConversionError as e:
        return jsonify({"error": str(e)}), 402
    work_dir = tempfile.mkdtemp(prefix=f"mc_pdfimg_{uuid.uuid4().hex[:8]}_")
    try:
        src_path = os.path.join(work_dir, "source.pdf")
        file.save(src_path)
        try:
            paths = pdf_to_images(src_path, os.path.join(work_dir, "pages"), password=request.form.get("password") or None)
        except ConversionError as e:
            return jsonify({"error": str(e)}), 422
        base = safe_download_name(file.filename).rsplit(".", 1)[0]
        if len(paths) == 1:
            return send_file(paths[0], mimetype="image/jpeg", as_attachment=True, download_name=f"{base}.jpg")
        zip_path = os.path.join(work_dir, "pages.zip")
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for n, p in enumerate(paths, 1):
                zf.write(p, arcname=f"{base} - page {n}.jpg")
        return send_file(zip_path, mimetype="application/zip", as_attachment=True, download_name=f"{base} (pages).zip")
    except Exception as e:
        logger.error(f"PDF to pictures failed unexpectedly: {e}", exc_info=True)
        return jsonify({"error": "Turning the PDF into pictures failed. Please try again."}), 500
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


@app.route("/api/word/open", methods=["POST", "OPTIONS"])
@limiter.limit(HEAVY_LIMIT, key_func=_account_or_ip_key)
@auth_required
def word_open_endpoint():
    """3.1 — the Word editor opens a file: 'file' (.docx; a .doc, .odt or .rtf is made into a .docx first,
    and that .docx comes back too, as 'docx', to be saved from). No file: a new, empty document."""
    work_dir = tempfile.mkdtemp(prefix=f"mc_wordopen_{uuid.uuid4().hex[:8]}_")
    try:
        file = request.files.get("file")
        path, made = None, None
        if file is not None and file.filename:
            ext = file.filename.rsplit(".", 1)[-1].lower() if "." in file.filename else ""
            if ext not in ("docx", "doc", "odt", "rtf"):
                return jsonify({"error": "Choose a Word file (.docx or .doc)."}), 400
            path = os.path.join(work_dir, "source." + ext)
            file.save(path)
            if ext != "docx":
                try:
                    path = _soffice_convert(path, "docx", os.path.join(work_dir, "made"))
                except ConversionError:
                    return jsonify({"error": "This file couldn't be opened — please save it as .docx in Word, then try again."}), 422
                with open(path, "rb") as fh:
                    made = base64.b64encode(fh.read()).decode("ascii")
        try:
            data = word_open(path)
        except ConversionError as e:
            return jsonify({"error": str(e)}), 422
        if made:
            data["docx"] = made
        return jsonify(data)
    except Exception as e:
        logger.error(f"Word editor open failed unexpectedly: {e}", exc_info=True)
        return jsonify({"error": "Opening the Word file failed. Please try again."}), 500
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


@app.route("/api/word/save", methods=["POST", "OPTIONS"])
@limiter.limit(HEAVY_LIMIT, key_func=_account_or_ip_key)
@auth_required
def word_save_endpoint():
    """3.1 — the Word editor's Save: 'file' (the .docx that was opened; none for a new document), 'plan'
    (JSON, see word_save), 'assets' (pictures added, by their names in the plan) and, for a new
    document, 'name'. Returns the new .docx — "<name> (edited).docx" for a file that was opened."""
    try:
        plan = json.loads(request.form.get("plan") or "{}")
    except json.JSONDecodeError:
        return jsonify({"error": "The changes couldn't be read — please try again."}), 400
    file = request.files.get("file")
    has_file = file is not None and bool(file.filename)
    if has_file and not file.filename.lower().endswith(".docx"):
        return jsonify({"error": "The file must be a .docx Word file."}), 400
    try:
        _log_conversion_and_consume(request.current_user, "docx", "docx-edit")
    except ConversionError as e:
        return jsonify({"error": str(e)}), 402
    work_dir = tempfile.mkdtemp(prefix=f"mc_wordsave_{uuid.uuid4().hex[:8]}_")
    try:
        src = None
        if has_file:
            src = os.path.join(work_dir, "source.docx")
            file.save(src)
        assets = {}
        for n, a in enumerate(request.files.getlist("assets")[:60]):
            if not a.filename:
                continue
            safe, _e = safe_upload_filename(a.filename)
            p = os.path.join(work_dir, f"asset{n}_{safe}")
            a.save(p)
            assets[a.filename] = p
        try:
            result = word_save(src, plan, os.path.join(work_dir, "out"), assets)
        except ConversionError as e:
            return jsonify({"error": str(e)}), 422
        if has_file:
            name = safe_download_name(file.filename).rsplit(".", 1)[0] + " (edited).docx"
        else:
            name = safe_download_name(str(request.form.get("name") or "Document")[:80] + ".docx")
        return send_file(result, mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                         as_attachment=True, download_name=name)
    except Exception as e:
        logger.error(f"Word editor save failed unexpectedly: {e}", exc_info=True)
        return jsonify({"error": "Saving the Word file failed. Please try again."}), 500
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
        cap = REMY_ATTACH_PRO if (request.form.get("limit") == str(REMY_ATTACH_PRO) and _is_pro(request.current_user)) else REMY_ATTACH_FREE
        cut = len(text) > cap
        if cut:
            text = text[:cap]
        return jsonify({"text": text, "filename": safe_download_name(file.filename), "cut": cut, "limit": cap})
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
    '  "header": the HEADING of this slide — the topic it is about, as a lecturer would title it '
    '(e.g. "Areas of Biomechanics", "Causes of Academic Stress"). Consecutive slides on the same '
    'topic may share a heading.\n'
    '  "sections": the body of this slide, in the lecture pattern — an array of 1 to 3 objects, each '
    '{"subheading": a SUB-HEADING naming the specific aspect discussed (at most 7 words, e.g. '
    '"Developmental biomechanics"), "points": an array of 1 to 4 short points about it (each at '
    'most 16 words)} — empty when this slide instead uses "chart" below\n'
    '  "bullets": leave this an empty array — "sections" replaces it\n'
    '  "icon": one single emoji that visually represents this slide\'s content (e.g. '
    '\U0001F4C8 for growth, \U0001F4A1 for an idea, \u2705 for a takeaway/action) — pick '
    "one that actually fits the specific content, not the same one repeatedly\n"
    '  "chart": either null, or an object {"type": "bar"|"line"|"pie", "categories": '
    '[...], "series_name": "...", "values": [...]} when — and only when — the source '
    "material contains real numeric/comparative data suited to a chart\n"
    '  "is_cta": true only on the final slide, false everywhere else\n\n'
    "Follow every one of these — each is a specific, checkable property, not a vague style "
    "goal:\n\n"
    "1. HEADING, SUB-HEADING, THEN THE POINTS. Every content slide reads like a well-made "
    "lecture slide: the heading says the topic we are on, each sub-heading says what we are "
    "about to talk about, and its points say it. At most 3 sub-headings and 8 points on a "
    "slide — split a big topic across several slides that share the heading rather than "
    "crowding one. Points are short and specific, never full paragraphs.\n"
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



# Smart Summarize for a student's or a tutor's presentation: the AI writes only
# the content (and the tutor's objectives and discussion questions, and any
# references that appear IN the material); the server then lays the deck out in
# a fixed order, so the structure is guaranteed rather than hoped for.
SUMMARY_STRUCTURE = {
    "student": "",
    "tutor": (
        '\n  "objectives": 3 to 5 learning objectives for this session, each beginning with a verb '
        '("Explain…", "Identify…", "Compare…"), drawn only from the material\n'
        '  "discussion": 2 or 3 short discussion questions about the material, for the end of the session\n'),
}


def _summary_structure_prompt(audience):
    return (SUMMARIZE_SYSTEM_PROMPT + "\n\nTHIS TIME, RETURN ONE JSON OBJECT INSTEAD OF A BARE ARRAY — nothing before or after it, "
            "starting with { and ending with } — with these keys:\n"
            '  "title": the presentation\'s title, taken from the material (its title or main topic), at most 12 words\n'
            '  "slides": the slide array exactly as described above, where the FIRST slide is an introduction / overview '
            'of the topic (its header begins "Introduction" or "Overview"), then the content in the order the material '
            'presents it. Do NOT include a cover, outline, references, objectives, discussion or thank-you slide — those '
            'are added separately — and no slide has is_cta true.\n'
            '  "references": every reference or citation that appears in the material, copied exactly as written there '
            "(with its web address \u2014 DOI or link \u2014 when the material gives one). Never invent, complete or add a reference; "
            "if the material cites nothing, return [].\n"
            "When the material uses in-text citations such as (Mensah & Boateng, 2023), keep them at the end of the points "
            "that use them, copied exactly.\n"
            + SUMMARY_STRUCTURE[audience])


def _json_object(raw_text):
    t = str(raw_text or "")
    a, b = t.find("{"), t.rfind("}")
    if a < 0 or b <= a:
        raise ConversionError("The summary came back in an unexpected format — please try again.")
    try:
        return json.loads(t[a:b + 1])
    except ValueError:
        raise ConversionError("The summary came back in an unexpected format — please try again.")


def _structure_summary(audience, obj, slides, presenters):
    """Lays the deck out in the fixed order for a student or a tutor."""
    clean = lambda xs, n, cap=300: [str(x).strip()[:cap] for x in (xs or []) if str(x).strip()][:n]
    title = str(obj.get("title") or (slides[0]["header"] if slides else "") or "Presentation").strip()[:120]
    refs = clean(obj.get("references"), 30)
    for s in slides:
        s["is_cta"] = False
    plain = lambda header, items, icon: [{"header": header, "bullets": items[i:i + 6], "icon": icon, "chart": None, "is_cta": False}
                                         for i in range(0, len(items), 6)]
    outline = list(dict.fromkeys(re.sub(r"\*\*", "", s["header"]).strip() for s in slides))
    out = [{"kind": "cover", "header": title, "bullets": [], "icon": "", "chart": None, "is_cta": False}]
    if audience == "student":
        out += plain("Presented by", presenters or ["Name \u2014 Index number"], "\U0001F465")
        out += plain("Presentation Outline", outline, "\U0001F5C2\ufe0f")
        out += slides
        out += plain("References", refs or [f"Source material: {title}"], "\U0001F4DA")
    else:
        objectives = clean(obj.get("objectives"), 6, 160) or [f"Describe the key ideas in: {h}" for h in outline[:4]]
        questions = clean(obj.get("discussion"), 4, 200) or ["What stood out to you most in this session, and why?"]
        out += plain("References & Recommended Books", refs or [f"Main source: {title}"], "\U0001F4DA")
        out += plain("Learning Objectives", objectives, "\U0001F3AF")
        out += plain("Presentation Outline", outline, "\U0001F5C2\ufe0f")
        out += slides
        out += plain("Discussion", questions, "\U0001F4AC")
        out.append({"kind": "thanks", "header": "Thank you", "bullets": [], "icon": "", "chart": None, "is_cta": False})
    return title, out


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
        sections = []
        for sec in (item.get("sections") or [])[:3]:
            if not isinstance(sec, dict):
                continue
            sh = str(sec.get("subheading") or "").strip()[:120]
            pts = [str(p).strip()[:240] for p in (sec.get("points") or []) if str(p).strip()][:4]
            if sh or pts:
                sections.append({"subheading": sh, "points": pts})
        if sections and not bullets:
            # Older readers of a slide (and the outline) still see sensible bullets.
            bullets = [x for sec in sections for x in ([sec["subheading"]] if sec["subheading"] else []) + sec["points"]][:10]
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
            "sections": sections or None,
            "icon": icon,
            "chart": chart,
            "is_cta": bool(item.get("is_cta")),
        })
    if not slides:
        raise ConversionError("Couldn't generate a presentation from that material — please try again.")
    return slides


SAGE_SYSTEM_PROMPT = (
    "You are Remy, Academic-Home's built-in AI assistant — not a narrow, single-purpose tool "
    "like the app's other AI features, but a genuine, open-ended assistant anyone can "
    "bring almost anything to: students and teachers, people at work, and anyone with an "
    "everyday question. Academic-Home also has three purpose-built AI "
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
    "rather than only handing over a final answer, writing and improving text (emails, "
    "letters, reports, essays, posts), planning (a week, a study timetable, a lesson, an "
    "event or a project), brainstorming, comparing options, translating, answering "
    "questions about material someone has pasted or uploaded, and anything else a person "
    "would reasonably bring to a knowledgeable, patient assistant. Don't assume who "
    "someone is: adapt to whoever they show themselves to be — a student, a teacher or "
    "lecturer, someone at work, a parent, or anyone else — and to the level they're "
    "working at. You have access to web search for anything that depends on "
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


# 3.4: a source for every sentence that states something.
SAGE_SYSTEM_PROMPT += (
    "\n\nEvidence, sentence by sentence: when someone asks about facts \u2014 explaining or defining something, how or "
    "why something happens, history, science, health, law, figures, or anything current \u2014 every sentence of your "
    "answer that states something must be backed by a source: a numbered paper [n] when papers are listed below, or a "
    "page you found with web search (search whenever the papers don't cover it, and write from what the pages say). "
    "Don't write a sentence you can't back with a source; if the evidence doesn't cover part of the question, say so in "
    "one sentence. A short lead-in that states nothing (\"Here's how it works:\") or a question back to the person "
    "needs no source. When the request doesn't call for evidence \u2014 small talk, writing or improving the person's own "
    "email, letter, post or story, translating, planning, brainstorming, code, or working through a calculation they "
    "gave you \u2014 answer as usual, with no citations."
)


# ------------------------------------------------ Remy: evidence first (3.0)
_REMY_NO_PAPERS = re.compile(
    r"\b(?:e-?mails?|letters?|cover letter|cv|resume|grammar|spelling|proofread|rewrite|rephrase|paraphrase|translat\w*|"
    r"timetable|schedule|plan my|my (?:week|day|weekend)|to-?do|brainstorm|ideas? for|names? for|poem|story|joke|song|lyrics|"
    r"caption|tweet|post for|code|program(?:me)? in|python|javascript|html|css|sql|excel formula|recipe|cook|weather|news|today|"
    r"tonight|tomorrow|latest|right now|this (?:week|month|year)|score|who won|exchange rate|price of)\b", re.I)


def _remy_wants_papers(message):
    """A question about a topic, where real papers are the right evidence."""
    m = str(message or "")
    if len(m) > 2000 or _REMY_NO_PAPERS.search(m):
        return False
    q = _search_topic(m)
    return _worth_searching(q) and len(_paper_terms(q)) >= 1


def _paper_reference(p):
    """One paper as an APA reference line."""
    try:
        return enrich_sources([dict(p)], "apa")[0]["ref_text"]
    except Exception:
        return f"{p.get('title', '')}."


def _remy_papers_block(papers):
    lines = []
    for i, p in enumerate(papers, 1):
        au = p.get("authors") or []
        if au:
            sur = [a.split()[-1] for a in au if a.split()]
            who = sur[0] + (" et al." if len(sur) > 2 else (" & " + sur[1] if len(sur) == 2 else ""))
        else:
            who = p.get("venue") or "Unknown authors"
        lines.append(f"[{i}] {who} ({p.get('year') or 'n.d.'}). {p['title']}." + (f" {p['venue']}." if p.get("venue") else ""))
        if p.get("abstract"):
            lines.append(f"    Abstract: {p['abstract'][:700]}")
    return ("EVIDENCE FIRST: below are real, open-access research papers found for this question (each one can be downloaded "
            "free), numbered. Ground your answer in them first — explain in your own words, never copy — and cite them "
            "inline as [1], [2] right after the sentences they support, using only these numbers. Cite a paper only for what its "
            "title or abstract actually supports. For anything they don't cover, search the web and write from what the pages say — never state a fact without a source. Never invent a paper, an author, a statistic or a quotation. Don't list the papers at the end "
            "— the app shows them under your answer.\n\nPAPERS:\n" + "\n".join(lines))


_REMY_CITE = re.compile(r"(?<![\w\]\[])\[(\d{1,2}(?:\s*[,–\-]\s*\d{1,2})*)\]")


def _number_cited_papers(text, papers):
    """Remy's [n] citations renumbered 1, 2, 3… in the order they first appear; a number that
    matches no paper is taken out. Returns (text, the cited papers in that order)."""
    if not papers or not text:
        return text, []

    def nums_of(group):
        out = []
        for part in re.split(r"\s*,\s*", group):
            r = re.match(r"^(\d+)\s*[–\-]\s*(\d+)$", part)
            if r:
                a, b = int(r.group(1)), int(r.group(2))
                out += list(range(a, b + 1)) if a <= b and b - a < 10 else [a, b]
            elif part.strip().isdigit():
                out.append(int(part))
        return out

    order = []
    for m in _REMY_CITE.finditer(text):
        for k in nums_of(m.group(1)):
            if 1 <= k <= len(papers) and k not in order:
                order.append(k)
    new_no = {k: i + 1 for i, k in enumerate(order)}

    def rewrite(m):
        nums = sorted({new_no[k] for k in nums_of(m.group(1)) if k in new_no})
        return "[" + ", ".join(str(n) for n in nums) + "]" if nums else ""

    out = _REMY_CITE.sub(rewrite, text)
    out = re.sub(r"[ \t]+([.,;:])", r"\1", out)
    out = re.sub(r"[ \t]{2,}", " ", out)
    used = []
    for k in order:
        p = papers[k - 1]
        used.append({"n": new_no[k], "title": p["title"], "authors": (p.get("authors") or [])[:6], "year": p.get("year"),
                     "venue": p.get("venue"), "url": p["url"], "pdf": p.get("pdf"), "doi": p.get("doi"), "text": _paper_reference(p)})
    return out, used


_SMALL_WORDS = {"a", "an", "and", "as", "at", "but", "by", "for", "in", "of", "on", "or", "the", "to", "vs", "via", "with", "from", "into"}


def _topic_title(message):
    """A document title from what the question is about ("hand hygiene in hospitals" ->
    "Hand Hygiene in Hospitals") — never the question itself; None when there's no clear topic."""
    t = re.sub(r"[?!.]+$", "", _search_topic(message)).strip()
    words = t.split()
    if not (1 <= len(words) <= 12) or len(t) > 90:
        return None
    out = []
    for i, w in enumerate(words):
        if w.isupper() and len(w) > 1:
            out.append(w)
        elif i and w.lower() in _SMALL_WORDS:
            out.append(w.lower())
        else:
            out.append(w[:1].upper() + w[1:])
    return " ".join(out)


# ------------------------------------------------ 3.4: a source for every sentence
C34_WMARK = re.compile(r"\u27e6w:(\d+)\u27e7")      # \u27e6w:i\u27e7 \u2014 a passage the web search grounded in page i (private, never shown)
_C34_SKIP_LINE = re.compile(r"^\s*(?:#{1,6}\s|\||```|~~~|\$\$|[-*_]{3,}\s*$|!\[|<|\u27e6sec:)")
_C34_ITEM = re.compile(r"^\s*(?:[-*+\u2022]|\d{1,3}[.)])\s+")
_C34_LABEL = re.compile(r"^\s*(?:\*\*|__)[^*_\n]{1,140}(?:\*\*|__)\s*:?\s*$")
_C34_OFFER = re.compile(r"^\W*(?:let me know|if you(?:'d|\u2019d| would)? like|would you like|do you want|want me to|"
                        r"i can (?:also )?|i could|feel free|here(?:'s|\u2019s| is| are)\b|below (?:is|are)\b|happy to|hope this)", re.I)
_C34_ABBR = re.compile(r"(?:\b(?:e\.g|i\.e|et al|etc|vs|cf|approx|ca|fig|figs|no|nos|vol|pp?|eds?|dr|mr|mrs|ms|prof|st|jr|sr|"
                       r"inc|ltd|co|u\.s|u\.k)|(?<![\w.])[A-Za-z])\.$", re.I)
_C34_END = re.compile(r"[.!?]+[\"\u201d\u2019')\]*_]*")
_C34_AFTER = re.compile(r"([.!?])((?:[ \t]*(?:\[\d{1,2}(?:\s*[,\u2013\-]\s*\d{1,2})*\]|\u27e6w:\d+\u27e7))+)")


def _c34_sentences(text):
    """Every sentence of the answer's prose, as (start, end) spans in text. Headings, tables, code, pictures and
    bold labels aren't sentences; a list item's marker isn't part of its sentence."""
    out, pos, fence = [], 0, False
    for line in text.split("\n"):
        start = pos
        pos += len(line) + 1
        st = line.strip()
        if st.startswith("```") or st.startswith("~~~"):
            fence = not fence
            continue
        if fence or not st or _C34_SKIP_LINE.match(line) or _C34_LABEL.match(line):
            continue
        m = _C34_ITEM.match(line)
        off = m.end() if m else len(line) - len(line.lstrip())
        body, b0 = line[off:], 0
        for e in _C34_END.finditer(body):
            end = e.end()
            if end < len(body) and not body[end].isspace():
                continue                                   # 3.5, e.g.,…
            rest = body[end:].lstrip()
            if rest and rest[0].islower():
                continue                                   # "… et al. (2020) found" carries on
            if e.group(0)[0] == "." and _C34_ABBR.search(body[max(b0, e.start() - 12):e.start() + 1]):
                continue
            if body[b0:end].strip():
                out.append([start + off + b0, start + off + end])
            b0 = end
        if body[b0:].strip():
            out.append([start + off + b0, start + off + len(body)])
    for s in out:
        while s[0] < s[1] and text[s[0]].isspace():
            s[0] += 1
    return [tuple(s) for s in out if s[0] < s[1]]


def _c34_needs(s):
    """A sentence that states something, so it needs a source: not a lead-in ending in ':', not a question, not a
    fragment of a few words, not an offer to help."""
    plain = re.sub(r"[*_`#>]|\[\d[^\]]*\]|\u27e6w:\d+\u27e7|\([^()]*\b(?:1[5-9]|20)\d{2}[a-z]?\b[^()]*\)", " ", s).strip()
    if len(re.findall(r"[^\W\d_]+(?:['\u2019-][^\W\d_]+)*|\d+(?:[.,]\d+)*", plain)) < 4:
        return False
    t = plain.rstrip(" \"'\u201d\u2019)")
    if t.endswith("?") or t.endswith(":") or _C34_OFFER.match(plain):
        return False
    return True


def _c34_cut(text, a, b):
    """Takes the sentence at [a, b) out, with the space next to it; a line (or list item) left empty goes too."""
    if b < len(text) and text[b] in " \t":
        b += 1
    elif a > 0 and text[a - 1] in " \t":
        a -= 1
    text = text[:a] + text[b:]
    ls = text.rfind("\n", 0, a) + 1
    le = text.find("\n", a)
    le = len(text) if le == -1 else le
    line = text[ls:le]
    if not line.strip() or (_C34_ITEM.match(line) and not line[_C34_ITEM.match(line).end():].strip()):
        text = text[:ls] + text[le + 1:]
    return text


def _c34_tidy(text):
    text = re.sub(r"[ \t]+(?=\n|$)", "", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def _c34_put(sentence, mark):
    """The sentence with its source put just before its full stop."""
    m = re.search(r"[.!?]+[\"\u201d\u2019')\]*_]*$", sentence)
    cut = m.start() if m else len(sentence.rstrip())
    return sentence[:cut].rstrip() + " " + mark + sentence[cut:]


def _c34_check(kind, request_text, evidence_lines, sentences, context, deadline):
    """Asks the AI, sentence by sentence, which listed evidence supports each one, or that none does.
    Returns {"skip": bool, "cite": {i: [ids]}, "drop": {i}} \u2014 or None when the check couldn't run."""
    if not sentences or not evidence_lines or time.time() > deadline - 4:
        return None
    system = (
        "You check an answer so that every sentence that states something is backed by evidence. Below: the EVIDENCE "
        "(each item with an id), then the person's REQUEST, then numbered SENTENCES from the answer that have no source "
        "yet. For each sentence: if one or more evidence items directly support what it says, give their ids (at most "
        "3); if none does, mark it drop. Never pair a sentence with evidence that doesn't support it. If the request "
        "doesn't call for evidence at all (small talk, the person's own email, letter or story, editing their text, "
        "translating, planning, brainstorming, code, a calculation), answer {\"skip\": true}. Reply with JSON only, "
        "like {\"skip\": false, \"sentences\": [{\"n\": 1, \"cite\": [\"P2\"]}, {\"n\": 2, \"drop\": true}]}.\n\n"
        "EVIDENCE:\n" + "\n".join(evidence_lines))
    user = ("REQUEST:\n" + str(request_text or "")[:1500] + "\n\nSENTENCES:\n"
            + "\n".join(f"{i + 1}. {s}" for i, s in enumerate(sentences))
            + "\n\nTHE WHOLE ANSWER, for context:\n" + str(context or "")[:6000])
    try:
        raw = str(call_claude(system_prompt=system, user_message=user, max_tokens=min(1500, 60 * len(sentences) + 150),
                              timeout=max(5, min(25, deadline - time.time())), feature=kind + "-check") or "")
        a, b = raw.find("{"), raw.rfind("}")
        obj = json.loads(raw[a:b + 1]) if 0 <= a < b else None
        if not isinstance(obj, dict):
            raise ValueError("its reply wasn't the JSON asked for")
    except Exception as e:
        logger.warning(f"Citations ({kind}): the sentence check didn't run ({str(e)[:100]})")
        return None
    out = {"skip": obj.get("skip") is True, "cite": {}, "drop": set()}
    for it in obj.get("sentences") or []:
        if not isinstance(it, dict):
            continue
        try:
            i = int(it.get("n")) - 1
        except (TypeError, ValueError):
            continue
        if not 0 <= i < len(sentences):
            continue
        ids = [str(x).strip().upper() for x in (it.get("cite") or []) if str(x).strip()][:3]
        if ids and not it.get("drop"):
            out["cite"][i] = ids
        else:
            out["drop"].add(i)
    return out


def _c34_enforce(kind, text, request_text, is_cited, evidence_lines, make_mark, deadline, allow_skip):
    """Every sentence that states something gets a source or goes. Returns (text, unsourced): unsourced counts the
    sentences left without one \u2014 only when taking them out would remove most of the answer."""
    need = [s for s in _c34_sentences(text) if _c34_needs(text[s[0]:s[1]])]
    todo = [s for s in need if not is_cited(text[s[0]:s[1]])]
    if not todo:
        return text, 0
    todo = todo[:30]
    any_cited = len(todo) < len(need)
    dec = _c34_check(kind, request_text, evidence_lines, [text[a:b] for a, b in todo], text, deadline)
    if dec and dec["skip"] and allow_skip and not any_cited:
        logger.info(f"Citations ({kind}): the request doesn't call for evidence \u2014 left as it is")
        return text, 0
    fixed = dropped = 0
    if dec:
        for i in sorted(range(len(todo)), key=lambda i: -todo[i][0]):
            a, b = todo[i]
            mark = make_mark(dec["cite"][i]) if i in dec["cite"] else None
            if mark:
                text = text[:a] + _c34_put(text[a:b], mark) + text[b:]
                fixed += 1
            elif i in dec["drop"] or i in dec["cite"]:
                text = _c34_cut(text, a, b)
                dropped += 1
        text = _c34_tidy(text)
    need = [s for s in _c34_sentences(text) if _c34_needs(text[s[0]:s[1]])]
    left = [s for s in need if not is_cited(text[s[0]:s[1]])]
    unsourced = 0
    if left:
        if not dec and not any_cited and allow_skip:
            left = []                    # the check couldn't run and nothing is cited: most likely not a request for evidence
        elif len(left) * 2 <= len(need):
            for a, b in sorted(left, key=lambda s: -s[0]):
                text = _c34_cut(text, a, b)
            text = _c34_tidy(text)
            dropped += len(left)
            left = []
        else:
            unsourced = len(left)
    logger.info(f"Citations ({kind}): {len(need) + dropped} sentences, {len(todo)} had no source \u2192 {fixed} given one, "
                f"{dropped} taken out, {unsourced} kept with a note")
    return text, unsourced


# ---- Remy
def _c34_marked_text(content):
    """The answer's text with a private mark after each passage the web search grounded (its page's index), and the
    pages: title, address, and the lines the answer drew on."""
    pages, idx, out = [], {}, []
    for b in content or []:
        if not isinstance(b, dict) or b.get("type") != "text":
            continue
        t, marks = str(b.get("text") or ""), []
        for c in b.get("citations") or []:
            url = str((c or {}).get("url") or "")[:600]
            if not url.startswith("http"):
                continue
            if url not in idx:
                idx[url] = len(pages)
                pages.append({"url": url, "title": str(c.get("title") or url).strip()[:200], "quotes": []})
            q = re.sub(r"\s+", " ", str(c.get("cited_text") or "")).strip()
            if q and len(pages[idx[url]]["quotes"]) < 3 and q not in pages[idx[url]]["quotes"]:
                pages[idx[url]]["quotes"].append(q[:320])
            if idx[url] not in marks:
                marks.append(idx[url])
        if marks:
            core = t.rstrip()
            t = core + "".join(f"\u27e6w:{i}\u27e7" for i in marks) + t[len(core):]
        out.append(t)
    return "".join(out), pages


def _c34_nums(group):
    out = []
    for part in re.split(r"\s*,\s*", group):
        r = re.match(r"^(\d+)\s*[\u2013\-]\s*(\d+)$", part)
        if r:
            a, b = int(r.group(1)), int(r.group(2))
            out += list(range(a, b + 1)) if a <= b and b - a < 10 else [a, b]
        elif part.strip().isdigit():
            out.append(int(part))
    return out


def _c34_remy_cited(papers):
    def cited(s):
        if C34_WMARK.search(s):
            return True
        return bool(papers) and any(1 <= k <= len(papers) for m in _REMY_CITE.finditer(s) for k in _c34_nums(m.group(1)))
    return cited


def _c34_number(text, papers, pages):
    """Papers and web pages numbered 1, 2, 3\u2026 in one series, in the order the answer first cites them. A number
    that matches no paper is taken out (when there are papers). Returns (text, papers cited, pages cited)."""
    pat = re.compile(r"\u27e6w:(\d+)\u27e7|(?<![\w\]\[])\[(\d{1,2}(?:\s*[,\u2013\-]\s*\d{1,2})*)\]")

    def keys(m):
        if m.group(1) is not None:
            i = int(m.group(1))
            return [("w", i)] if 0 <= i < len(pages) else []
        return [("p", k) for k in _c34_nums(m.group(2)) if 1 <= k <= len(papers)]

    order, new = [], {}
    for m in pat.finditer(text):
        if m.group(2) is not None and not papers:
            continue
        for k in keys(m):
            if k not in new:
                new[k] = len(order) + 1
                order.append(k)

    def rewrite(m):
        if m.group(2) is not None and not papers:
            return m.group(0)
        nums = sorted({new[k] for k in keys(m)})
        if not nums:
            return ""
        pre = m.string[m.start() - 1] if m.start() else " "
        return ("" if pre.isspace() or pre in "([\u27e7" else " ") + "\u27e6c:" + ",".join(map(str, nums)) + "\u27e7"

    out = pat.sub(rewrite, text)
    # side-by-side numbers become one group: [1][3] -> [1, 3]
    join = re.compile(r"\u27e6c:([\d,]+)\u27e7\s*\u27e6c:([\d,]+)\u27e7")
    while join.search(out):
        out = join.sub(lambda m: "\u27e6c:" + ",".join(str(n) for n in sorted({int(x) for x in (m.group(1) + "," + m.group(2)).split(",")})) + "\u27e7", out)
    out = re.sub(r"\u27e6c:([\d,]+)\u27e7", lambda m: "[" + ", ".join(m.group(1).split(",")) + "]", out)
    out = re.sub(r"[ \t]+([.,;:])", r"\1", out)
    out = re.sub(r"[ \t]{2,}", " ", out)
    used, web = [], []
    for k in order:
        n = new[k]
        if k[0] == "p":
            p = papers[k[1] - 1]
            used.append({"n": n, "title": p["title"], "authors": (p.get("authors") or [])[:6], "year": p.get("year"),
                         "venue": p.get("venue"), "url": p["url"], "pdf": p.get("pdf"), "doi": p.get("doi"), "text": _paper_reference(p)})
        else:
            pg = pages[k[1]]
            web.append({"n": n, "title": pg["title"], "url": pg["url"], "cited": True})
    return out, used, web


def _c34_remy(text, papers, pages, found, message, topic_q, started, attachment):
    """Remy's answer with a source on every sentence that states something (see the top of this patch).
    Returns (text, papers cited, web pages: the cited ones numbered, then others the search found, unsourced)."""
    text = _C34_AFTER.sub(lambda m: (" " if m.group(2).strip().startswith("[") else "") + m.group(2).strip() + m.group(1), text)
    is_cited = _c34_remy_cited(papers)
    has_cite = bool(C34_WMARK.search(text)) or (bool(papers) and any(is_cited(m.group(0)) for m in _REMY_CITE.finditer(text)))
    unsourced = 0
    if not attachment and (papers or pages) and (has_cite or topic_q):
        lines = []
        for i, p in enumerate(papers, 1):
            lines.append(f"[P{i}] {p.get('title', '')} ({p.get('year') or 'n.d.'})" + (f" \u2014 {str(p.get('abstract'))[:600]}" if p.get("abstract") else ""))
        for i, pg in enumerate(pages, 1):
            if pg["quotes"]:
                lines.append(f"[W{i}] {pg['title']} \u2014 " + " / ".join(f"\u201c{q}\u201d" for q in pg["quotes"]))

        def mark(ids):
            ms = []
            for x in ids:
                if x[:1] == "P" and x[1:].isdigit() and 1 <= int(x[1:]) <= len(papers):
                    ms.append(f"[{int(x[1:])}]")
                elif x[:1] == "W" and x[1:].isdigit() and 1 <= int(x[1:]) <= len(pages) and pages[int(x[1:]) - 1]["quotes"]:
                    ms.append(f"\u27e6w:{int(x[1:]) - 1}\u27e7")
            return "".join(ms) or None

        text, unsourced = _c34_enforce("remy", text, message, is_cited, lines, mark, started + 50, allow_skip=not has_cite)
    elif topic_q and not has_cite:
        # a question about facts, and nothing found to back the answer: it stays, and says so
        need = [s for s in _c34_sentences(text) if _c34_needs(text[s[0]:s[1]])]
        unsourced = len(need) if len(need) >= 2 else 0
        if unsourced:
            logger.info(f"Citations (remy): no evidence found for a question about facts — {unsourced} sentences kept with a note")
    text, used, web = _c34_number(text, papers, pages)
    seen = {w["url"] for w in web}
    rest = [dict(s, cited=False) for s in (found or []) if isinstance(s, dict) and s.get("url") not in seen]
    return text, used, (web + rest)[:max(8, len(web))], unsourced


# ---- Evidence-Based Writing
def _c34_write_parts(sources):
    lines = []
    for i, s in enumerate(sources or [], 1):
        line = f"[S{i}] {s.get('citation_key', '')} {s.get('title', '')}"
        if s.get("abstract"):
            line += f" \u2014 {str(s['abstract'])[:600]}"
        lines.append(line)

    def mark(ids):
        inner = []
        for x in ids:
            if x[:1] == "S" and x[1:].isdigit() and 1 <= int(x[1:]) <= len(sources):
                key = str(sources[int(x[1:]) - 1].get("citation_key") or "")
                if key.startswith("(") and key.endswith(")") and key[1:-1] not in inner:
                    inner.append(key[1:-1])
        return "(" + "; ".join(inner) + ")" if inner else None
    return lines, mark


def _c34_write(text, sources, style, request_text, note, started):
    """Evidence-Based Writing: every sentence cited (see the top of this patch). Returns (text, cited indices)."""
    if not sources or not text:
        return text, keep_known_citations(text, sources, style)[1] if text else set()
    is_cited = lambda s: bool(keep_known_citations(s, sources, style)[1])
    lines, mark = _c34_write_parts(sources)
    has_cite = bool(keep_known_citations(text, sources, style)[1])
    text, unsourced = _c34_enforce("writing", text, request_text, is_cited, lines, mark, started + 150, allow_skip=not has_cite)
    if unsourced:
        note["text"] = (note["text"] + " " if note["text"] else "") + \
            "Some sentences in this piece have no source \u2014 check them before you use it."
    return keep_known_citations(text, sources, style)


def _c34_write_sections(outline, sources, style, request_text, note, started):
    """The same for an outline's subsections, checked together. Returns the cited indices."""
    joined = "\n\n".join(f"\u27e6sec:{i}\u27e7\n\n" + str(s.get("text") or "") for i, s in enumerate(outline))
    text, cited = _c34_write(joined, sources, style, request_text, note, started)
    parts = re.split(r"\u27e6sec:(\d+)\u27e7", text)
    for j in range(1, len(parts) - 1, 2):
        i = int(parts[j])
        if 0 <= i < len(outline):
            outline[i]["text"] = parts[j + 1].strip() or "(No sentence here could be backed by the sources.)"
    return cited


@app.route("/api/sage", methods=["POST", "OPTIONS"])
@auth_required
def sage_endpoint():
    started = time.time()
    data = request.get_json(silent=True) or {}
    message = (data.get("message") or "").strip()
    if not message:
        return jsonify({"error": "No message provided"}), 400
    if len(message) > 20000:
        return jsonify({"error": "That's too long for one message (20,000 character limit)"}), 400

    attachment = (data.get("attachment") or "").strip()
    att_max = REMY_ATTACH_PRO if _is_pro(request.current_user) else REMY_ATTACH_FREE
    if len(attachment) > att_max:
        return jsonify({"error": f"That attached document is too long ({att_max:,} character limit)"}), 400

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
        charged_from = _log_sage_message_and_consume(request.current_user)
    except ConversionError as e:
        return jsonify({"error": str(e)}), 429

    # Remy's memory (when the person allows it): what Remy noted before, and
    # permission to note one new durable thing — taken out of the reply below.
    memory_on = _memory_on(request.current_user)
    system_prompt = SAGE_SYSTEM_PROMPT
    if memory_on:
        notes = get_user_memory(request.current_user["id"])
        if notes:
            system_prompt += (
                "\n\nWhat you remember about this person from earlier — use it naturally where it "
                "genuinely helps; don't recite it back: " + "; ".join(notes))
        system_prompt += "\n\n" + MEMORY_INSTRUCTION
    system_prompt += "\n\n" + THINKING_INSTRUCTION
    # Evidence first (3.0): a question about a topic gets real open papers to ground the answer.
    papers = []
    topic_q = not attachment and _remy_wants_papers(message)
    if topic_q:
        try:
            papers, _ = paper_topic_search(message, max_results=REMY_PAPERS_PRO if _is_pro(request.current_user) else REMY_PAPERS_FREE)
        except Exception as e:
            logger.warning(f"Remy's paper search failed: {e}")
            papers = []
    if papers:
        system_prompt += "\n\n" + _remy_papers_block(papers)
    try:
        meta = call_claude(
            system_prompt=system_prompt,
            messages=turns,
            max_tokens=4000,
            use_search=True,
            model=ANTHROPIC_MODEL_STRONG,
            return_meta=True,
            feature="remy",
        )
        marked, pages = _c34_marked_text(meta.get("content"))
        result = marked if marked.strip() else meta["text"]
    except Exception as e:
        # The message was counted above, before the AI call — but no reply
        # arrived, so give it back. Otherwise a timeout or AI outage quietly
        # spent one of a free user's 10 monthly messages, and tapping Retry
        # spent another.
        _refund_sage_message(request.current_user, charged_from)
        if isinstance(e, ConversionError):
            return jsonify({"error": str(e) + AI_MSG_NOT_COUNTED}), 502
        raise
    thinking, result = split_thinking_anywhere(result or "")
    thinking = C34_WMARK.sub("", thinking) if thinking else thinking
    note, reply = split_memory_note(result or "")
    note = C34_WMARK.sub("", note).strip() if note else note
    if note and memory_on:
        add_user_memory(request.current_user["id"], note)
    reply, used, web, unsourced = _c34_remy((reply or "").strip() or result, papers, pages, meta.get("sources") or [],
                                           message, topic_q, started, bool(attachment))
    reply = C34_WMARK.sub("", reply)
    return jsonify({"reply": reply, "thinking": thinking, "papers": used, "title": _topic_title(message) if papers else None,
                    "sources": web, "searched": (meta.get("queries") or [])[:5], "unsourced": unsourced})


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

    # For a student's or a tutor's presentation, the deck follows a fixed order
    # (see _structure_summary); without an audience, the original behaviour.
    audience = data.get("audience") if data.get("audience") in ("student", "tutor") else None
    reqs = str(data.get("requirements") or "").strip()[:1200]
    reqs_block = ("\n\nTHE PRESENTER'S REQUIREMENTS \u2014 follow them exactly (number of slides, sections to include, focus, "
                  "level of language), while keeping the JSON format described above:\n<<<\n" + reqs + "\n>>>") if reqs else ""
    presenters = [str(x).strip()[:80] for x in (data.get("presenters") or []) if str(x).strip()][:10] if audience == "student" else []
    try:
        result = call_claude(
            system_prompt=(_summary_structure_prompt(audience) if audience else SUMMARIZE_SYSTEM_PROMPT) + reqs_block,
            user_message=user_message,
            max_tokens=max_tokens + (900 if audience else 0),
        )
        if not audience:
            return jsonify({"slides": _parse_summary_json(result)})
        obj = _json_object(result)
        slides = _parse_summary_json(json.dumps(obj.get("slides") or []))
        title, deck = _structure_summary(audience, obj, slides, presenters)
        return jsonify({"slides": deck, "title": title, "audience": audience})
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
    style = data.get("style") if data.get("style") in ("visual", "plain") else "plain"   # 3.0: Plain White

    work_dir = tempfile.mkdtemp(prefix=f"mc_sum_{uuid.uuid4().hex[:8]}_")
    try:
        result_path = summary_slides_to_pptx(slides, work_dir, style=style)
        try:
            _colour_not_bold(result_path)
        except Exception as e:
            logger.warning(f"Colour-not-bold pass skipped: {e}")
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
    '  "explanation": written for the student who has just answered: one or two short sentences on '
    "why the correct answer is right and, for multiple_choice, why the most tempting wrong option is "
    "wrong — so a mistake becomes something learned\n"
    '  "difficulty": "easy", "medium" or "hard" — how demanding this particular question is\n\n'
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


QUIZ_STANDALONE_RULES = (
    "\n\nSTAND-ALONE QUESTIONS — the most important rule. The quiz is downloaded, printed and "
    "used on its own by people who never see the source text, so no question, option, answer or "
    "explanation may refer to the source: never \"the material\", \"the document\", \"the text\", "
    "\"the passage\", \"the notes\", \"the lecture\", \"the reading\", \"the article\", \"the author\", "
    "\"the source\" or \"the information given\", and never \"according to\" any of those. Name the "
    "subject itself instead (\"In biomechanics, …\", \"Under the WHO guidelines, …\"). Every question "
    "must make complete sense to someone holding only the printed quiz.\n\n"
    "RETURN ONE JSON OBJECT (not a bare array): {\"title\": a short, specific title naming the subject "
    "(at most 9 words, e.g. \"Introduction to Biomechanics\"), \"questions\": [the question objects]}."
)


def _quiz_case_rules(n_case, total, tutor=True):
    share = (f"Exactly {n_case} of the {total} questions must be CASE STUDIES, spread through the quiz. " if n_case < total
             else f"EVERY one of the {total} questions must be a CASE STUDY, each with its own scenario. ")
    return (
        "\n\nCASE STUDIES — " + ("this quiz is for a tutor to set. " if tutor else "") + share +
        "A case study has \"case_study\": true and a "
        "\"case\" field: a realistic scenario of 4 to 7 sentences with concrete details (a named person, "
        "age, setting, measurements or figures — whatever suits the subject; Ghanaian names and settings "
        "are welcome). Its \"question\" must only be answerable by applying the subject's ideas to that "
        "scenario — assess, decide, calculate, explain or predict — never by recalling a definition. "
        "Case studies may be multiple_choice or short_answer; their explanations refer to the case's "
        "details." + (" Every other question has \"case_study\": false and no \"case\"." if n_case < total else "")
    )


# Leftover references to the source, cleaned out of anything the AI still words that way.
_SRC_WORDS = (r"(?:study\s+)?(?:material|materials|document|documents|text|passage|notes|note|lecture(?:\s+notes)?|reading|"
              r"article|source(?:\s+text)?|content|information\s+(?:given|provided)|(?:given|provided)\s+(?:text|material|notes|information))")
_SRC_LEAD = re.compile(r"^\s*(?:according\s+to|based\s+on|as\s+(?:stated|described|mentioned|explained|discussed|noted|shown|outlined|defined)\s+in|"
                       r"as\s+per|from|in|per|using)\s+(?:the|this|these|your)\s+" + _SRC_WORDS + r"\s*,\s*", re.I)
_SRC_MID = re.compile(r",?\s+(?:according\s+to|as\s+(?:stated|described|mentioned|explained|discussed|noted|outlined|defined)\s+in|as\s+per|based\s+on)\s+"
                      r"(?:the|this|these|your)\s+" + _SRC_WORDS + r"(?=[\s,.?!;:)]|$)", re.I)
_SRC_PAIR = re.compile(r",\s+(?:according\s+to|as\s+(?:stated|described|mentioned|explained|discussed|noted|outlined|defined)\s+in|as\s+per|based\s+on)\s+"
                       r"(?:the|this|these|your)\s+" + _SRC_WORDS + r"\s*,\s*", re.I)
_SRC_TAIL = re.compile(r"\s+(?:in|from|within|by)\s+(?:the|this|these|your)\s+" + _SRC_WORDS + r"(?=\s*[?.!,;:)]|\s*$)", re.I)


def _standalone(text):
    t = str(text or "").strip()
    if not t:
        return t
    t2 = _SRC_LEAD.sub("", t)
    t2 = _SRC_PAIR.sub(" ", t2)       # "Which term, as described in the notes, refers…" → "Which term refers…"
    t2 = _SRC_MID.sub("", t2)
    t2 = _SRC_TAIL.sub("", t2)
    t2 = re.sub(r"\s{2,}", " ", t2).strip()
    if t2 != t and t2:
        t2 = t2[0].upper() + t2[1:]
    return t2 or t


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
    if not text.startswith(("[", "{")):
        a = min([i for i in (text.find("{"), text.find("[")) if i >= 0] or [0])
        text = text[a:]
    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        raise ConversionError("Couldn't generate a quiz from that material — please try again.")
    if isinstance(data, dict):        # {"title": …, "questions": […]}
        data = data.get("questions")
    if not isinstance(data, list) or not data:
        raise ConversionError("Couldn't generate a quiz from that material — please try again.")

    questions = []
    for item in data:
        if not isinstance(item, dict):
            continue
        q_type = item.get("type") if item.get("type") in ("multiple_choice", "short_answer") else "short_answer"
        q_text = _standalone(item.get("question"))
        if not q_text:
            continue
        options = item.get("options") if isinstance(item.get("options"), list) else None
        if q_type == "multiple_choice" and (not options or len(options) < 2):
            q_type = "short_answer"  # a malformed MC question still becomes a usable question, not a dropped one
            options = None
        if options:
            options = [_standalone(o) for o in options]
        correct_answer = _standalone(item.get("correct_answer"))
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
        if final_options and sorted(o.strip().lower() for o in final_options) == ["false", "true"]:
            final_options = ["True", "False"]
            correct_answer = "True" if correct_answer.strip().lower() == "true" else "False"
        elif final_options:
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
            "explanation": _standalone(item.get("explanation")),
            "difficulty": item.get("difficulty") if item.get("difficulty") in ("easy", "medium", "hard") else None,
        })
        case = _standalone(item.get("case")) if item.get("case_study") else ""
        questions[-1]["case_study"] = bool(case)
        questions[-1]["case"] = case[:2400] if case else None
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

    # How hard, and the lecturer's own requirements (topics, focus, format) — both optional.
    difficulty = data.get("difficulty") if data.get("difficulty") in ("easy", "medium", "hard", "mixed") else "mixed"
    level = {
        "easy": "DIFFICULTY: make every question introductory \u2014 clear understanding of the key ideas, one simple step of reasoning. Mark each \"easy\".",
        "medium": "DIFFICULTY: make every question course-level \u2014 understanding and applying ideas, one or two steps of reasoning. Mark each \"medium\".",
        "hard": ("DIFFICULTY: make every question exam-level and challenging \u2014 applying ideas to new cases, analysing and comparing, "
                 "multi-step reasoning. Mark each \"hard\"."),
        "mixed": "DIFFICULTY: a genuine range from easy to hard across the set, each question marked with its own level.",
    }[difficulty]
    audience = "tutor" if data.get("audience") == "tutor" else "student"
    reqs = str(data.get("requirements") or "").strip()[:1200]
    reqs_block = ("\n\nTHE LECTURER'S REQUIREMENTS \u2014 follow them exactly (topics to cover, question focus, examples, "
                  "format), while keeping the JSON format above and basing every question on the material:\n<<<\n" + reqs + "\n>>>") if reqs else ""
    style = str(data.get("question_style") or "mixed")
    if style == "multiple_choice":
        question_style = "only multiple_choice"
    elif style == "short_answer":
        question_style = "only short_answer"
    elif style == "true_false":
        question_style = ("only TRUE/FALSE questions: each is \"type\": \"multiple_choice\" with \"options\" exactly "
                          "[\"True\", \"False\"] and a clear statement as its \"question\"; roughly half true and half false, "
                          "and a false statement must be wrong in one specific, checkable way")
    else:
        question_style = "a mix of multiple_choice and short_answer"

    user_message = (
        f"Generate exactly {num_questions} questions ({question_style}) on the subject below — remember: "
        f"no question may mention the source.\n{level}\n\nSOURCE TEXT (for you only — the quiz-taker never sees it):\n\n{text}"
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
    n_case = num_questions if style == "case_study" else (max(1, round(num_questions / 3)) if audience == "tutor" else 0)
    # A little more room: explanations now teach, and a case study's scenario runs long.
    max_tokens = min(64000, max(4000, num_questions * 360 + n_case * 260 + 600))

    try:
        result = call_claude(
            system_prompt=QUIZ_SYSTEM_PROMPT + QUIZ_STANDALONE_RULES + (_quiz_case_rules(n_case, num_questions, audience == "tutor") if n_case else "") + reqs_block,
            user_message=user_message,
            max_tokens=max_tokens,
        )
        questions = _parse_quiz_json(result)
        if title in ("", "Practice Quiz"):
            try:
                made = json.loads(result[result.find("{"):result.rfind("}") + 1]) if "{" in result else {}
                named = str(made.get("title") or "").strip() if isinstance(made, dict) else ""
            except ValueError:
                named = ""
            named = _standalone(re.sub(r"\s+", " ", named))[:90]
            if named:
                title = named if "quiz" in named.lower() else f"{named} — Practice Quiz"
        return jsonify({"questions": questions, "title": title, "difficulty": difficulty, "audience": audience, "question_style": style})
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
    "You're Academic-Home's academic writing partner for a student or researcher — help with "
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
    "and why — as 2 to 4 short steps, one per line, each a single sentence in your own "
    "natural voice (the app shows them one at a time). Wrap it in the exact "
    f"markers {THINK_START} and {THINK_END} with nothing else on those lines, then "
    "immediately after the closing marker write your actual answer in full — the thinking "
    "is shown to the student separately and is never a substitute for the complete answer."
)


def split_thinking_anywhere(text):
    """Like split_thinking, but the block may come after a few words (a web-searching
    answer sometimes says something before it thinks)."""
    t = text or ""
    a = t.find(THINK_START)
    if a == -1:
        return None, t
    b = t.find(THINK_END, a)
    if b == -1:
        return None, t
    thinking = t[a + len(THINK_START):b].strip()
    rest = (t[:a] + t[b + len(THINK_END):]).strip()
    return (thinking or None), rest


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
    "person's ongoing work — their field of study or work, what they teach, a specific project "
    "or thesis they're on, a recurring topic, a citation-style or formatting preference — add ONE short sentence "
    "noting it, wrapped in the exact markers "
    f"{MEMORY_START} and {MEMORY_END}, placed after your full answer and any references. "
    "Only include this when you've learned something genuinely new and worth remembering for "
    "next time — omit it entirely otherwise, and never repeat something you already noted "
    "before. Never note sensitive personal details (health, money, religion, relationships, "
    "passwords or anything they'd keep private). They can see, edit and delete these notes "
    "in Profile -> Remy's memory."
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
    started = time.time()
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

    # Citation style and how recent the sources must be (both optional).
    style = str(data.get("style") or "apa").lower()
    style = style if style in CITE_STYLES else "apa"
    CS = CITE_STYLES[style]
    try:
        recent_years = int(data.get("recent_years") or 0)
    except (TypeError, ValueError):
        recent_years = 0
    recent_years = recent_years if recent_years in (5, 10) else 0
    note = {"text": ""}

    since = (datetime.now(timezone.utc).year - recent_years + 1) if recent_years else None

    def _find(text, n=8, topic=True):
        """Real, downloadable papers for this request, cited in the chosen style — recent ones
        when the student asked for them, older ones added (with a note) when too few came up."""
        look = paper_topic_search if topic else search_papers
        papers, diag = look(text, max_results=n, since=since)
        if since and len(papers) < 3 and diag != "not a topic":
            older, _ = look(text, max_results=n)
            have = {p["url"] for p in papers}
            fresh = len(papers)
            papers = papers + [p for p in older if p["url"] not in have][: n - len(papers)]
            if len(papers) > fresh:
                note["text"] = (f"Only {fresh} paper{'' if fresh == 1 else 's'} from the last {recent_years} years came up, "
                                "so older ones were added to support the writing properly.")
            if papers:
                diag = None
        return enrich_sources(papers, style), diag

    def _prep(srcs):
        if recent_years and srcs:
            cutoff = datetime.now(timezone.utc).year - recent_years
            fresh = [s for s in srcs if (s.get("year") or 0) >= cutoff]
            if len(fresh) >= 3:
                srcs = fresh
            else:
                note["text"] = (f"Only {len(fresh)} source{'' if len(fresh) == 1 else 's'} from the last {recent_years} years came up, "
                                "so older ones were kept to support the writing properly.")
        return enrich_sources(srcs, style)

    # The student's own brief (word count, sections, focus…) and a target
    # length, added to EVERY writing call below — including continuations —
    # by this local wrapper, so no path can miss them. They never override
    # the real-sources and exact-citation rules.
    reqs = str(data.get("requirements") or "").strip()[:1500]
    try:
        length_words = int(data.get("length_words") or 0)
    except (TypeError, ValueError):
        length_words = 0
    length_words = length_words if length_words in (500, 800, 1000, 1500, 2000, 3000) else 0
    extra = ""
    if reqs:
        extra += ("\n\nTHE STUDENT'S REQUIREMENTS (from their assignment brief) \u2014 follow them exactly: structure, sections, "
                  "focus, tone, audience and any word count. They never override the rules above about using only the "
                  "listed real sources with their exact citations.\n<<<\n" + reqs + "\n>>>")
    if length_words:
        extra += (f"\n\nLENGTH: write about {length_words} words in total (within 10%), not counting the reference list. "
                  "If the piece needs more room than one answer allows, stop at a natural break \u2014 it will be continued.")
    _call_claude = globals()["call_claude"]

    def call_claude(*args, **kw):
        if extra:
            if "system_prompt" in kw:
                kw["system_prompt"] = kw["system_prompt"] + extra
            elif args:
                args = (args[0] + extra,) + tuple(args[1:])
        return _call_claude(*args, **kw)

    # Continuation of a truncated previous answer — a distinct, simpler path:
    # no clarify/thinking/outline logic, just pick the citation-bearing
    # writing back up exactly where it stopped.
    if data.get("continue_previous"):
        previous_text = (data.get("previous_text") or "").strip()
        if not previous_text:
            return jsonify({"error": "Nothing to continue"}), 400
        try:
            cont_sources, cont_diag = _find(topic)
            if cont_sources:
                cont_sources_block = build_sources_block(cont_sources)
                cont_system_prompt = (
                    f"{ROLE_DESCRIPTION} You are continuing your own previous response, which "
                    "was cut off partway through. Below is a list of real sources for this "
                    f"topic, each labeled with its exact {CS['name']} in-text citation, e.g. "
                    f"{CS['example']}, listed most-recent-first. Continue writing directly from "
                    "where the previous text left off — do not repeat, restate, or summarize "
                    "anything already written. Keep citing EVERY sentence using the EXACT citation key next to its supporting source, copied "
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
            cont_text, cont_cited = keep_known_citations(cont_text, cont_sources, style)
            cont_text, cont_cited = _c34_write(cont_text, cont_sources, style, topic, note, started)
            return jsonify({
                "type": "result",
                "text": cont_text,
                "references": build_references_list(cont_sources, cont_cited), "references_title": CS["heading"], "source_note": note["text"], "style": style,
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
            sources, sources_diag = _find(combined_headings, topic=False)

            if sources:
                sources_block = build_sources_block(sources)
                base_prompt = (
                    "You are writing one subsection of a longer academic piece on "
                    f"\"{combined_headings}\". Write ONLY the subsection given below as the "
                    "user message — 100-160 words, formal academic prose. Below is a list of "
                    "real sources relevant to the overall piece, each labeled with its exact "
                    f"{CS['name']} in-text citation, e.g. {CS['example']}, listed most-recent-first. Write "
                    "the way you actually would — ground every factual or evidentiary claim in "
                    "one of these sources, using the EXACT key next to it, copied exactly — "
                    "never alter an author name or year, and never invent a citation not in "
                    "this list. Every sentence you write must carry at least one of these citations — openings, links between ideas and conclusions too: a sentence that draws sources together cites each of them. Write no sentence you can't cite; leave it out instead. When more than one source could support the same claim, prefer the "
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
            cited_all = set()
            for sec in outline:
                sec["text"], c = keep_known_citations(sec["text"], sources, style)
                cited_all |= c
            if sources:
                cited_all = _c34_write_sections(outline, sources, style, topic, note, started)

            return jsonify({
                "type": "result",
                "sections": [
                    {"number": s["number"], "heading": s["heading"], "level": s["level"], "text": s["text"]}
                    for s in outline
                ],
                "references": build_references_list(sources, cited_all), "references_title": CS["heading"], "source_note": note["text"], "style": style,
                "title": next((s["heading"] for s in outline if s.get("level", 1) <= 1 and s.get("heading")), None),
            })

        # Chat mode (history present) and simple mode (first message, no history
        # yet) share the same citation and clarify-guardrail logic below — they
        # differ only in whether prior turns are folded into the prompt.
        user_id = request.current_user["id"] if getattr(request, "current_user", None) else None
        memory_on = _memory_on(request.current_user)
        memory_notes = get_user_memory(user_id) if memory_on else []
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
            sources, search_diag = _find(search_query)

            if sources:
                sources_block = build_sources_block(sources)
                system_prompt = (
                    f"{role_with_memory} {THINKING_INSTRUCTION} {MEMORY_INSTRUCTION} {CLARIFY_INSTRUCTION} You're in an ongoing "
                    "conversation with the student — below is the conversation so far, then "
                    "a list of real sources found for their latest message, each labeled "
                    f"with its exact {CS['name']} in-text citation, e.g. {CS['example']}, listed "
                    "most-recent-first. When your response is generating written academic "
                    "content (a draft, an explanation, a literature summary, an argument), "
                    "write the way you actually would — ground every factual or evidentiary "
                    "claim in one of these sources, using the EXACT citation key next to it, "
                    "copied exactly, never altering an author name or year, and never "
                    "inventing a citation not in this list. Every sentence you write must carry at least one of these citations — openings, links between ideas and conclusions too: a sentence that draws sources together cites each of them. Write no sentence you can't cite; leave it out instead. When more than one source could support "
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
                    f"{role_with_memory} {THINKING_INSTRUCTION} {MEMORY_INSTRUCTION} {CLARIFY_INSTRUCTION} "
                    + _no_sources_note(search_diag, "the student's latest message") + "\n\n"
                    f"CONVERSATION SO FAR:\n{history_block}"

                )
        else:
            sources, search_diag2 = _find(topic)

            if sources:
                sources_block = build_sources_block(sources)
                system_prompt = (
                    f"{role_with_memory} {THINKING_INSTRUCTION} {MEMORY_INSTRUCTION} {CLARIFY_INSTRUCTION} Below is a list of real "
                    f"sources found for this exact topic, each labeled with its exact {CS['name']} "
                    f"in-text citation, e.g. {CS['example']}, listed most-recent-first. When "
                    "your response is generating written academic content (a draft, an "
                    "explanation, a literature summary, an argument), write the way you "
                    "actually would — ground every factual or evidentiary claim in one of "
                    "these sources, using the EXACT citation key next to it, copied exactly, "
                    "never altering an author name or year, and never inventing a citation "
                    "not in this list. Every sentence you write must carry at least one of these citations — openings, links between ideas and conclusions too: a sentence that draws sources together cites each of them. Write no sentence you can't cite; leave it out instead. When more than one source could support the same "
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
                    f"{role_with_memory} {THINKING_INSTRUCTION} {MEMORY_INSTRUCTION} {CLARIFY_INSTRUCTION} "
                    + _no_sources_note(search_diag2, "this exact topic")
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
        if memory_note and memory_on:
            add_user_memory(user_id, memory_note)

        if not text:
            raise ConversionError("Couldn't generate a reply — try rephrasing")
        if not is_clarify and len(text) < 15:
            raise ConversionError("Couldn't generate a reply — try rephrasing")

        if is_clarify:
            return jsonify({"type": "clarify", "text": text, "thinking": thinking})

        text, cited = keep_known_citations(text, sources, style)
        text, cited = _c34_write(text, sources, style, topic, note, started)
        return jsonify({
            "type": "result",
            "text": text,
            "thinking": thinking,
            "references": build_references_list(sources, cited), "references_title": CS["heading"], "source_note": note["text"], "style": style,
            "title": _topic_title(topic) if not attachment else None,
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
            {"title": title, "text": text, "sections": sections, "references": references, "references_title": str(data.get("references_title") or "References")[:40],
             "numbered": data.get("numbered") is True},
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
        # The file's own design, kept (2.9). Rendering can fail on an unusual file: then it
        # opens in one of our designs, as before.
        original = None
        try:
            original = parse_pptx_original(path, work_dir, to_pdf=lambda src_path, out: _soffice_convert(src_path, "pdf", out))
        except Exception as e:
            logger.warning(f"Keeping the deck's own design failed: {e}")
        return jsonify({"deck": deck, "original": original})
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
    original = request.files.get("original")
    if original is not None:          # a deck that keeps its own design: changes go into the original file
        try:
            deck = json.loads(request.form.get("deck") or "")
        except Exception:
            deck = None
    else:
        deck = request.get_json(silent=True)
    if not isinstance(deck, dict):
        return jsonify({"error": "The slides couldn't be read — please try again."}), 400
    work_dir = tempfile.mkdtemp(prefix=f"mc_build_{uuid.uuid4().hex[:8]}_")
    try:
        if original is not None:
            orig_path = os.path.join(work_dir, "original.pptx")
            original.save(orig_path)
            out_dir = os.path.join(work_dir, "out")
            os.makedirs(out_dir, exist_ok=True)
            path = build_original_deck(deck, orig_path, out_dir, str(deck.get("filename") or "Presentation.pptx"))
        else:
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


# Remy inside Slide Studio: one tap to tighten a slide's points, put them in
# simpler words, draft points from its title, or write its speaker notes.
# Like Smart Summarize it's a slide tool, free on every plan, so it doesn't
# use the Remy message allowance — only the per-account AI rate limit.
_ASSIST_ACTIONS = {
    "tighten": ("Rewrite the slide's points so each is short and punchy — ideally under 12 words — keeping every idea "
                "and its meaning. Keep sub-points as sub-points (level 1). Use at most 7 points.", 700),
    "simplify": ("Rewrite the slide's points in plain, simple words a first-year student understands at once, keeping "
                 "the meaning. Keep roughly the same points and their levels.", 700),
    "expand": ("Write 3 to 5 clear, accurate points for this slide from its title, keeping and improving any points "
               "already there. Suitable for a class presentation or lecture.", 700),
    "notes": ("Write speaker notes for this slide: what the presenter says while it is shown — 2 to 4 short paragraphs "
              "of natural spoken English, about 80 to 150 words, covering each point.", 900),
}
_ASSIST_SYSTEM = (
    "You are Remy, the study assistant in Academic-Home, improving one slide of a student's or lecturer's presentation. "
    "Reply with one JSON object only — no preamble and no code fences. For points reply "
    '{"bullets": [{"text": "...", "level": 0}]} (level 0 for main points, 1 for sub-points); for speaker notes reply '
    '{"notes": "..."}. Only rework what the slide already says or clearly implies: never invent statistics, dates, '
    "citations or facts. Write in the slide's own language."
)


def _parse_assist_json(raw, action):
    cleaned = (raw or "").strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
        cleaned = re.sub(r"\s*```$", "", cleaned)
    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError:
        raise ConversionError("Remy couldn't improve this slide — please try again.")
    if not isinstance(data, dict):
        raise ConversionError("Remy couldn't improve this slide — please try again.")
    if action == "notes":
        notes = data.get("notes")
        if not isinstance(notes, str) or not notes.strip():
            raise ConversionError("Remy couldn't write notes for this slide — please try again.")
        return {"notes": notes.strip()[:2000]}
    points = []
    for b in data.get("bullets") or []:
        if isinstance(b, str):
            b = {"text": b, "level": 0}
        if not isinstance(b, dict):
            continue
        text = str(b.get("text") or "").strip()
        if not text:
            continue
        try:
            level = max(0, min(int(b.get("level") or 0), 3))
        except (TypeError, ValueError):
            level = 0
        points.append({"text": text[:400], "level": level})
        if len(points) >= 10:
            break
    if not points:
        raise ConversionError("Remy couldn't improve this slide — please try again.")
    return {"bullets": points}


@app.route("/api/slides/assist", methods=["POST", "OPTIONS"])
@limiter.limit(AI_LIMIT, key_func=_account_or_ip_key)
@auth_required
def slides_assist_endpoint():
    data = request.get_json(silent=True) or {}
    action = str(data.get("action") or "")
    if action not in _ASSIST_ACTIONS:
        return jsonify({"error": "Unknown request for Remy."}), 400
    slide = data.get("slide") if isinstance(data.get("slide"), dict) else {}
    title = str(slide.get("title") or "").strip()[:200]
    points = []
    for b in (slide.get("bullets") or [])[:40]:
        if isinstance(b, dict) and str(b.get("text") or "").strip():
            try:
                level = max(0, min(int(b.get("level") or 0), 3))
            except (TypeError, ValueError):
                level = 0
            points.append(("  " * level) + "- " + str(b["text"]).strip()[:600])
    if action in ("tighten", "simplify") and not points:
        return jsonify({"error": "Add some points to this slide first."}), 400
    if not title and not points:
        return jsonify({"error": "Give this slide a title first."}), 400
    deck = str((data.get("context") or {}).get("deck_title") or "").strip()[:160]
    instruction, max_tokens = _ASSIST_ACTIONS[action]
    user_message = (
        (f"Presentation: {deck}\n" if deck else "")
        + f"Slide title: {title or '(none)'}\n"
        + ("Points:\n" + "\n".join(points) + "\n" if points else "Points: (none yet)\n")
        + f"\nTask: {instruction}"
    )
    try:
        raw = call_claude(_ASSIST_SYSTEM, user_message, max_tokens=max_tokens)
        return jsonify(_parse_assist_json(raw, action))
    except ConversionError as e:
        return jsonify({"error": str(e)}), 502
    except Exception as e:
        logger.error(f"Slide assist failed unexpectedly: {e}", exc_info=True)
        return jsonify({"error": "Remy couldn't reach the AI just now — please try again."}), 502


# --------------------------------------------------------- drafting a whole deck
# Starting a new deck in Slide Studio no longer has to mean a blank slide:
# given a topic, Remy drafts a full outline (title slide plus content slides)
# that opens straight into the editor to fix up. It uses one Remy message per
# draft (the same allowance and Pay-per-Use messages as Remy chat), since a
# whole deck is a substantial AI request; editing the deck afterwards, and
# Remy's per-slide help, stay free.
_DRAFT_MAX_SLIDES = 10
_DRAFT_SYSTEM = (
    "You are Remy, the study assistant in Academic-Home, drafting a full slide-deck outline from a short topic someone "
    "gives you, for a {group} to open and edit. Reply with one JSON object only — no preamble, no code fences: "
    '{{"title_slide": {{...}}, "slides": [...]}}.\n'
    "{title_fields}"
    "slides: a JSON array, {min_slides} to " + str(_DRAFT_MAX_SLIDES - 1) + ' items, each one '
    '{{"kind": "content"|"columns"|"section", "title": "...", ...}} — '
    'for "content": "bullets": [{{"text": "...", "level": 0}}] (3 to 6 per slide, each under about 18 words, level 1 for '
    'a sub-point); for "columns": "left_heading", "left": [...same shape...], "right_heading", "right": [...]; for '
    '"section": a divider between parts of the talk, plus optional "sub". The LAST item in the array must be '
    '{{"kind": "end", "title": "Any questions" or "End", "big": "Thank you"}}.\n'
    "Rules: only develop what the topic already says or clearly implies — never invent statistics, dates, named "
    "studies, or quotations attributed to a real person; never use a \"quote\" or \"picture\" slide, since this tool "
    "has no real quotation or photo to put there; leave a title-slide field blank rather than inventing a course "
    "name, code, lecturer, or institution the person didn't mention; write in the topic's own language."
)
_DRAFT_TITLE_FIELDS = {
    "student": ('title_slide: {"lines": [up to 3 short strings — institution, department, programme; empty strings '
                'if not mentioned], "main": "the topic, as a slide title", "sub": "" (leave blank unless a '
                'presenter name was given)}.\n'),
    "tutor": ('title_slide: {"course": "the subject", "code": "" (leave blank unless given), "subtitle": "the '
              'lecture title", "lecturer": "" (leave blank unless given), "date": ""}.\n'),
}


def _clip_bullets(raw, limit=6):
    out = []
    for b in (raw or [])[:limit]:
        if isinstance(b, str):
            b = {"text": b, "level": 0}
        if not isinstance(b, dict):
            continue
        text = re.sub(r"\s+", " ", str(b.get("text") or "")).strip()
        if not text:
            continue
        try:
            level = max(0, min(int(b.get("level") or 0), 1))
        except (TypeError, ValueError):
            level = 0
        out.append({"text": text[:220], "level": level})
    return out


def _parse_draft_deck_json(raw, group):
    cleaned = (raw or "").strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
        cleaned = re.sub(r"\s*```$", "", cleaned)
    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError:
        data = None
    if not isinstance(data, dict):
        raise ConversionError("Remy couldn't draft this topic — please try again, or add a bit more detail.")

    ts_raw = data.get("title_slide") if isinstance(data.get("title_slide"), dict) else {}
    clip = lambda v, n: re.sub(r"\s+", " ", str(v or "")).strip()[:n]
    if group == "tutor":
        title_slide = {
            "course": clip(ts_raw.get("course"), 80), "code": clip(ts_raw.get("code"), 30),
            "subtitle": clip(ts_raw.get("subtitle"), 120), "lecturer": clip(ts_raw.get("lecturer"), 80),
            "date": clip(ts_raw.get("date"), 40),
        }
    else:
        lines = [clip(l, 120) for l in (ts_raw.get("lines") or []) if clip(l, 120)][:3]
        title_slide = {"lines": lines, "main": clip(ts_raw.get("main"), 200), "sub": clip(ts_raw.get("sub"), 200)}

    slides = []
    for sd in (data.get("slides") or [])[: _DRAFT_MAX_SLIDES]:
        if not isinstance(sd, dict):
            continue
        kind_raw = sd.get("kind")
        if kind_raw in ("quote", "picture"):
            continue  # never let these through — not even as an empty filler slide
        kind = kind_raw if kind_raw in ("content", "columns", "section", "end") else "content"
        title = clip(sd.get("title"), 140)
        if kind == "end":
            slides.append({"kind": "end", "title": title or ("Any questions" if group == "tutor" else "End"),
                            "big": clip(sd.get("big"), 60) or "Thank you"})
        elif kind == "section":
            slides.append({"kind": "section", "title": title, "sub": clip(sd.get("sub"), 200)})
        elif kind == "columns":
            slides.append({"kind": "columns", "title": title, "list_style": "bullets",
                            "left_heading": clip(sd.get("left_heading"), 60), "left": _clip_bullets(sd.get("left")),
                            "right_heading": clip(sd.get("right_heading"), 60), "right": _clip_bullets(sd.get("right"))})
        else:
            slides.append({"kind": "content", "title": title, "list_style": "bullets", "bullets": _clip_bullets(sd.get("bullets"))})

    if slides and slides[-1]["kind"] == "end":
        body, end = slides[:-1], slides[-1]
    else:
        body, end = slides, {"kind": "end", "title": "Any questions" if group == "tutor" else "End", "big": "Thank you"}
    if not body:
        raise ConversionError("Remy couldn't draft this topic — please try again, or add a bit more detail.")
    return {
        "template": ("tutor" if group == "tutor" else "student") + "-classic",
        "title_slide": title_slide, "slide_numbers": True, "slides": body[: _DRAFT_MAX_SLIDES - 1] + [end],
    }


# ------------------------------------------------ 3.0: drawings for slides
# A drawing arrives as a short list of shapes on a 1000 x 600 canvas (stick figures, lines,
# arrows, boxes, circles, labels) and is drawn here as a clear black-on-white picture.
_DRAW_W, _DRAW_H, _DRAW_SCALE = 1000, 600, 3
_DRAW_COLOURS = {"black": (17, 17, 17), "grey": (105, 105, 105), "gray": (105, 105, 105), "red": (196, 30, 30),
                 "blue": (25, 85, 200), "green": (20, 132, 60), "orange": (214, 110, 0)}
# a stick figure's joints in each pose, as (forward, up) in heights, feet at (0, 0)
_POSES = {
    "stand": {"head": (0, .91), "neck": (0, .80), "hip": (0, .45), "el1": (-.10, .62), "ha1": (-.13, .46), "el2": (.10, .62), "ha2": (.13, .46),
              "kn1": (-.05, .22), "ft1": (-.08, 0), "kn2": (.05, .22), "ft2": (.08, 0)},
    "walk": {"head": (.02, .91), "neck": (.02, .80), "hip": (0, .45), "el1": (-.10, .63), "ha1": (-.16, .50), "el2": (.12, .63), "ha2": (.18, .52),
             "kn1": (-.08, .23), "ft1": (-.16, 0), "kn2": (.10, .24), "ft2": (.14, 0)},
    "run": {"head": (.10, .88), "neck": (.08, .77), "hip": (0, .45), "el1": (-.10, .66), "ha1": (-.04, .56), "el2": (.18, .68), "ha2": (.26, .74),
            "kn1": (-.16, .30), "ft1": (-.26, .14), "kn2": (.18, .30), "ft2": (.16, .02)},
    "sit": {"head": (-.02, .68), "neck": (-.02, .57), "hip": (-.04, .25), "el1": (.06, .42), "ha1": (.16, .36), "el2": (.04, .44), "ha2": (.14, .38),
            "kn1": (.22, .27), "ft1": (.22, 0), "kn2": (.20, .27), "ft2": (.24, 0)},
    "squat": {"head": (.06, .66), "neck": (.05, .55), "hip": (-.08, .28), "el1": (.16, .50), "ha1": (.30, .52), "el2": (.14, .48), "ha2": (.28, .50),
              "kn1": (.14, .24), "ft1": (.02, 0), "kn2": (.12, .22), "ft2": (.06, 0)},
    "lift": {"head": (0, .91), "neck": (0, .80), "hip": (0, .45), "el1": (-.10, .90), "ha1": (-.08, 1.02), "el2": (.10, .90), "ha2": (.08, 1.02),
             "kn1": (-.06, .22), "ft1": (-.10, 0), "kn2": (.06, .22), "ft2": (.10, 0)},
    "push": {"head": (.12, .86), "neck": (.10, .75), "hip": (-.02, .44), "el1": (.24, .70), "ha1": (.38, .70), "el2": (.22, .66), "ha2": (.36, .66),
             "kn1": (-.02, .22), "ft1": (-.16, 0), "kn2": (.08, .22), "ft2": (.06, 0)},
    "pull": {"head": (-.10, .86), "neck": (-.08, .75), "hip": (.02, .44), "el1": (.10, .68), "ha1": (.26, .68), "el2": (.10, .64), "ha2": (.26, .64),
             "kn1": (.08, .22), "ft1": (.16, 0), "kn2": (-.02, .22), "ft2": (-.04, 0)},
    "reach": {"head": (.02, .91), "neck": (.02, .80), "hip": (0, .45), "el1": (-.10, .62), "ha1": (-.13, .46), "el2": (.16, .88), "ha2": (.30, .98),
              "kn1": (-.05, .22), "ft1": (-.08, 0), "kn2": (.05, .22), "ft2": (.08, 0)},
    "point": {"head": (0, .91), "neck": (0, .80), "hip": (0, .45), "el1": (-.10, .62), "ha1": (-.13, .46), "el2": (.17, .76), "ha2": (.34, .76),
              "kn1": (-.05, .22), "ft1": (-.08, 0), "kn2": (.05, .22), "ft2": (.08, 0)},
    "wave": {"head": (0, .91), "neck": (0, .80), "hip": (0, .45), "el1": (-.10, .62), "ha1": (-.13, .46), "el2": (.16, .80), "ha2": (.18, .98),
             "kn1": (-.05, .22), "ft1": (-.08, 0), "kn2": (.05, .22), "ft2": (.08, 0)},
    "jump": {"head": (0, 1.01), "neck": (0, .90), "hip": (0, .55), "el1": (-.12, 1.0), "ha1": (-.18, 1.12), "el2": (.12, 1.0), "ha2": (.18, 1.12),
             "kn1": (-.10, .34), "ft1": (-.06, .14), "kn2": (.10, .34), "ft2": (.06, .14)},
    "lie": {"head": (-.86, .07), "neck": (-.75, .05), "hip": (-.38, .05), "el1": (-.60, .02), "ha1": (-.46, .02), "el2": (-.62, .10), "ha2": (-.48, .12),
            "kn1": (-.18, .05), "ft1": (0, .03), "kn2": (-.18, .10), "ft2": (0, .08)},
}


def _draw_font(size):
    from PIL import ImageFont
    for p in ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
              "/usr/share/fonts/truetype/crosextra/Carlito-Regular.ttf", "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf"):
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                pass
    try:
        return ImageFont.load_default(size=size)
    except TypeError:
        return ImageFont.load_default()


def _num(v, lo=-200, hi=1200, default=0.0):
    try:
        v = float(v)
    except (TypeError, ValueError):
        return default
    return max(lo, min(hi, v)) if v == v else default


def render_drawing(spec):
    """A drawing spec ({"items": [...]}) as a PNG data URL; None when there is nothing to draw."""
    from PIL import Image, ImageDraw
    items = spec.get("items") if isinstance(spec, dict) else spec
    if not isinstance(items, list):
        return None
    S = _DRAW_SCALE
    img = Image.new("RGB", (_DRAW_W * S, _DRAW_H * S), (255, 255, 255))
    d = ImageDraw.Draw(img)
    drawn = 0
    for it in items[:60]:
        if not isinstance(it, dict):
            continue
        t = str(it.get("t") or it.get("type") or "").lower()
        col = _DRAW_COLOURS.get(str(it.get("color") or "black").lower(), _DRAW_COLOURS["black"])
        w = int(max(1, min(12, _num(it.get("w"), 1, 12, 4))) * S)
        fill = col if it.get("fill") else None
        try:
            if t == "person":
                x, y, h = _num(it.get("x"), default=500), _num(it.get("y"), default=560), _num(it.get("h"), 40, 600, 300)
                pose = _POSES.get(str(it.get("pose") or "stand").lower(), _POSES["stand"])
                s = -1 if str(it.get("face") or "right").lower() == "left" else 1
                P = lambda k: ((x + s * pose[k][0] * h) * S, (y - pose[k][1] * h) * S)
                r = h * 0.085 * S
                hx, hy = P("head")
                d.ellipse((hx - r, hy - r, hx + r, hy + r), outline=col, width=w)
                for a, b in (("neck", "hip"), ("neck", "el1"), ("el1", "ha1"), ("neck", "el2"), ("el2", "ha2"),
                             ("hip", "kn1"), ("kn1", "ft1"), ("hip", "kn2"), ("kn2", "ft2")):
                    d.line((P(a), P(b)), fill=col, width=w, joint="curve")
                for k in ("ha1", "ha2", "ft1", "ft2", "el1", "el2", "kn1", "kn2", "hip", "neck"):
                    cx, cy = P(k)
                    d.ellipse((cx - w / 2, cy - w / 2, cx + w / 2, cy + w / 2), fill=col)
                if it.get("label"):
                    f = _draw_font(int(22 * S))
                    d.text((x * S, (y + 30) * S), str(it["label"])[:40], fill=col, font=f, anchor="mm")
            elif t in ("line", "arrow"):
                p = [_num(v) * S for v in (it.get("p") or [])[:4]]
                if len(p) < 4:
                    continue
                d.line(p, fill=col, width=w)
                if t == "arrow" or it.get("head"):
                    import math
                    ang = math.atan2(p[3] - p[1], p[2] - p[0])
                    L = (16 + w / S * 2) * S
                    a1 = (p[2] - L * math.cos(ang - 0.45), p[3] - L * math.sin(ang - 0.45))
                    a2 = (p[2] - L * math.cos(ang + 0.45), p[3] - L * math.sin(ang + 0.45))
                    d.polygon([(p[2], p[3]), a1, a2], fill=col)
            elif t == "rect":
                x, y, rw, rh = [_num(v) for v in ((it.get("p") or [0, 0, 0, 0]) + [0, 0, 0, 0])[:4]]
                if rw <= 0 or rh <= 0:
                    continue
                d.rectangle((x * S, y * S, (x + rw) * S, (y + rh) * S), outline=col, width=w, fill=fill)
            elif t in ("circle", "ellipse"):
                c = it.get("c") or [500, 300]
                cx, cy = _num(c[0] if len(c) > 0 else 500), _num(c[1] if len(c) > 1 else 300)
                rx = _num(it.get("r") if t == "circle" else it.get("rx"), 1, 600, 30)
                ry = _num(it.get("r") if t == "circle" else it.get("ry"), 1, 600, 30)
                d.ellipse(((cx - rx) * S, (cy - ry) * S, (cx + rx) * S, (cy + ry) * S), outline=col, width=w, fill=fill)
            elif t in ("poly", "polygon", "polyline"):
                pts = [(_num(q[0]) * S, _num(q[1]) * S) for q in (it.get("pts") or [])[:40] if isinstance(q, (list, tuple)) and len(q) >= 2]
                if len(pts) < 2:
                    continue
                if it.get("closed", t != "polyline") and len(pts) >= 3:
                    d.polygon(pts, outline=col, fill=fill, width=w)
                else:
                    d.line(pts, fill=col, width=w, joint="curve")
            elif t == "text":
                p = it.get("p") or [500, 300]
                s_ = str(it.get("s") or it.get("text") or "")[:60]
                if not s_:
                    continue
                f = _draw_font(int(_num(it.get("size"), 12, 60, 26) * S))
                anchor = {"start": "lm", "left": "lm", "end": "rm", "right": "rm"}.get(str(it.get("align") or "").lower(), "mm")
                d.text((_num(p[0]) * S, _num(p[1] if len(p) > 1 else 300) * S), s_, fill=col, font=f, anchor=anchor)
            else:
                continue
            drawn += 1
        except Exception:
            continue
    if not drawn:
        return None
    img = img.resize((1500, 900), Image.LANCZOS)
    import io as _io
    buf = _io.BytesIO()
    img.save(buf, "PNG", optimize=True)
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode("ascii")


# ------------------------------------------------ 3.0: a complete deck from a few words
_DECK_MAX = 16
_DRAW_WORDS = re.compile(r"\b(draw\w*|diagram\w*|illustrat\w*|stick ?(?:man|men|figures?)|sketch\w*|picture\w*|visual\w*|figure\w*|image\w*)\b", re.I)
_CHART_WORDS = re.compile(r"\b(chart\w*|graph\w*|data|statistic\w*|figures?|numbers?)\b", re.I)
_DECK_ORDER_FIELDS = {
    "student": ('presenters: the presenters\' names with their index numbers, as strings, ONLY if the request gives them; '
                'otherwise [].\n'),
    "tutor": ('objectives: 3 to 5 learning objectives for the lecture, each starting with a verb ("Explain ...", "Compare ...") '
              'and at most 14 words.\ndiscussion: 2 or 3 short questions for a class discussion at the end.\n'),
}
_DECK_SYSTEM = (
    "You are Remy, the assistant in Academic-Home, writing a COMPLETE slide deck from a short request, for a {group} to present. "
    "Reply with ONE JSON object only — no preamble, no code fences: {{\"title_slide\": {{...}}, \"slides\": [...]}}.\n"
    "{title_fields}"
    "slides: {min_slides} to {max_slides} items in teaching order: an introduction or overview first, then the content, then key "
    "takeaways or a summary. Each item is one of:\n"
    '  {{"kind": "content", "title": "...", "bullets": [{{"text": "...", "level": 0}}], "notes": "..."}} — 3 to 5 points, each at most '
    "16 words (level 1 for a sub-point, at most 2 on a slide);\n"
    '  {{"kind": "columns", "title": "...", "left_heading": "...", "left": [...], "right_heading": "...", "right": [...], "notes": "..."}} '
    "— for a comparison, 2 to 4 points a side;\n"
    '  {{"kind": "section", "title": "...", "sub": "..."}} — a divider between big parts, only in decks of 10 or more slides.\n'
    "Every title names its slide's real topic (never \"Slide 3\", \"Continued\" or \"Overview\" twice). \"notes\": 2 or 3 sentences "
    "of what to say on that slide. Don't add a cover, a presenters slide, an outline, objectives, a discussion, a references "
    "slide or a thank-you slide as slides — they are added for you, in their places.\n"
    "Rules: never invent statistics, dates, named studies or quotations; write in the request's own language; leave a "
    "title-slide field blank rather than inventing a course, code, lecturer or institution the request didn't mention."
)
_DECK_EVIDENCE = (
    "\n\nEVIDENCE: base the content on these real, open-access papers first. When a point uses a paper, end that point with "
    "the paper's exact citation key, e.g. \"Short naps improve recall {example}\". Use only these keys, copied exactly. Every "
    "point that states a fact carries the key of the paper behind it; leave out a fact no paper here supports \u2014 make fewer points "
    "rather than uncited ones (a title, a section divider, an agenda or a closing slide needs no key).\nPAPERS:\n{papers}"
)
_DECK_DRAWINGS = (
    "\n\nDRAWINGS: the request asks for drawings or visuals, so give the slides that need one (an experiment with stick "
    "figures, a labelled diagram, a simple process) a \"drawing\": {\"items\": [...]} on a 1000 x 600 canvas (x to the right, y "
    "down), black lines on white, at most 40 items:\n"
    '  {"t": "person", "x": 300, "y": 540, "h": 300, "pose": "stand"|"walk"|"run"|"sit"|"squat"|"lift"|"push"|"pull"|"reach"|"point"|"wave"|"jump"|"lie", '
    '"face": "left"|"right", "label": "optional"} — a stick figure, (x, y) where its feet touch the ground;\n'
    '  {"t": "line", "p": [x1, y1, x2, y2]}, {"t": "arrow", "p": [x1, y1, x2, y2]}, {"t": "rect", "p": [x, y, w, h]}, '
    '{"t": "circle", "c": [x, y], "r": 30}, {"t": "ellipse", "c": [x, y], "rx": 60, "ry": 30}, '
    '{"t": "poly", "pts": [[x, y], ...], "closed": true}, {"t": "text", "p": [x, y], "s": "a short label", "size": 26}.\n'
    'Any item may add "w" (line width 2-10), "fill": true, or "color": "red"|"blue"|"green"|"orange"|"grey" to pick something '
    "out. Draw the ground as a line under people and objects; keep labels short, beside what they name and clear of the lines. "
    "A slide with a drawing has at most 3 points."
)
_DECK_CHARTS = (
    "\n\nCHARTS: when a paper's abstract above states real numbers worth comparing, a slide may add \"chart\": {\"chart_type\": "
    "\"column\"|\"bar\"|\"line\"|\"pie\", \"title\": \"...\", \"categories\": [...], \"series\": [{\"name\": \"...\", \"values\": [numbers]}], "
    "\"from\": \"the citation key of that paper\"} — every number exactly as that abstract states it, never estimated."
)


def _deck_points(raw, limit=8):
    out = []
    for b in (raw or [])[:limit]:
        if isinstance(b, str):
            b = {"text": b, "level": 0}
        if not isinstance(b, dict):
            continue
        text = re.sub(r"\s+", " ", str(b.get("text") or "")).strip()
        if not text:
            continue
        try:
            level = max(0, min(int(b.get("level") or 0), 1))
        except (TypeError, ValueError):
            level = 0
        out.append({"text": text[:260], "level": level})
    return out


def _chart_from_paper(spec, sources):
    """A chart only when every number in it is stated in the abstract of the paper it cites."""
    if not isinstance(spec, dict):
        return None
    key = str(spec.get("from") or "").strip()
    paper = next((s for s in sources if s.get("citation_key") == key), None)
    abstract = (paper or {}).get("abstract") or ""
    cats = [str(c)[:40] for c in (spec.get("categories") or [])][:12]
    series = []
    for s in (spec.get("series") or [])[:4]:
        if not isinstance(s, dict):
            continue
        vals = []
        for v in (s.get("values") or [])[:len(cats)]:
            try:
                vals.append(float(v))
            except (TypeError, ValueError):
                return None
        series.append({"name": str(s.get("name") or "Value")[:40], "values": vals})
    if not paper or not cats or not series or any(len(s["values"]) != len(cats) for s in series):
        return None
    shown = set(re.findall(r"\d+(?:\.\d+)?", abstract.replace(",", "")))
    for s in series:
        for v in s["values"]:
            forms = {f"{v:g}"} | ({str(int(v))} if v == int(v) else set())
            if not forms & shown:          # a number the paper doesn't state: no chart
                return None
    kind = spec.get("chart_type") if spec.get("chart_type") in ("column", "bar", "line", "pie") else "column"
    return {"type": "chart", "chart_type": kind, "categories": cats, "series": series, "title": str(spec.get("title") or "")[:80]}


def _parse_full_deck(raw, group, sources, style="apa", deck_max=None):
    """The AI's deck as Slide Studio's deck: points checked against the papers, more than 5 points
    split onto a second slide, drawings drawn, charts kept only with real numbers, a References
    slide for the papers cited, and the closing slide."""
    data = None
    cleaned = re.sub(r"^```(?:json)?\s*|\s*```$", "", (raw or "").strip())
    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError:
        a, b = cleaned.find("{"), cleaned.rfind("}")
        if a >= 0 and b > a:
            try:
                data = json.loads(cleaned[a:b + 1])
            except json.JSONDecodeError:
                data = None
    if not isinstance(data, dict):
        raise ConversionError("Remy couldn't write this deck — please try again, or add a bit more detail.")
    ts_raw = data.get("title_slide") if isinstance(data.get("title_slide"), dict) else {}
    clip = lambda v, n: re.sub(r"\s+", " ", str(v or "")).strip()[:n]
    if group == "tutor":
        title_slide = {"course": clip(ts_raw.get("course"), 80), "code": clip(ts_raw.get("code"), 30), "subtitle": clip(ts_raw.get("subtitle"), 120),
                       "lecturer": clip(ts_raw.get("lecturer"), 80), "date": clip(ts_raw.get("date"), 40)}
    else:
        lines = [clip(l, 120) for l in (ts_raw.get("lines") or []) if clip(l, 120)][:3]
        title_slide = {"lines": lines, "main": clip(ts_raw.get("main"), 200), "sub": clip(ts_raw.get("sub"), 200)}
    cited = set()

    def check(points):
        for b in points:
            b["text"], c = keep_known_citations(b["text"], sources, style)
            cited.update(c)
        return [b for b in points if b["text"].strip()]

    slides = []
    deck_max = deck_max or _DECK_MAX
    for sd in (data.get("slides") or [])[:deck_max + 4]:
        if not isinstance(sd, dict):
            continue
        kind = sd.get("kind") if sd.get("kind") in ("content", "columns", "section") else "content"
        title = clip(sd.get("title"), 140)
        notes = clip(sd.get("notes"), 900)
        if kind == "section":
            slides.append({"kind": "section", "title": title, "sub": clip(sd.get("sub"), 200)})
            continue
        if kind == "columns":
            slides.append({"kind": "columns", "title": title, "list_style": "bullets", "notes": notes,
                           "left_heading": clip(sd.get("left_heading"), 60), "left": check(_deck_points(sd.get("left"), 5)),
                           "right_heading": clip(sd.get("right_heading"), 60), "right": check(_deck_points(sd.get("right"), 5))})
            continue
        points = check(_deck_points(sd.get("bullets")))
        media = None
        if sd.get("drawing"):
            url = render_drawing(sd.get("drawing"))
            if url:
                media = {"type": "image", "data": url, "fit": "whole"}
        if media is None and sd.get("chart"):
            media = _chart_from_paper(sd.get("chart"), sources)
        if not title and not points and not media:
            continue
        per = 3 if media else 5           # big text: split rather than shrink
        chunks = [points[i:i + per] for i in range(0, len(points), per)] or [[]]
        for k, chunk in enumerate(chunks):
            s = {"kind": "content", "title": title, "list_style": "bullets", "bullets": chunk, "notes": notes if k == 0 else ""}
            if media and k == 0:
                s["media"] = media
                s["media_position"] = "right" if chunk else "full"
            slides.append(s)
    body = slides[:deck_max]
    if not any(s["kind"] != "section" for s in body):
        raise ConversionError("Remy couldn't write this deck — please try again, or add a bit more detail.")
    refs = build_references_list(sources, cited) if sources else []
    lines = lambda rs: [re.sub(r"[*_]", "", r["text"]) + (" " + r["url"] if r.get("url") else "") for r in rs]

    def plain(title, items, per=5):
        return [{"kind": "content", "title": title, "list_style": "bullets", "bullets": [{"text": t, "level": 0} for t in items[i:i + per]]}
                for i in range(0, len(items), per)]
    # the outline is the deck's own parts: its sections when it has them, else its slide titles
    sections = [s["title"] for s in body if s["kind"] == "section" and s.get("title")]
    outline = sections if len(sections) >= 2 else list(dict.fromkeys(s["title"] for s in body if s["kind"] in ("content", "columns") and s.get("title")))
    items = lambda key, n, cap: [clip(x, cap) for x in (data.get(key) or []) if isinstance(x, (str, int, float)) and clip(x, cap)][:n]
    if group == "tutor":
        # every paper found is worth reading (those the slides cite, and the others), all of them real
        reading = lines(build_references_list(sources, None)) if sources else []
        out = plain("References & Recommended Books", reading, 4) + plain("Learning Objectives", items("objectives", 6, 160))
        out += plain("Presentation Outline", outline, 6) + body + plain("Discussion", items("discussion", 4, 200))
        out.append({"kind": "end", "title": "Thank you", "big": "Thank you"})
    else:
        out = plain("Presented by", items("presenters", 10, 120) or ["Name \u2014 Index number"], 6)
        out += plain("Presentation Outline", outline, 6) + body + plain("References", lines(refs), 4)
    return {"template": ("tutor" if group == "tutor" else "student") + "-plain", "title_slide": title_slide, "slide_numbers": True,
            "slides": out}, refs


@app.route("/api/slides/draft", methods=["POST", "OPTIONS"])
@limiter.limit(AI_LIMIT, key_func=_account_or_ip_key)
@auth_required
def slides_draft_endpoint():
    data = request.get_json(silent=True) or {}
    group = "tutor" if data.get("group") == "tutor" else "student"
    topic = re.sub(r"\s+", " ", str(data.get("topic") or "")).strip()[:300]
    if not topic:
        return jsonify({"error": "Tell Remy what the deck is about first."}), 400
    pro = _is_pro(request.current_user)
    content_max = DECK_CONTENT_PRO if pro else DECK_CONTENT_FREE   # room for 2 section dividers on top
    # Evidence first: real papers for the topic, and the request's own wishes (drawings, charts).
    try:
        papers, _ = paper_topic_search(topic, max_results=6)
    except Exception as e:
        logger.warning(f"Deck paper search failed: {e}")
        papers = []
    sources = enrich_sources(papers, "apa") if papers else []
    system = _DECK_SYSTEM.format(
        group="tutor preparing a lecture" if group == "tutor" else "student preparing a presentation or report",
        title_fields=_DRAFT_TITLE_FIELDS[group] + _DECK_ORDER_FIELDS[group], min_slides=8, max_slides=content_max,
    )
    if sources:
        system += _DECK_EVIDENCE.format(example=sources[0]["citation_key"], papers="\n".join(
            f"{s['citation_key']} {s['title']} \u2014 {s.get('venue') or 'paper'}, {s.get('year') or 'n.d.'}"
            + (f"\n    Abstract: {s['abstract'][:600]}" if s.get("abstract") else "") for s in sources))
    if _DRAW_WORDS.search(topic):
        system += _DECK_DRAWINGS
    if sources and _CHART_WORDS.search(topic):
        system += _DECK_CHARTS
    # A whole drafted deck is real AI work, so it draws on the same Remy
    # messages as chatting with Remy: the month's allowance first, then any
    # Pay-per-Use messages. Checked before the AI call (never spend on a
    # request already over its limit), and given back below if no usable
    # deck arrives — a message only counts when the answer does.
    try:
        charged_from = _log_sage_message_and_consume(request.current_user)
    except ConversionError as e:
        return jsonify({"error": str(e)}), 429
    try:
        raw = call_claude(system, f"Request: {topic}", max_tokens=12000 if pro else 9000, timeout=150 if pro else None, feature="deck")
        deck, refs = _parse_full_deck(raw, group, sources, deck_max=content_max + 2)
        # named after the deck's own title — never the request
        ts = deck["title_slide"]
        name = (ts.get("main") or ts.get("subtitle") or ts.get("course") or _topic_title(topic) or "").strip()
        deck["filename"] = (re.sub(r"[^\w \-]+", "", name).strip()[:60] or ("Lecture" if group == "tutor" else "Presentation"))
        return jsonify({"deck": deck, "references": refs})
    except ConversionError as e:
        _refund_sage_message(request.current_user, charged_from)
        return jsonify({"error": str(e) + AI_MSG_NOT_COUNTED}), 502
    except Exception as e:
        _refund_sage_message(request.current_user, charged_from)
        logger.error(f"Slide draft failed unexpectedly: {e}", exc_info=True)
        return jsonify({"error": AI_MSG_FAILED + AI_MSG_NOT_COUNTED}), 502


# ------------------------------------------- the stores' notices to this server
# When a subscription renews, lapses, is cancelled or refunded, Apple and
# Google tell this server. A notice is never trusted on its own word: it
# only says which purchase to look at, and the server then asks the store
# itself — over the same authenticated APIs used to verify purchases — what
# that purchase's state is now. A forged notice can at most cause a harmless
# re-check. When the store can't be reached, nothing is changed and the
# notice is answered 503, so the store sends it again later.
#   Apple:  App Store Connect -> App Information -> App Store Server
#           Notifications (Version 2): https://<server>/api/store/apple/notifications
#   Google: Play Console -> Monetisation setup -> Real-time developer
#           notifications, through a Pub/Sub push subscription to
#           https://<server>/api/store/google/notifications?token=<GOOGLE_RTDN_TOKEN>
GOOGLE_RTDN_TOKEN = os.environ.get("GOOGLE_RTDN_TOKEN", "").strip()


def _apple_api_token():
    now = int(time.time())
    return pyjwt.encode(
        {"iss": APPLE_APP_STORE_ISSUER_ID, "iat": now, "exp": now + 300, "aud": "appstoreconnect-v1", "bid": APPLE_BUNDLE_ID},
        APPLE_APP_STORE_PRIVATE_KEY, algorithm="ES256", headers={"kid": APPLE_APP_STORE_KEY_ID, "typ": "JWT"},
    )


def _apple_subscription_entitled(original_transaction_id):
    """True while Apple says the subscription is active or in its billing
    grace period; False once it has expired, is stuck in billing retry, or
    was revoked (refunded); None when Apple can't be asked right now."""
    if not (APPLE_APP_STORE_KEY_ID and APPLE_APP_STORE_ISSUER_ID and APPLE_APP_STORE_PRIVATE_KEY):
        return None
    host = ("https://api.storekit-sandbox.itunes.apple.com" if APPLE_APP_STORE_ENVIRONMENT == "sandbox"
            else "https://api.storekit.itunes.apple.com")
    try:
        resp = requests.get(f"{host}/inApps/v1/subscriptions/{original_transaction_id}",
                            headers={"Authorization": f"Bearer {_apple_api_token()}"}, timeout=15)
    except requests.RequestException:
        return None
    if resp.status_code != 200:
        return None
    statuses = []
    for group in (resp.json().get("data") or []):
        for tx in (group.get("lastTransactions") or []):
            if str(tx.get("originalTransactionId")) == str(original_transaction_id):
                statuses.append(tx.get("status"))
    if not statuses:
        return None
    ends = []
    for group in (resp.json().get("data") or []):
        for tx in (group.get("lastTransactions") or []):
            if str(tx.get("originalTransactionId")) == str(original_transaction_id):
                try:
                    info = pyjwt.decode(tx.get("signedTransactionInfo", ""), options={"verify_signature": False})
                    if info.get("expiresDate"):
                        ends.append(datetime.fromtimestamp(info["expiresDate"] / 1000, timezone.utc))
                except Exception:
                    pass
    _STORE_PERIOD_END[str(original_transaction_id)] = max(ends) if ends else None
    return any(s in (1, 4) for s in statuses)   # 1 active, 4 grace period; 2 expired, 3 billing retry, 5 revoked


# The end of the paid period the store reported on its last check, by original
# transaction id (Apple) or purchase token (Google).
_STORE_PERIOD_END = {}


def _google_access_token():
    service_account = json.loads(GOOGLE_PLAY_SERVICE_ACCOUNT_JSON)
    now = int(time.time())
    assertion = pyjwt.encode(
        {"iss": service_account["client_email"], "scope": "https://www.googleapis.com/auth/androidpublisher",
         "aud": "https://oauth2.googleapis.com/token", "iat": now, "exp": now + 3600},
        service_account["private_key"], algorithm="RS256",
    )
    token_resp = requests.post("https://oauth2.googleapis.com/token",
                               data={"grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer", "assertion": assertion},
                               timeout=15)
    token_resp.raise_for_status()
    return token_resp.json()["access_token"]


def _google_subscription_entitled(purchase_token):
    """True while Google says the subscription gives access — active, in its
    grace period, or cancelled but still inside the period already paid for;
    False once it has lapsed; None when Google can't be asked right now."""
    if not (GOOGLE_PLAY_SERVICE_ACCOUNT_JSON and GOOGLE_PLAY_PACKAGE_NAME):
        return None
    try:
        token = _google_access_token()
        resp = requests.get(
            f"https://androidpublisher.googleapis.com/androidpublisher/v3/applications/{GOOGLE_PLAY_PACKAGE_NAME}"
            f"/purchases/subscriptionsv2/tokens/{purchase_token}",
            headers={"Authorization": f"Bearer {token}"}, timeout=15)
    except (requests.RequestException, KeyError, ValueError):
        return None
    if resp.status_code != 200:
        return None
    data = resp.json()
    state = data.get("subscriptionState")
    ends = []
    for item in (data.get("lineItems") or []):
        try:
            ends.append(datetime.fromisoformat(str(item.get("expiryTime", "")).replace("Z", "+00:00")))
        except ValueError:
            pass
    _STORE_PERIOD_END[str(purchase_token)] = max(ends) if ends else None
    if state in ("SUBSCRIPTION_STATE_ACTIVE", "SUBSCRIPTION_STATE_IN_GRACE_PERIOD"):
        return True
    if state == "SUBSCRIPTION_STATE_CANCELED":
        # Cancelled means "won't renew" — access continues until the paid period ends.
        for item in (data.get("lineItems") or []):
            try:
                expiry = datetime.fromisoformat(str(item.get("expiryTime", "")).replace("Z", "+00:00"))
            except ValueError:
                continue
            if expiry > datetime.now(timezone.utc):
                return True
        return False
    if state in ("SUBSCRIPTION_STATE_EXPIRED", "SUBSCRIPTION_STATE_ON_HOLD", "SUBSCRIPTION_STATE_PAUSED"):
        return False
    return None   # pending, unknown: leave things as they are


def _store_purchase_owner(keys):
    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT user_id, product FROM store_purchases WHERE transaction_key = ANY(%s) LIMIT 1", (list(keys),))
            return cur.fetchone()
    finally:
        conn.close()


def _apply_store_state(owner, entitled=None, refunded_pass=False, period_end=None):
    """Brings one account in line with what the store said. Only the plan the
    notice is about is touched: a lapsed Monthly never downgrades someone who
    has since moved to Yearly, and access returns only to an account that
    lost it (it's on Free), never overriding a different plan."""
    conn = get_db()
    try:
        with conn.cursor() as cur:
            if refunded_pass:
                cur.execute("UPDATE users SET remy_credits = GREATEST(COALESCE(remy_credits, 0) - %s, 0) WHERE id = %s",
                            (PAYPERUSE_MESSAGES, owner["user_id"]))
            elif entitled is False:
                cur.execute("UPDATE users SET plan = 'free', plan_expires_at = NULL, plan_source = NULL WHERE id = %s AND plan = %s", (owner["user_id"], owner["product"]))
            elif entitled is True:
                # Renewed: the plan runs to the store's new period end (plus a
                # day), never shortening a later date already there.
                until = (period_end + timedelta(days=1)) if period_end else (datetime.now(timezone.utc) + timedelta(days=2))
                cur.execute("UPDATE users SET plan = %s, plan_expires_at = GREATEST(COALESCE(plan_expires_at, %s), %s), plan_source = 'store' "
                            "WHERE id = %s AND (plan = 'free' OR plan = %s)",
                            (owner["product"], until, until, owner["user_id"], owner["product"]))
        conn.commit()
    finally:
        conn.close()


@app.route("/api/store/apple/notifications", methods=["POST"])
@limiter.limit("120 per minute")
def apple_store_notification():
    body = request.get_json(silent=True) or {}
    try:
        payload = pyjwt.decode(body.get("signedPayload", ""), options={"verify_signature": False})
        tx = pyjwt.decode((payload.get("data") or {}).get("signedTransactionInfo", ""), options={"verify_signature": False})
    except Exception:
        return jsonify({"ok": True, "ignored": "unreadable notification"}), 200
    original, current = str(tx.get("originalTransactionId") or ""), str(tx.get("transactionId") or "")
    owner = _store_purchase_owner([k for k in ("ios:" + original, "ios:" + current) if k != "ios:"])
    if not owner:
        return jsonify({"ok": True, "ignored": "unknown purchase"}), 200
    if owner["product"] == "payperuse":
        if payload.get("notificationType") == "REFUND":
            _apply_store_state(owner, refunded_pass=True)
        return jsonify({"ok": True}), 200
    entitled = _apple_subscription_entitled(original or current)   # asked of Apple, never taken from the notice
    if entitled is None:
        return jsonify({"error": "Couldn't check with Apple right now."}), 503
    _apply_store_state(owner, entitled=entitled, period_end=_STORE_PERIOD_END.pop(str(original or current), None))
    return jsonify({"ok": True, "entitled": entitled}), 200


@app.route("/api/store/google/notifications", methods=["POST"])
@limiter.limit("120 per minute")
def google_store_notification():
    if not GOOGLE_RTDN_TOKEN:
        return jsonify({"error": "Google Play notifications aren't configured (GOOGLE_RTDN_TOKEN)."}), 503
    if not hmac.compare_digest(request.args.get("token", ""), GOOGLE_RTDN_TOKEN):
        return jsonify({"error": "Not allowed."}), 403
    try:
        message = (request.get_json(silent=True) or {}).get("message") or {}
        note = json.loads(base64.b64decode(message.get("data", "")).decode("utf-8"))
    except Exception:
        return jsonify({"ok": True, "ignored": "unreadable notification"}), 200
    sub_note, voided = note.get("subscriptionNotification"), note.get("voidedPurchaseNotification")
    token = (sub_note or voided or {}).get("purchaseToken")
    if not token:
        return jsonify({"ok": True, "ignored": "nothing to act on"}), 200   # e.g. a test notification
    owner = _store_purchase_owner(["android:" + token])
    if not owner:
        return jsonify({"ok": True, "ignored": "unknown purchase"}), 200
    if owner["product"] == "payperuse":
        if voided:
            _apply_store_state(owner, refunded_pass=True)
        return jsonify({"ok": True}), 200
    entitled = False if voided else _google_subscription_entitled(token)   # asked of Google, never taken from the notice
    if entitled is None:
        return jsonify({"error": "Couldn't check with Google Play right now."}), 503
    _apply_store_state(owner, entitled=entitled, period_end=_STORE_PERIOD_END.pop(str(token), None))
    return jsonify({"ok": True, "entitled": entitled}), 200


# ------------------------------------------------ colour, not bold, in slides
def _slide_is_dark(slide):
    """Whether a slide's background is dark (its own, else its layout's or
    master's solid fill); white when it can't be told."""
    for owner in (slide, slide.slide_layout, slide.slide_layout.slide_master):
        try:
            fill = owner.background.fill
            if fill.type == 1:   # solid
                rgb = fill.fore_color.rgb
                r, g, b = int(str(rgb)[0:2], 16), int(str(rgb)[2:4], 16), int(str(rgb)[4:6], 16)
                return (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255 < 0.45
        except Exception:
            continue
    return False


def _colour_not_bold(pptx_path):
    """Presentations never use bold for emphasis — it reads heavily and is
    hard to make out from the back of a room. Every bold run in a finished
    deck becomes regular weight and, unless it already has its own colour,
    takes a clear emphasis colour picked against its slide's background: a
    strong blue on light slides, a soft gold on dark ones. Titles keep the
    colours they have."""
    from pptx import Presentation as _Prs
    from pptx.dml.color import RGBColor as _RGB
    from pptx.enum.shapes import PP_PLACEHOLDER as _PH
    prs = _Prs(pptx_path)
    changed = False

    def text_shapes(shapes):
        for sh in shapes:
            if getattr(sh, "shape_type", None) == 6 and hasattr(sh, "shapes"):   # group
                yield from text_shapes(sh.shapes)
            elif getattr(sh, "has_text_frame", False):
                yield sh
            elif getattr(sh, "has_table", False):
                for row in sh.table.rows:
                    for cell in row.cells:
                        yield cell

    for slide in prs.slides:
        emph = _RGB(0xFD, 0xE6, 0x8A) if _slide_is_dark(slide) else _RGB(0x1D, 0x4E, 0xD8)
        for sh in text_shapes(slide.shapes):
            is_title = bool(getattr(sh, "is_placeholder", False) and sh.placeholder_format.type in (_PH.TITLE, _PH.CENTER_TITLE))
            in_table = not hasattr(sh, "shape_id")
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if r.font.bold:
                        r.font.bold = False
                        changed = True
                        if not is_title and not in_table and r.font.color.type is None:
                            r.font.color.rgb = emph
    if changed:
        prs.save(pptx_path)
    return changed


# ------------------------------------------------------------ deleting an account
# Required by the App Store and Google Play for any app where people create
# an account: deleting it from inside the app. Everything stored with the
# account goes with it — conversion history and Remy's memory notes (removed
# by the database with the account), and the store purchase records; promo
# codes someone created stay usable but no longer point at them. The
# password is asked again, so a borrowed phone that's already signed in
# can't do it by accident. (Store subscriptions are billed by Apple or
# Google, so the app reminds people to cancel those in the store.)
@app.route("/api/account", methods=["DELETE", "OPTIONS"])
@limiter.limit("5 per hour", key_func=_account_or_ip_key)
@auth_required
def delete_account_endpoint():
    data = request.get_json(silent=True) or {}
    user = request.current_user
    stored = user.get("password_hash") if hasattr(user, "get") else None
    if stored:
        if not check_password_hash(stored, str(data.get("password") or "")):
            return jsonify({"error": "That password isn't right — your account hasn't been deleted."}), 403
    elif str(data.get("confirm") or "").strip().upper() != "DELETE":
        return jsonify({"error": "Type DELETE to confirm."}), 400
    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM store_purchases WHERE user_id = %s", (user["id"],))
            cur.execute("DELETE FROM users WHERE id = %s", (user["id"],))
        conn.commit()
    except Exception as e:
        conn.rollback()
        logger.error(f"Account deletion failed for user {user['id']}: {e}", exc_info=True)
        return jsonify({"error": "Your account couldn't be deleted just now — please try again."}), 500
    finally:
        conn.close()
    logger.info(f"Account {user['id']} deleted at the user's request")
    return jsonify({"ok": True})


# ------------------------------------------------------------ app error reports
# When the app hits an unexpected error on someone's phone, it sends a short
# technical report here so problems on real devices get noticed and fixed.
# Only technical facts are accepted: the error message, where in the app it
# happened, the app version and the browser type — no names, emails or file
# content (the privacy policy's "basic technical logs"). Reports are written
# to the server log, capped per phone, and trimmed to fixed lengths.
@app.route("/api/client-error", methods=["POST", "OPTIONS"])
@limiter.limit("20 per hour")
def client_error_report():
    data = request.get_json(silent=True) or {}
    clip = lambda v, n: re.sub(r"\s+", " ", str(v or ""))[:n]
    message = clip(data.get("message"), 300)
    if not message:
        return jsonify({"ok": True})
    where = clip(data.get("where"), 200)
    stack = clip(data.get("stack"), 900)
    version = clip(data.get("version"), 20)
    agent = clip(request.headers.get("User-Agent"), 160)
    logger.warning(f"[app error] v{version} | {message} | at {where} | {stack} | {agent}")
    return jsonify({"ok": True})

# ============================================================ 3.4: Academic-Home Meet
# Sound and video go through LiveKit; this server is the front desk (see patches/server_meet34.py).
def _meet_env_int(name, default):
    try:
        return max(1, int(os.environ.get(name, default)))
    except (TypeError, ValueError):
        return default


LIVEKIT_URL = os.environ.get("LIVEKIT_URL", "").strip().rstrip("/")
LIVEKIT_API_KEY = os.environ.get("LIVEKIT_API_KEY", "").strip()
LIVEKIT_API_SECRET = os.environ.get("LIVEKIT_API_SECRET", "").strip()
LIVEKIT_HTTP = re.sub(r"^ws", "http", LIVEKIT_URL)                  # wss://x → https://x for the server API
MEET_MAX_PEOPLE = _meet_env_int("MEET_MAX_PEOPLE", 100)             # in one meeting
MEET_GLOBAL_MAX = _meet_env_int("MEET_GLOBAL_MAX", 100)             # across every meeting at once (LiveKit Build: 100)
MEET_MONTH_MINUTES = _meet_env_int("MEET_MONTH_MINUTES", 5000)      # participant-minutes a month (LiveKit Build: 5,000)
MEET_FREE_MINUTES = _meet_env_int("MEET_FREE_MINUTES", 60)          # how long a meeting lasts, hosted on Free
MEET_PRO_MINUTES = _meet_env_int("MEET_PRO_MINUTES", 300)           # … and on Pro
MEET_AUDIENCE_OVER = _meet_env_int("MEET_AUDIENCE_OVER", 30)        # bigger than this, newcomers listen until let speak
MEET_KEEP_DAYS = 90                                                  # meetings and attendance are deleted after this
MEET_LIVE_PER_HOST = 3                                               # meetings one person can have running at once
MEET_TOKEN_TTL = 600                                                 # a pass is good for joining within 10 minutes
MEET_WAIT_MINUTES = 30                                               # how long someone can wait in the lobby

MEET_MSG_OFF = "Meetings aren’t switched on yet."
MEET_MSG_DOWN = "Meetings can’t be reached right now. Please try again in a moment."
MEET_MSG_BUSY = "Meetings are very busy right now. Please try again in a few minutes."
MEET_MSG_MONTH = "Meetings have reached this month’s limit. They’ll be back next month."
MEET_MSG_NOCODE = "That meeting code doesn’t exist. Check it and try again."
MEET_MSG_ENDED = "This meeting has ended."
MEET_MSG_LOCKED = "The host has locked this meeting."
MEET_MSG_PASS = "That passcode isn’t right."
MEET_MSG_FULL = "This meeting is full ({n} people)."
MEET_MSG_DENIED = "The host didn’t let you in."
MEET_MSG_REMOVED = "The host removed you from this meeting."
MEET_LETTERS = "abcdefghjkmnpqrstuvwxyz"                             # no i, l or o: nothing to mix up when read aloud
MEET_ROLES = ("host", "cohost", "speaker", "attendee", "listener")   # listener: a big meeting's newcomer, until let speak


class MeetError(Exception):
    def __init__(self, message, status=400, **extra):
        super().__init__(message)
        self.status, self.extra = status, extra


def _meet_on():
    return bool(LIVEKIT_URL and LIVEKIT_API_KEY and LIVEKIT_API_SECRET and DATABASE_URL)


def _meet_tables(cur):
    cur.execute("""
        CREATE TABLE IF NOT EXISTS meetings (
            id SERIAL PRIMARY KEY,
            code TEXT UNIQUE NOT NULL,
            room TEXT UNIQUE NOT NULL,
            host_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            title TEXT NOT NULL DEFAULT '',
            passcode_hash TEXT,
            lobby BOOLEAN NOT NULL DEFAULT TRUE,
            locked BOOLEAN NOT NULL DEFAULT FALSE,
            e2ee BOOLEAN NOT NULL DEFAULT FALSE,
            speak TEXT NOT NULL DEFAULT 'all',
            chat BOOLEAN NOT NULL DEFAULT TRUE,
            max_people INTEGER NOT NULL DEFAULT 100,
            limit_minutes INTEGER NOT NULL DEFAULT 60,
            created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
            started_at TIMESTAMPTZ,
            ends_at TIMESTAMPTZ,
            ended_at TIMESTAMPTZ,
            peak INTEGER NOT NULL DEFAULT 0,
            seconds BIGINT NOT NULL DEFAULT 0
        )""")
    cur.execute("""
        CREATE TABLE IF NOT EXISTS meeting_people (
            id SERIAL PRIMARY KEY,
            meeting_id INTEGER NOT NULL REFERENCES meetings(id) ON DELETE CASCADE,
            identity TEXT UNIQUE NOT NULL,
            user_id INTEGER REFERENCES users(id) ON DELETE SET NULL,
            name TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'attendee',
            status TEXT NOT NULL DEFAULT 'waiting',
            ticket_hash TEXT,
            seat_hash TEXT,
            asked_at TIMESTAMPTZ NOT NULL DEFAULT now(),
            admitted_at TIMESTAMPTZ,
            joined_at TIMESTAMPTZ,
            left_at TIMESTAMPTZ,
            in_since TIMESTAMPTZ,
            seconds INTEGER NOT NULL DEFAULT 0
        )""")
    cur.execute("CREATE INDEX IF NOT EXISTS meeting_people_meeting ON meeting_people (meeting_id)")
    cur.execute("CREATE INDEX IF NOT EXISTS meetings_host ON meetings (host_id)")
    cur.execute("CREATE TABLE IF NOT EXISTS meeting_usage (month TEXT PRIMARY KEY, seconds BIGINT NOT NULL DEFAULT 0)")


def _meet_hash(secret):
    return hashlib.sha256(("meet|" + str(secret or "")).encode()).hexdigest()


def _meet_month():
    return datetime.now(timezone.utc).strftime("%Y-%m")


def _meet_clean_name(v, fallback="Guest"):
    v = re.sub(r"[\x00-\x1f\x7f<>]", "", _text(v)).strip()
    v = re.sub(r"\s+", " ", v)[:40]
    return v or fallback


# ------------------------------------------------------------ LiveKit: passes and the room API
def _lk_jwt(extra, ttl):
    now = int(time.time())
    claims = {"iss": LIVEKIT_API_KEY, "nbf": now - 30, "exp": now + int(ttl)}
    claims.update(extra)
    tok = pyjwt.encode(claims, LIVEKIT_API_SECRET, algorithm="HS256")
    return tok.decode() if isinstance(tok, bytes) else tok


def _meet_can_talk(m, role):
    return role in ("host", "cohost", "speaker") or (role == "attendee" and m["speak"] == "all")


def _meet_grant(m, p):
    talk = _meet_can_talk(m, p["role"])
    # canUpdateOwnMetadata off: nobody can rename themselves (to look like the host, say) or change their own details;
    # a raised hand and what someone presents are set by the server (POST /me), which checks they may
    g = {"room": m["room"], "roomJoin": True, "canSubscribe": True, "canPublishData": True,
         "canUpdateOwnMetadata": False, "canPublish": talk}
    if talk:
        g["canPublishSources"] = ["camera", "microphone", "screen_share", "screen_share_audio"]
    return g


def _meet_token(m, p):
    return _lk_jwt({"sub": p["identity"], "name": p["name"], "video": _meet_grant(m, p),
                    "metadata": json.dumps({"role": p["role"]})}, MEET_TOKEN_TTL)


def _lk_call(method, body, room=None, timeout=8):
    """One call to LiveKit's room API. Not found answers as empty; anything else wrong is logged for the
    owner and met with plain words."""
    grant = {"roomCreate": True, "roomList": True, "roomAdmin": True}
    if room:
        grant["room"] = room
    tok = _lk_jwt({"video": grant}, 120)
    try:
        r = requests.post(f"{LIVEKIT_HTTP}/twirp/livekit.RoomService/{method}", json=body, timeout=timeout,
                          headers={"Authorization": "Bearer " + tok, "Content-Type": "application/json"})
    except requests.exceptions.RequestException as e:
        logger.warning(f"Meet: couldn't reach LiveKit ({method}): {e}")
        raise MeetError(MEET_MSG_DOWN, 503)
    if r.status_code == 404:
        return {}
    if r.status_code != 200:
        detail = (r.text or "")[:300]
        low = detail.lower()
        if r.status_code in (401, 403) and "permission" not in low:
            logger.error(f"Meet is OFF for everyone: LiveKit refused the key ({r.status_code}). Check LIVEKIT_API_KEY and "
                         f"LIVEKIT_API_SECRET on Render. ({method})")
        elif "limit" in low or "quota" in low or r.status_code == 429:
            logger.error(f"Meet: LiveKit says a limit was reached ({method} {r.status_code}): {detail[:200]}. Upgrade the "
                         "LiveKit plan, or lower MEET_GLOBAL_MAX / MEET_MONTH_MINUTES on Render to match it.")
            raise MeetError(MEET_MSG_BUSY, 503)
        else:
            logger.warning(f"Meet: LiveKit {method} -> {r.status_code}: {detail}")
        raise MeetError(MEET_MSG_DOWN, 503)
    try:
        return r.json() if r.content else {}
    except ValueError:
        return {}


_meet_rooms_cache = {"at": 0.0, "total": 0}


def _meet_people_now(room):
    """How many people are in one meeting right now (LiveKit's own count)."""
    return len(_lk_call("ListParticipants", {"room": room}, room=room).get("participants") or [])


def _meet_everyone_now():
    """How many people are in all meetings right now — asked at most every 5 seconds."""
    if time.time() - _meet_rooms_cache["at"] > 5:
        rooms = _lk_call("ListRooms", {}).get("rooms") or []
        _meet_rooms_cache.update(at=time.time(), total=sum(int(r.get("num_participants") or r.get("numParticipants") or 0) for r in rooms))
    return _meet_rooms_cache["total"]


def _meet_month_seconds(cur):
    cur.execute("SELECT seconds FROM meeting_usage WHERE month = %s", (_meet_month(),))
    row = cur.fetchone()
    return int(row["seconds"]) if row else 0


def _meet_check_traffic(cur, m, joining_role, here=None):
    """Turn people away kindly before LiveKit would: this month's minutes, everyone in meetings now, this meeting."""
    used = _meet_month_seconds(cur) // 60
    if used >= MEET_MONTH_MINUTES:
        logger.error(f"Meet: this month's {MEET_MONTH_MINUTES} participant-minutes are used up — new joins are turned away "
                     "until next month. Upgrade the LiveKit plan and raise MEET_MONTH_MINUTES on Render to allow more.")
        raise MeetError(MEET_MSG_MONTH, 503)
    if used >= MEET_MONTH_MINUTES * 0.8:
        logger.warning(f"Meet: {used} of this month's {MEET_MONTH_MINUTES} participant-minutes used.")
    if _meet_everyone_now() >= MEET_GLOBAL_MAX:
        logger.warning(f"Meet: {MEET_GLOBAL_MAX} people are in meetings at once (MEET_GLOBAL_MAX) — a newcomer was asked to try later.")
        raise MeetError(MEET_MSG_BUSY, 503)
    if joining_role != "host" and _meet_running(m):
        if here is None:
            here = _meet_people_now(m["room"])
        if here >= m["max_people"]:
            raise MeetError(MEET_MSG_FULL.format(n=m["max_people"]), 409)


def _meet_meta(cur, m):
    """The meeting's shared state, as only this server can set it (LiveKit room metadata)."""
    cur.execute("SELECT identity, role FROM meeting_people WHERE meeting_id = %s AND status = 'admitted' AND role <> 'attendee'", (m["id"],))
    roles = {r["identity"]: r["role"] for r in cur.fetchall()}
    return json.dumps({"title": m["title"], "locked": m["locked"], "lobby": m["lobby"], "speak": m["speak"], "chat": m["chat"],
                       "e2ee": m["e2ee"], "endsAt": m["ends_at"].isoformat() if m.get("ends_at") else None, "roles": roles,
                       "v": int(time.time() * 1000)})


def _meet_push(cur, m):
    if _meet_running(m):
        _lk_call("UpdateRoomMetadata", {"room": m["room"], "metadata": _meet_meta(cur, m)}, room=m["room"])


def _meet_open_room(cur, m):
    _lk_call("CreateRoom", {"name": m["room"], "empty_timeout": 600, "departure_timeout": 600,
                            "max_participants": m["max_people"] + 2, "metadata": _meet_meta(cur, m)}, room=m["room"])


# ------------------------------------------------------------ who is asking
def _meet_user():
    """The signed-in Academic-Home account behind this request, or None (guests may join meetings)."""
    header = request.headers.get("Authorization", "")
    token = header.split(" ", 1)[1] if header.startswith("Bearer ") else None
    uid = verify_token(token) if token else None
    if not uid:
        return None
    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM users WHERE id = %s", (uid,))
            row = cur.fetchone()
    finally:
        conn.close()
    if not row or _token_version(token) != (row.get("token_version") or 0):
        return None
    return _with_promo(row)


def _meet_get(cur, code, lock=False):
    code = _text(code).strip().lower()
    if not re.fullmatch(r"[a-z]{3}-[a-z]{4}-[a-z]{3}", code):
        raise MeetError(MEET_MSG_NOCODE, 404)
    cur.execute("SELECT * FROM meetings WHERE code = %s" + (" FOR UPDATE" if lock else ""), (code,))
    m = cur.fetchone()
    if not m:
        raise MeetError(MEET_MSG_NOCODE, 404)
    return m


def _meet_actor(cur, m, data):
    """Host controls: the host's own account, or a co-host's seat."""
    user = _meet_user()
    if user and user["id"] == m["host_id"]:
        cur.execute("SELECT * FROM meeting_people WHERE meeting_id = %s AND user_id = %s AND role = 'host' ORDER BY id DESC LIMIT 1",
                    (m["id"], user["id"]))
        return cur.fetchone() or {"identity": None, "role": "host", "user_id": user["id"]}
    seat = _text(data.get("seat") or request.headers.get("X-Meet-Seat"))
    if seat:
        cur.execute("SELECT * FROM meeting_people WHERE meeting_id = %s AND seat_hash = %s AND status = 'admitted'",
                    (m["id"], _meet_hash(seat)))
        p = cur.fetchone()
        if p and p["role"] in ("host", "cohost"):
            return p
    raise MeetError("Only the host can do that.", 403)


def _meet_admit(cur, m, p):
    """Let someone in: a seat (for coming back and co-host controls) and a pass to LiveKit."""
    seat = secrets.token_urlsafe(24)
    cur.execute("UPDATE meeting_people SET status = 'admitted', admitted_at = COALESCE(admitted_at, now()), seat_hash = %s, "
                "ticket_hash = NULL WHERE id = %s RETURNING *", (_meet_hash(seat), p["id"]))
    p = cur.fetchone()
    return {"status": "admitted", "seat": seat, "token": _meet_token(m, p), "url": LIVEKIT_URL, "identity": p["identity"],
            "role": p["role"], "name": p["name"], "e2ee": m["e2ee"], "title": m["title"], "code": m["code"]}


def _meet_running(m):
    return bool(m.get("started_at") and not m.get("ended_at"))


def _meet_public(m, host_name=None):
    state = "live" if _meet_running(m) else ("ended" if m["ended_at"] else "waiting")
    return {"code": m["code"], "title": m["title"], "host": host_name, "state": state, "passcode": bool(m["passcode_hash"]),
            "lobby": m["lobby"], "e2ee": m["e2ee"], "locked": m["locked"], "maxPeople": m["max_people"],
            "minutes": m["limit_minutes"], "endsAt": m["ends_at"].isoformat() if m["ends_at"] else None}


def _meet_reply(fn):
    """Meeting endpoints: switched off, plain-words errors, one database connection."""
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if request.method == "OPTIONS":
            return "", 204
        if not _meet_on():
            return jsonify({"error": MEET_MSG_OFF, "off": True}), 503
        conn = get_db()
        conn.autocommit = False
        try:
            with conn.cursor() as cur:
                out = fn(cur, *args, **kwargs)
            conn.commit()
            return out
        except MeetError as e:
            conn.rollback()
            return jsonify(dict({"error": str(e)}, **e.extra)), e.status
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()
    return wrapper


_meet_misses = {}


def _meet_miss(code_ok):
    """Guessing codes is pointless (23^10 of them), and one network trying many wrong ones is slowed down."""
    ip, now = _client_ip(), time.time()
    hits = [t for t in _meet_misses.get(ip, []) if now - t < 3600]
    if not code_ok:
        hits.append(now)
    _meet_misses[ip] = hits
    if len(_meet_misses) > 5000:
        _meet_misses.clear()
    if len(hits) > 40:
        raise MeetError("Too many wrong meeting codes from this network. Please wait a while, then try again.", 429)


def _meet_code_key():
    return f"{_client_ip()}|{(request.view_args or {}).get('code', '')}"


def _meet_person_key():
    """One person in one meeting: their waiting ticket or seat (both secret), else their network."""
    data = request.get_json(silent=True) or {}
    secret = _text(data.get("ticket") or data.get("seat") or request.headers.get("X-Meet-Seat") or request.headers.get("Authorization"))
    if secret:
        return "meet-person:" + hashlib.sha256(secret.encode()).hexdigest()[:24]
    return _meet_code_key()


# ------------------------------------------------------------ the endpoints
@app.route("/api/meet/config", methods=["GET", "OPTIONS"])
def meet_config():
    if request.method == "OPTIONS":
        return "", 204
    return jsonify({"enabled": _meet_on(), "maxPeople": MEET_MAX_PEOPLE, "freeMinutes": MEET_FREE_MINUTES,
                    "proMinutes": MEET_PRO_MINUTES, "audienceOver": MEET_AUDIENCE_OVER})


@app.route("/api/meet", methods=["POST", "OPTIONS"])
@limiter.limit("20 per hour", key_func=_account_or_ip_key)
@auth_required
@_meet_reply
def meet_create(cur):
    user = request.current_user
    data = _json_body()
    cur.execute("SELECT count(*) AS n FROM meetings WHERE host_id = %s AND started_at IS NOT NULL AND ended_at IS NULL", (user["id"],))
    if (cur.fetchone() or {}).get("n", 0) >= MEET_LIVE_PER_HOST:
        raise MeetError(f"You already have {MEET_LIVE_PER_HOST} meetings running. End one first.", 409)
    title = _meet_clean_name(data.get("title"), fallback=f"{(user.get('name') or 'Academic-Home').split()[0]}’s meeting")[:80]
    passcode = re.sub(r"\s+", "", _text(data.get("passcode")))[:20]
    speak = "chosen" if data.get("speak") == "chosen" else "all"
    minutes = MEET_PRO_MINUTES if _is_pro(user) else MEET_FREE_MINUTES
    for _ in range(5):
        code = "".join(secrets.choice(MEET_LETTERS) for _ in range(10))
        code = f"{code[:3]}-{code[3:7]}-{code[7:]}"
        cur.execute("SELECT 1 FROM meetings WHERE code = %s", (code,))
        if not cur.fetchone():
            break
    cur.execute("INSERT INTO meetings (code, room, host_id, title, passcode_hash, lobby, e2ee, speak, max_people, limit_minutes) "
                "VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s) RETURNING *",
                (code, "m_" + secrets.token_hex(12), user["id"], title,
                 generate_password_hash(passcode) if passcode else None, data.get("lobby", True) is not False,
                 bool(data.get("e2ee")), speak, MEET_MAX_PEOPLE, minutes))
    m = cur.fetchone()
    logger.info(f"Meet: meeting {m['id']} made (lobby={m['lobby']} e2ee={m['e2ee']} speak={speak} minutes={minutes})")
    return jsonify({"meeting": _meet_public(m, user.get("name")), "link": f"{FRONTEND_URL}/?meet={code}"})


@app.route("/api/meet/mine", methods=["GET", "OPTIONS"])
@auth_required
@_meet_reply
def meet_mine(cur):
    user = request.current_user
    cur.execute("SELECT m.*, (SELECT count(*) FROM meeting_people p WHERE p.meeting_id = m.id AND p.joined_at IS NOT NULL) AS came "
                "FROM meetings m WHERE host_id = %s ORDER BY created_at DESC LIMIT 30", (user["id"],))
    out = []
    for m in cur.fetchall():
        d = _meet_public(m, user.get("name"))
        d.update(created=m["created_at"].isoformat(), came=int(m["came"] or 0), peak=m["peak"])
        out.append(d)
    return jsonify({"meetings": out})


@app.route("/api/meet/<code>", methods=["GET", "OPTIONS"])
@limiter.limit("240 per hour", key_func=_meet_code_key)
@_meet_reply
def meet_info(cur, code):
    try:
        m = _meet_get(cur, code)
    except MeetError:
        _meet_miss(False)
        raise
    cur.execute("SELECT name FROM users WHERE id = %s", (m["host_id"],))
    host = ((cur.fetchone() or {}).get("name") or "").split(" ")[0] or None
    user = _meet_user()
    d = _meet_public(m, host)
    d["youHost"] = bool(user and user["id"] == m["host_id"])
    return jsonify({"meeting": d})


@app.route("/api/meet/<code>/join", methods=["POST", "OPTIONS"])
@limiter.limit("300 per hour", key_func=_meet_code_key)
@_meet_reply
def meet_join(cur, code):
    data = _json_body()
    try:
        m = _meet_get(cur, code, lock=True)
    except MeetError:
        _meet_miss(False)
        raise
    user = _meet_user()
    is_host = bool(user and user["id"] == m["host_id"])
    # coming back with a seat (a dropped connection, a reload): no waiting room again
    seat = _text(data.get("seat"))
    if seat:
        cur.execute("SELECT * FROM meeting_people WHERE meeting_id = %s AND seat_hash = %s", (m["id"], _meet_hash(seat)))
        p = cur.fetchone()
        if p and p["status"] == "removed":
            raise MeetError(MEET_MSG_REMOVED, 403)
        if p and p["status"] == "admitted" and _meet_running(m):
            _meet_check_traffic(cur, m, p["role"])
            return jsonify(_meet_admit(cur, m, p))
    if not is_host:
        if m["locked"]:
            raise MeetError(MEET_MSG_LOCKED, 423)
        if m["passcode_hash"] and not check_password_hash(m["passcode_hash"], re.sub(r"\s+", "", _text(data.get("passcode")))):
            raise MeetError(MEET_MSG_PASS, 403, passcode=True)
    if user and not is_host:
        cur.execute("SELECT 1 FROM meeting_people WHERE meeting_id = %s AND user_id = %s AND status = 'removed' "
                    "AND asked_at > COALESCE(%s, now())", (m["id"], user["id"], m["started_at"]))
        if cur.fetchone():
            raise MeetError(MEET_MSG_REMOVED, 403)
    name = _meet_clean_name(user.get("name") if user else data.get("name"))
    if user and data.get("name"):
        name = _meet_clean_name(data.get("name"), fallback=name)
    here = None
    if is_host:
        role = "host"
    else:
        role = "attendee"
        if _meet_running(m):
            here = _meet_people_now(m["room"])
            if m["speak"] == "all" and here >= MEET_AUDIENCE_OVER:
                role = "listener"     # a big meeting: newcomers listen (and chat, and raise a hand) until let speak
    _meet_check_traffic(cur, m, role, here)
    identity = ("m" if user else "g") + secrets.token_hex(8)    # m: an Academic-Home account, g: a guest
    cur.execute("INSERT INTO meeting_people (meeting_id, identity, user_id, name, role) VALUES (%s, %s, %s, %s, %s) RETURNING *",
                (m["id"], identity, user["id"] if user else None, name, role))
    p = cur.fetchone()
    if is_host:
        if not _meet_running(m):       # start it — or start it again: a meeting's link can be used for the next class too
            cur.execute("UPDATE meetings SET started_at = now(), ends_at = now() + make_interval(mins => limit_minutes), "
                        "ended_at = NULL WHERE id = %s RETURNING *", (m["id"],))
            m = cur.fetchone()
            logger.info(f"Meet: meeting {m['id']} started by its host")
        reply = _meet_admit(cur, m, p)
        _meet_open_room(cur, m)
        return jsonify(reply)
    if not _meet_running(m) or m["lobby"]:
        ticket = secrets.token_urlsafe(24)
        cur.execute("UPDATE meeting_people SET ticket_hash = %s WHERE id = %s", (_meet_hash(ticket), p["id"]))
        return jsonify({"status": "waiting", "ticket": ticket, "started": _meet_running(m), "title": m["title"]}), 202
    return jsonify(_meet_admit(cur, m, p))


@app.route("/api/meet/<code>/wait", methods=["POST", "OPTIONS"])
@limiter.limit("1500 per hour", key_func=_meet_person_key)
@_meet_reply
def meet_wait(cur, code):
    """Someone in the waiting room asks how it's going (every few seconds)."""
    data = _json_body()
    m = _meet_get(cur, code)
    cur.execute("SELECT * FROM meeting_people WHERE meeting_id = %s AND ticket_hash = %s", (m["id"], _meet_hash(_text(data.get("ticket")))))
    p = cur.fetchone()
    if not p:
        raise MeetError("You’re no longer waiting for this meeting. Please ask to join again.", 404)
    if p["status"] == "denied":
        raise MeetError(MEET_MSG_DENIED, 403)
    if datetime.now(timezone.utc) - p["asked_at"] > timedelta(minutes=MEET_WAIT_MINUTES):
        raise MeetError("You waited a long time, so we stopped. Please ask to join again.", 408)
    if p["status"] == "removed":
        raise MeetError(MEET_MSG_REMOVED, 403)
    if _meet_running(m) and (p["status"] == "admitted" or not m["lobby"]):
        _meet_check_traffic(cur, m, p["role"])
        return jsonify(_meet_admit(cur, m, p))
    return jsonify({"status": "waiting", "started": _meet_running(m)}), 202


@app.route("/api/meet/<code>/lobby", methods=["POST", "OPTIONS"])
@limiter.limit("1500 per hour", key_func=_meet_person_key)
@_meet_reply
def meet_lobby(cur, code):
    """The host's view of the waiting room."""
    data = _json_body()
    m = _meet_get(cur, code)
    _meet_actor(cur, m, data)
    cur.execute("SELECT identity, name, asked_at, user_id IS NOT NULL AS member FROM meeting_people WHERE meeting_id = %s "
                "AND status = 'waiting' AND ticket_hash IS NOT NULL AND asked_at > now() - make_interval(mins => %s) ORDER BY asked_at",
                (m["id"], MEET_WAIT_MINUTES))
    return jsonify({"waiting": [{"identity": r["identity"], "name": r["name"], "member": r["member"],
                                 "since": r["asked_at"].isoformat()} for r in cur.fetchall()]})


@app.route("/api/meet/<code>/me", methods=["POST", "OPTIONS"])
@limiter.limit("2400 per hour", key_func=_meet_person_key)
@_meet_reply
def meet_me(cur, code):
    """Your raised hand, and what you're presenting (and its page) \u2014 set by the server, so nobody can set anyone
    else's, or present when the host has made them a listener."""
    data = _json_body()
    m = _meet_get(cur, code)
    seat, p = _text(data.get("seat") or request.headers.get("X-Meet-Seat")), None
    if seat:
        cur.execute("SELECT * FROM meeting_people WHERE meeting_id = %s AND seat_hash = %s AND status = 'admitted'",
                    (m["id"], _meet_hash(seat)))
        p = cur.fetchone()
    if not p or not _meet_running(m):
        raise MeetError("You\u2019re not in this meeting.", 403)
    attrs = {}
    if "hand" in data:
        attrs["hand"] = str(int(time.time() * 1000)) if data.get("hand") else ""
    if "present" in data:
        v = data.get("present") or ""
        if v:
            if not _meet_can_talk(m, p["role"]):
                raise MeetError("The host has made you a listener \u2014 raise your hand to ask to present.", 403)
            try:
                s = json.loads(v) if isinstance(v, str) else v
                s = {"id": re.sub(r"[^a-z0-9]", "", str(s.get("id") or ""))[:16], "kind": "pdf",
                     "name": _meet_clean_name(s.get("name"), fallback="Slides")[:80],
                     "page": max(0, min(int(s.get("page") or 0), 5000)), "pages": max(1, min(int(s.get("pages") or 1), 5000))}
            except (TypeError, ValueError, AttributeError):
                raise MeetError("That presentation couldn\u2019t be shown.", 400)
            if not s["id"]:
                raise MeetError("That presentation couldn\u2019t be shown.", 400)
            v = json.dumps(s, separators=(",", ":"))
        attrs["present"] = v
    if attrs:
        _lk_call("UpdateParticipant", {"room": m["room"], "identity": p["identity"], "attributes": attrs}, room=m["room"])
    return jsonify({"ok": True})


def _meet_find_track(room, identity, kind):
    for part in _lk_call("ListParticipants", {"room": room}, room=room).get("participants") or []:
        if part.get("identity") == identity:
            for t in part.get("tracks") or []:
                src = str(t.get("source") or "").upper()
                if (kind == "audio" and src == "MICROPHONE") or (kind == "video" and src in ("CAMERA", "SCREEN_SHARE")):
                    yield t.get("sid")


def _meet_set_talk(m, identity, role):
    talk = _meet_can_talk(m, role)
    perm = {"can_subscribe": True, "can_publish_data": True, "can_update_metadata": False, "can_publish": talk}
    if talk:
        perm["can_publish_sources"] = ["CAMERA", "MICROPHONE", "SCREEN_SHARE", "SCREEN_SHARE_AUDIO"]
    body = {"room": m["room"], "identity": identity, "permission": perm, "metadata": json.dumps({"role": role})}
    if not talk:
        body["attributes"] = {"present": ""}      # someone who may no longer speak stops presenting too
    _lk_call("UpdateParticipant", body, room=m["room"])


@app.route("/api/meet/<code>/control", methods=["POST", "OPTIONS"])
@limiter.limit("900 per hour", key_func=_meet_person_key)
@_meet_reply
def meet_control(cur, code):
    """What the host (or a co-host) can do."""
    data = _json_body()
    m = _meet_get(cur, code, lock=True)
    actor = _meet_actor(cur, m, data)
    action, who = _text(data.get("action")), _text(data.get("identity"))
    if not _meet_running(m) and action not in ("end", "delete", "lock", "unlock", "lobby_on", "lobby_off", "chat_on", "chat_off",
                                               "speak_all", "speak_chosen", "admit", "deny", "admit_all"):
        raise MeetError("This meeting isn’t running.", 409)
    target = None
    if who:
        cur.execute("SELECT * FROM meeting_people WHERE meeting_id = %s AND identity = %s", (m["id"], who))
        target = cur.fetchone()
        if not target:
            raise MeetError("That person isn’t in this meeting.", 404)
        if target["role"] == "host" and (action != "mute" or actor["role"] != "host"):
            raise MeetError("The host can’t be changed.", 403)
        if target["role"] == "cohost" and actor["role"] != "host" and action in ("remove", "attendee", "speaker", "listener", "cohost"):
            raise MeetError("Only the host can change a co-host.", 403)
    if action in ("end", "delete", "cohost") and actor["role"] != "host":
        raise MeetError("Only the host can do that.", 403)
    flags = {"lock": ("locked", True), "unlock": ("locked", False), "lobby_on": ("lobby", True), "lobby_off": ("lobby", False),
             "chat_on": ("chat", True), "chat_off": ("chat", False)}
    if action in flags:
        col, val = flags[action]
        cur.execute(f"UPDATE meetings SET {col} = %s WHERE id = %s RETURNING *", (val, m["id"]))
        m = cur.fetchone()
        if action == "lobby_off":    # everyone waiting comes in
            cur.execute("UPDATE meeting_people SET status = 'admitted' WHERE meeting_id = %s AND status = 'waiting'", (m["id"],))
        _meet_push(cur, m)
    elif action in ("speak_all", "speak_chosen"):
        cur.execute("UPDATE meetings SET speak = %s WHERE id = %s RETURNING *", ("all" if action == "speak_all" else "chosen", m["id"]))
        m = cur.fetchone()
        cur.execute("SELECT identity, role FROM meeting_people WHERE meeting_id = %s AND status = 'admitted' AND role = 'attendee' "
                    "AND in_since IS NOT NULL", (m["id"],))
        for p in cur.fetchall()[:300]:
            try:
                _meet_set_talk(m, p["identity"], p["role"])
            except MeetError:
                pass
        _meet_push(cur, m)
    elif action in ("admit", "deny", "admit_all"):
        if action == "admit_all":
            cur.execute("UPDATE meeting_people SET status = 'admitted' WHERE meeting_id = %s AND status = 'waiting'", (m["id"],))
        elif target:
            cur.execute("UPDATE meeting_people SET status = %s WHERE id = %s AND status = 'waiting'",
                        ("admitted" if action == "admit" else "denied", target["id"]))
    elif action == "mute" and target:
        for sid in list(_meet_find_track(m["room"], target["identity"], _text(data.get("kind")) or "audio")):
            _lk_call("MutePublishedTrack", {"room": m["room"], "identity": target["identity"], "track_sid": sid, "muted": True}, room=m["room"])
    elif action == "remove" and target:
        cur.execute("UPDATE meeting_people SET status = 'removed' WHERE id = %s", (target["id"],))
        _lk_call("RemoveParticipant", {"room": m["room"], "identity": target["identity"]}, room=m["room"])
        if not target["user_id"] and not m["lobby"]:
            # a guest can't be recognised if they come back with the link — so the waiting room goes on
            cur.execute("UPDATE meetings SET lobby = TRUE WHERE id = %s RETURNING *", (m["id"],))
            m = cur.fetchone()
            _meet_push(cur, m)
    elif action in ("speaker", "attendee", "listener", "cohost") and target:
        cur.execute("UPDATE meeting_people SET role = %s WHERE id = %s", (action, target["id"]))
        _meet_set_talk(m, target["identity"], action)
        _meet_push(cur, m)
    elif action == "end":
        _meet_end(cur, m, "the host ended it")
    elif action == "delete":
        _meet_end(cur, m, "the host deleted it")
        cur.execute("DELETE FROM meetings WHERE id = %s", (m["id"],))
        return jsonify({"ok": True, "deleted": True})
    else:
        raise MeetError("Unknown meeting control.", 400)
    cur.execute("SELECT * FROM meetings WHERE id = %s", (m["id"],))
    return jsonify({"ok": True, "meeting": _meet_public(cur.fetchone())})


def _meet_end(cur, m, why):
    if _meet_running(m):
        cur.execute("UPDATE meetings SET ended_at = now() WHERE id = %s", (m["id"],))
        logger.info(f"Meet: meeting {m['id']} ended ({why}) — peak {m['peak']} people")
    _meet_settle(cur, m)
    try:
        _lk_call("DeleteRoom", {"room": m["room"]}, room=m["room"])
    except MeetError:
        pass


def _meet_settle(cur, m, identity=None):
    """Time spent in the meeting so far is added up: for each person who left (or everyone), for the meeting, and for
    this month's participant-minutes."""
    q = ("UPDATE meeting_people SET left_at = now(), seconds = seconds + GREATEST(0, EXTRACT(EPOCH FROM now() - in_since))::int, "
         "in_since = NULL WHERE meeting_id = %s AND in_since IS NOT NULL")
    args = [m["id"]]
    if identity:
        q += " AND identity = %s"
        args.append(identity)
    cur.execute(q, args)
    cur.execute("SELECT seconds FROM meetings WHERE id = %s FOR UPDATE", (m["id"],))
    before = int((cur.fetchone() or {}).get("seconds") or 0)
    cur.execute("SELECT COALESCE(SUM(seconds), 0) AS s FROM meeting_people WHERE meeting_id = %s", (m["id"],))
    total = int((cur.fetchone() or {}).get("s") or 0)
    if total > before:
        cur.execute("UPDATE meetings SET seconds = %s WHERE id = %s", (total, m["id"]))
        cur.execute("INSERT INTO meeting_usage (month, seconds) VALUES (%s, %s) ON CONFLICT (month) DO UPDATE SET "
                    "seconds = meeting_usage.seconds + EXCLUDED.seconds", (_meet_month(), total - before))


@app.route("/api/meet/<code>/attendance", methods=["POST", "OPTIONS"])
@_meet_reply
def meet_attendance(cur, code):
    """Who came, when, and for how long — for the host, as a spreadsheet file (CSV)."""
    import csv as _csv
    import io as _io
    data = _json_body()
    m = _meet_get(cur, code)
    _meet_actor(cur, m, data)
    cur.execute("SELECT name, role, joined_at, left_at, seconds, in_since, user_id IS NOT NULL AS member FROM meeting_people "
                "WHERE meeting_id = %s AND joined_at IS NOT NULL ORDER BY joined_at", (m["id"],))
    rows = cur.fetchall()
    buf = _io.StringIO()
    w = _csv.writer(buf)
    w.writerow(["Name", "Role", "Academic-Home account", "First joined (UTC)", "Last left (UTC)", "Minutes"])
    now = datetime.now(timezone.utc)
    for r in rows:
        secs = int(r["seconds"] or 0) + (int((now - r["in_since"]).total_seconds()) if r["in_since"] else 0)
        w.writerow([("'" + r["name"]) if r["name"][:1] in "=+-@\t\r" else r["name"], r["role"], "yes" if r["member"] else "guest",
                    r["joined_at"].strftime("%Y-%m-%d %H:%M"), r["left_at"].strftime("%Y-%m-%d %H:%M") if r["left_at"] and not r["in_since"] else "still in",
                    round(secs / 60, 1)])
    name = re.sub(r"[^\w \-]+", "", m["title"]).strip()[:50] or "Meeting"
    resp = app.response_class(buf.getvalue(), mimetype="text/csv")
    resp.headers["Content-Disposition"] = f'attachment; filename="{name} attendance.csv"'
    return resp


@app.route("/api/meet/<code>/notes", methods=["POST", "OPTIONS"])
@limiter.limit(AI_LIMIT, key_func=_account_or_ip_key)
@auth_required
@_meet_reply
def meet_notes(cur, code):
    """Remy's notes of a meeting, from its chat, questions and slide titles — each point says where it came from."""
    data = _json_body()
    m = _meet_get(cur, code)
    if request.current_user["id"] != m["host_id"]:
        raise MeetError("Only the host can ask Remy for the notes.", 403)
    lines = []
    for it in (data.get("chat") or [])[:400]:
        if isinstance(it, dict) and _text(it.get("text")).strip():
            lines.append(f"[{_text(it.get('time'))[:5]} {_meet_clean_name(it.get('name'))}] {_text(it.get('text'))[:600]}")
    slides = [f"Slide {i + 1}: {_text(s)[:200]}" for i, s in enumerate((data.get("slides") or [])[:200]) if _text(s).strip()]
    if not lines and not slides:
        raise MeetError("There’s nothing to make notes from yet — no chat or slides.", 400)
    material = "\n".join(lines)[:30000] + ("\n\nSLIDES SHOWN:\n" + "\n".join(slides) if slides else "")
    try:
        charged = _log_sage_message_and_consume(request.current_user)
    except ConversionError as e:
        raise MeetError(str(e), 429)
    system = ("You write clear meeting notes for an Academic-Home meeting called \"" + m["title"] + "\" from its chat and the titles of the "
              "slides that were shown. Use only that material — never add facts it doesn't contain. Every point ends with where it "
              "came from in square brackets: the person and time for chat — like [Ama 10:42] — or the slide, like [Slide 3]. Sections: "
              "## Summary (3–5 points), ## Decisions, ## Questions asked (with the answers given, if any), ## Next steps (who does "
              "what, by when, if said). Leave a section out when there's nothing for it. Plain, short sentences.")
    try:
        notes = call_claude(system, material, max_tokens=1800, feature="meet-notes")
    except ConversionError as e:
        _refund_sage_message(request.current_user, charged)
        raise MeetError(str(e) + AI_MSG_NOT_COUNTED, 502)
    return jsonify({"notes": (notes or "").strip()})


@app.route("/api/meet/webhook", methods=["POST"])
@limiter.exempt
def meet_webhook():
    """LiveKit tells us who joined and left (attendance, minutes this month). Only LiveKit can sign these."""
    if not _meet_on():
        return jsonify({"ok": False}), 503
    body = request.get_data(cache=True) or b""
    auth = request.headers.get("Authorization", "")
    token = auth.split(" ", 1)[1] if auth.startswith("Bearer ") else auth
    try:
        claims = pyjwt.decode(token, LIVEKIT_API_SECRET, algorithms=["HS256"], leeway=120, options={"verify_aud": False})
    except Exception:
        return jsonify({"error": "bad signature"}), 401
    digest = base64.b64encode(hashlib.sha256(body).digest()).decode()
    if claims.get("iss") != LIVEKIT_API_KEY or not hmac.compare_digest(str(claims.get("sha256") or ""), digest):
        return jsonify({"error": "bad signature"}), 401
    try:
        ev = json.loads(body.decode("utf-8") or "{}")
    except ValueError:
        return jsonify({"error": "bad body"}), 400
    kind = ev.get("event")
    room = (ev.get("room") or {}).get("name") or ""
    who = (ev.get("participant") or {}).get("identity") or ""
    if not room.startswith("m_"):
        return jsonify({"ok": True})
    conn = get_db()
    conn.autocommit = False
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM meetings WHERE room = %s FOR UPDATE", (room,))
            m = cur.fetchone()
            if not m:
                return jsonify({"ok": True})
            if kind == "participant_joined" and who:
                cur.execute("UPDATE meeting_people SET joined_at = COALESCE(joined_at, now()), in_since = COALESCE(in_since, now()), "
                            "left_at = NULL WHERE meeting_id = %s AND identity = %s", (m["id"], who))
                cur.execute("UPDATE meetings SET peak = GREATEST(peak, (SELECT count(*) FROM meeting_people WHERE meeting_id = %s "
                            "AND in_since IS NOT NULL)) WHERE id = %s", (m["id"], m["id"]))
            elif kind in ("participant_left", "participant_connection_aborted") and who:
                _meet_settle(cur, m, who)
            elif kind == "room_finished":
                _meet_settle(cur, m)
                if _meet_running(m):
                    cur.execute("UPDATE meetings SET ended_at = now() WHERE id = %s", (m["id"],))
                    logger.info(f"Meet: meeting {m['id']} closed (everyone left) — peak {m['peak']} people")
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
    return jsonify({"ok": True})


_meet_hook_warned = {"at": 0}


def _meet_webhook_check(cur, rooms):
    """Attendance, and the monthly minutes limit, come from LiveKit's webhooks. People in a meeting for a while with
    none of them counted means the webhooks aren't arriving: say so in the owner's log (at most once an hour)."""
    busy = [r.get("name") for r in rooms if int(r.get("num_participants") or r.get("numParticipants") or 0) > 0]
    if not busy or time.time() - _meet_hook_warned["at"] < 3600:
        return
    cur.execute("SELECT m.room FROM meetings m WHERE m.room = ANY(%s) AND m.ended_at IS NULL AND m.started_at < now() - interval '3 minutes' "
                "AND NOT EXISTS (SELECT 1 FROM meeting_people p WHERE p.meeting_id = m.id AND p.joined_at IS NOT NULL)", (busy,))
    if cur.fetchone():
        _meet_hook_warned["at"] = time.time()
        logger.error("Meet: LiveKit's webhooks aren't arriving, so attendance and the monthly minutes limit aren't counted. In LiveKit "
                     "Cloud open Settings \u2192 Webhooks, add https://<this server>/api/meet/webhook and sign it with the same API key "
                     "as LIVEKIT_API_KEY on Render.")


def _meet_reap_once():
    """Meetings past their time end; meetings LiveKit no longer has close; old ones are deleted."""
    if not _meet_on():
        return
    conn = get_db()
    conn.autocommit = False
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM meetings WHERE ended_at IS NULL AND ends_at < now() LIMIT 50")
            for m in cur.fetchall():
                _meet_end(cur, m, "its time was up")
            cur.execute("SELECT * FROM meetings WHERE ended_at IS NULL AND started_at IS NOT NULL AND started_at < now() - interval '15 minutes' LIMIT 200")
            open_ = cur.fetchall()
            if open_:
                rooms = _lk_call("ListRooms", {}).get("rooms") or []
                live = {r.get("name") for r in rooms}
                for m in open_:
                    if m["room"] not in live:
                        _meet_end(cur, m, "no one was left in it")
                _meet_webhook_check(cur, rooms)
            cur.execute("DELETE FROM meetings WHERE (ended_at IS NOT NULL OR started_at IS NULL) "
                        "AND COALESCE(started_at, created_at) < now() - make_interval(days => %s)", (MEET_KEEP_DAYS,))
        conn.commit()
    except MeetError:
        conn.rollback()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def _meet_reap_loop():
    time.sleep(30)
    while True:
        try:
            _meet_reap_once()
        except Exception as e:
            logger.warning(f"Meet: tidy-up pass failed: {e}")
        time.sleep(60)


if _meet_on() and "fake" not in os.environ.get("DATABASE_URL", "fake"):
    _threading.Thread(target=_meet_reap_loop, daemon=True, name="meet-tidy").start()


init_db()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    app.run(host="0.0.0.0", port=port)
