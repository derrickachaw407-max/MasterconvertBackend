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
import unicodedata
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
from flask import Flask, request, send_file, jsonify, Response, stream_with_context, g
from flask_compress import Compress
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from itsdangerous import URLSafeTimedSerializer, BadSignature, SignatureExpired
from werkzeug.security import generate_password_hash, check_password_hash

from converters import (
    convert, text_to_pptx, summary_slides_to_pptx, academic_essay_to_docx, extract_text, ConversionError,
    PPT_EXTS, presentation_kind, as_pptx,
    apply_pdf_operations, pdf_split, images_to_pdf, pdf_shrink, pdf_get_page_thumbnails, quiz_to_docx, preview_file,
    warm_up_libreoffice,
)
from converters import _soffice_convert   # the server's own LibreOffice runner
from converters import pdf_apply_plan, pdf_plan_needs_pro, pdf_to_images   # 3.1: the PDF editor
from converters import word_open, word_save   # 3.1: the Word editor
from converters import CONVERTERS as _OLD_PAIRS   # 3.15: the pairs Convert always had (kept as they were)
import convert315 as _any                          # 3.15: every other kind of file
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
# quotas (RamzAI's monthly allowance is exact; it's counted in the database).
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
    "https://acadhome.vercel.app",       # 3.15: the one address
    "https://getdocente.vercel.app",
    "https://docently.vercel.app",
    "https://masterconvert-tau.vercel.app",
}
# Additional origins can be added without a code change via this env var
# (comma-separated) — set on Render if a custom domain is added later.
ALLOWED_ORIGINS |= {o.strip() for o in os.environ.get("EXTRA_ALLOWED_ORIGINS", "").split(",") if o.strip()}
# The app's address, used in emails (password resets, meeting invitations). 3.15: acadhome.vercel.app is the one
# address; an old one still set on Render is read as the new one (the old addresses only move people there).
FRONTEND_HOME = "https://acadhome.vercel.app"
FRONTEND_URL = os.environ.get("FRONTEND_URL", FRONTEND_HOME).rstrip("/")
if re.fullmatch(r"https?://(getdocente|docently|docently-1|docente-1-dusky|masterconvert-tau)\.vercel\.app", FRONTEND_URL):
    FRONTEND_URL = FRONTEND_HOME


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
    #   Pay-per-Use — one full use of Pro: PAYPERUSE_MESSAGES RamzAI messages
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
# 4.0: Scholar is bought as a pass — Day, Week, Month, Semester or Year. The earlier Pro plans still count while they last.
# 4.7: Month, Half a year and Year are sold; Day, Week and Semester passes bought before still count while they last.
SCHOLAR_PASSES = ("day", "week", "month", "half", "semester", "year")
PRO_PLANS = SCHOLAR_PASSES + ("monthly", "halfyear", "yearly")
_PRO_SQL = "('day', 'week', 'month', 'half', 'semester', 'year', 'monthly', 'halfyear', 'yearly')"
# research writing as often as the points allow: every Scholar pass since 5.0 (4.7 to 4.9: Half a year and Year only)
RESEARCH_PLANS = PRO_PLANS        # 5.0 (the owner, 10 Nov 2026): every Scholar plan opens everything, Month too; only points limit it
# 5.0 (the owner, 10 Nov 2026, 08:24): "I want monthly points to be 100 with limits applied so that out of the 100 only 10
# point is for research (user can decide to spend them on other features as well), the rest shouldn't be restricted on the
# rest of the ai features. Use the same idea on the half year … except yearly which will have no limits."
# A pass's research points: how many of its points may go on research (Research a question, automatic research writing,
# manual research writing). They are part of the pass's points, not a second purse: whatever isn't spent on research
# works on every other AI feature. A pass that isn't listed has no research limit (Year; and a pass bought before 5.0
# keeps what it was bought with).
PASS_RESEARCH_POINTS = {"month": 10, "half": 60}
SCHOLAR_MORE_RESEARCH = "Scholar gives you more research."
PLAN_NAMES = {"day": "Scholar Day", "week": "Scholar Week", "month": "Scholar Month", "half": "Scholar Half Year", "semester": "Scholar Semester", "year": "Scholar Year",
              "monthly": "Scholar Month", "halfyear": "Scholar Semester", "yearly": "Scholar Year", "free": "Free"}
# Points: the unit of AI work. A RamzAI answer, a summary or a quiz is 1; a whole deck is 3; ten cited statements are 2.
POINTS_REMY, POINTS_DECK, POINTS_WRITE, POINTS_RESEARCH = 1, 3, 2, 2
# 4.6: Research a question (RamzAI's Research) is 1 point (the owner, 9 Nov 2026: "1 point each time"); on Free it
# still counts as one of the month's research writings
POINTS_ASK_RESEARCH = 1
# 4.5 (the owner, 9 Nov 2026): Free gets 5 points every calendar month (UTC), once per phone, and at most 2 of its uses
# a month may be research writing — an automatic draft (/api/write) or a batch of statements for manual research
# writing (/api/guide/statements). Research writing still costs its usual 2 points each.
FREE_MONTH_POINTS = 5
FREE_MONTH_RESEARCH = 2
TRIAL_POINTS = FREE_MONTH_POINTS   # the old name; older apps read "trial_points" from /api/paystack/config
SCHOLAR_DAY_CAP = 60       # the most points any pass spends in one day (catches runaway or automated use)
BONUS_MONTH_POINTS = 100   # a bonus month given before 5.0 works like Scholar Month (5.0: 100 points, 10 of them for research)
# 5.0 (the owner, 10 Nov 2026, 08:44): "Inviting a friend fetches you 10 points including two research points."
# The person whose code a friend joined with gets INVITE_POINTS points that never end; INVITE_RESEARCH of them may go on
# research (users.bonus_research, used after a pass's own research points or Free's research writings of the month).
# A friend counts once: when they have confirmed their email on a phone that has had no account before (the rule Free's
# points already follow), or the first time they buy a pass with the code. At most INVITE_MONTH_MAX friends count in a
# calendar month (the owner, 10 Nov 10:05: "make 2 friends count a month"): every point is real AI work that is paid for.
# "What if the person signed up already, will they have to sign up again? Watch that well." — no: an account that
# exists can enter a friend's code once (/api/referral/use); and signing up again on the same phone earns nothing.
# "Protect users' invite code": a code ends in four random characters (1.6 million endings), checking and using codes
# is limited per hour, a right code shows only its owner's first name, nobody can use their own, and it counts once.
INVITE_POINTS = 10
INVITE_RESEARCH = 2
INVITE_MONTH_MAX = 2
# What Pro adds beyond Free (none of Free's own allowances change).
REMY_ATTACH_FREE, REMY_ATTACH_PRO = 20000, 60000     # characters of an attached file RamzAI reads (~8 / ~25 pages)
REMY_PAPERS_FREE, REMY_PAPERS_PRO = 6, 12            # research papers behind each answer
DECK_CONTENT_FREE, DECK_CONTENT_PRO = 14, 22         # content slides in a deck RamzAI drafts
PAYPERUSE_MESSAGES = 30   # one Pay-per-Use purchase: RamzAI messages that never expire

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
            # Pay-per-Use: RamzAI messages bought once, used after the monthly
            # allowance, never expiring — and every store transaction that
            # granted them, so the same purchase is never credited twice.
            cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS remy_credits INTEGER NOT NULL DEFAULT 0")
            # 4.0: a pass's points (end with the pass), the day's spending, the free trial (once per phone)
            cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS pass_points INTEGER NOT NULL DEFAULT 0")
            cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS day_points_used INTEGER NOT NULL DEFAULT 0")
            cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS day_points_at DATE")
            cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS trial_points_used INTEGER NOT NULL DEFAULT 0")
            # 4.5: trial_points_used counts Free's points in the month named by free_month (its first day, UTC);
            # free_research_used counts that month's research writings
            cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS free_month DATE")
            cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS free_research_used INTEGER NOT NULL DEFAULT 0")
            cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS trial_blocked BOOLEAN NOT NULL DEFAULT FALSE")
            # 5.0: how many of the running pass's points may still go on research; NULL = no research limit
            cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS research_points INTEGER")
            # 5.0: research points that never end (from inviting friends), and each friend who counted
            cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS bonus_research INTEGER NOT NULL DEFAULT 0")
            cur.execute("""
                CREATE TABLE IF NOT EXISTS invite_rewards (
                    friend_id INTEGER PRIMARY KEY,
                    inviter_id INTEGER NOT NULL,
                    points INTEGER NOT NULL,
                    research INTEGER NOT NULL,
                    created_at TIMESTAMPTZ NOT NULL DEFAULT now()
                )
            """)
            cur.execute("CREATE INDEX IF NOT EXISTS invite_rewards_inviter ON invite_rewards (inviter_id, created_at)")
            cur.execute("""
                CREATE TABLE IF NOT EXISTS trial_devices (
                    device_id TEXT PRIMARY KEY,
                    user_id INTEGER REFERENCES users(id) ON DELETE SET NULL,
                    created_at TIMESTAMPTZ NOT NULL DEFAULT now()
                )
            """)
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
            # 4.7: the points a payment was offered, written when it starts (prices can change while someone pays)
            cur.execute("ALTER TABLE paystack_payments ADD COLUMN IF NOT EXISTS points INTEGER")
            # 5.0: the research points offered when the payment started (NULL: no research limit, or started before 5.0)
            cur.execute("ALTER TABLE paystack_payments ADD COLUMN IF NOT EXISTS research_points INTEGER")
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
            # 4.9: a request sent again (a slow or dropped connection) is answered once — see _once below
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS request_once (
                    user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
                    key TEXT NOT NULL,
                    route TEXT NOT NULL,
                    state TEXT NOT NULL,
                    status INTEGER,
                    body BYTEA,
                    mimetype TEXT,
                    created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
                    PRIMARY KEY (user_id, key)
                )
                """
            )
            cur.execute("CREATE INDEX IF NOT EXISTS request_once_created ON request_once (created_at)")
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
        return False
    try:
        conn = get_db()
        try:
            with conn.cursor() as cur:
                # No duplicates, and at most MEMORY_LIMIT notes per person.
                cur.execute("SELECT note FROM user_memory WHERE user_id = %s", (user_id,))
                existing = [r["note"].strip().lower() for r in cur.fetchall()]
                if note.lower() in existing or len(existing) >= MEMORY_LIMIT:
                    return False
                cur.execute(
                    "INSERT INTO user_memory (user_id, note) VALUES (%s, %s)",
                    (user_id, note),
                )
                return True             # 4.5: the app says "RamzAI will remember…" only for a note really saved
        finally:
            conn.close()
    except Exception:
        return False  # memory is a nice-to-have — never let it break the actual response


MEMORY_LIMIT = 50


def _memory_on(user_row):
    """Whether this person lets RamzAI remember (Profile -> RamzAI's memory)."""
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


def make_reset_token(user_id, email=None, token_version=None):
    # 4.5: the account's token version goes in, so a link works once: the reset itself moves the version on
    data = {"reset_user_id": user_id, "email": (email or "").lower()}
    if token_version is not None:
        data["tv"] = int(token_version)
    return _serializer.dumps(data, salt="password-reset")


def _reset_token_version(token):
    try:
        v = (_serializer.loads(token, max_age=RESET_TOKEN_MAX_AGE, salt="password-reset") or {}).get("tv")
        return int(v) if v is not None else None
    except (BadSignature, SignatureExpired, TypeError, ValueError):
        return None


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
        logger.warning(f"Couldn't send a confirmation code to user {uid}: {e.__class__.__name__}")
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


def _find_referrer(cur, code, user_id=None, any_age=False):
    """The account a referral code belongs to: someone's personal code, or a
    14-day code they made. Never the person's own code. (any_age: a 14-day code that was good when the friend signed
    up still counts when the friend confirms their email a little later.)"""
    code = _clean_code(code)
    if not code:
        return None
    cur.execute("SELECT id, name FROM users WHERE referral_code = %s", (code,))
    ref = cur.fetchone()
    if not ref:
        cur.execute("SELECT u.id, u.name FROM promo_codes p JOIN users u ON u.id = p.created_by "
                    "WHERE p.code = %s" + ("" if any_age else " AND p.expires_at > now()"), (code,))
        ref = cur.fetchone()
    if ref and user_id and ref["id"] == user_id:
        return None
    return ref


def _reward_referral(cur, user_id, code=None):
    """A friend's code (5.0, the owner on 10 Nov 2026: "Inviting a friend fetches you 10 points including two research
    points"): the person who invited gets INVITE_POINTS points that never end, INVITE_RESEARCH of them usable on research.
    Once per friend account (user_id is the friend), with the code typed at checkout or the one saved at sign-up.
    Returns True when it was given. (4.9 and before: both got a bonus month; one already banked still starts.)"""
    cur.execute("SELECT id, referred_code, referral_rewarded FROM users WHERE id = %s FOR UPDATE", (user_id,))
    me = cur.fetchone()
    if not me or me["referral_rewarded"]:
        return False
    code = _clean_code(code) or _clean_code(me.get("referred_code"))
    ref = _find_referrer(cur, code, user_id, any_age=True) if code else None
    if not ref:
        return False
    # a month's worth of friends at most; one more friend waits for next month (nothing is marked, so it still counts then)
    cur.execute("SELECT count(*) AS n FROM invite_rewards WHERE inviter_id = %s AND created_at >= date_trunc('month', now())", (ref["id"],))
    if int((cur.fetchone() or {}).get("n") or 0) >= INVITE_MONTH_MAX:
        return False
    cur.execute("INSERT INTO invite_rewards (friend_id, inviter_id, points, research) VALUES (%s, %s, %s, %s) ON CONFLICT (friend_id) DO NOTHING RETURNING friend_id",
                (user_id, ref["id"], INVITE_POINTS, INVITE_RESEARCH))
    if not cur.fetchone():
        return False
    cur.execute("UPDATE users SET remy_credits = COALESCE(remy_credits, 0) + %s, bonus_research = bonus_research + %s WHERE id = %s",
                (INVITE_POINTS, INVITE_RESEARCH, ref["id"]))
    cur.execute("UPDATE users SET referral_rewarded = TRUE, referred_code = %s WHERE id = %s", (code, user_id))
    logger.info(f"Invite: user {ref['id']} got {INVITE_POINTS} points for user {user_id}")
    return True


def _invite_joined(row, device_id):
    """The friend has joined: they signed up with a code, their email is confirmed, and theirs is the first account on
    this phone (the rule Free's points follow — so many addresses on one phone earn nothing). Then the person who
    invited them gets the points. Looked at whenever the app opens, so it arrives as soon as the email is confirmed."""
    try:
        if (not row.get("referred_code") or row.get("referral_rewarded") or row.get("trial_blocked") or _needs_verification(row)
                or not device_id or not _DEVICE_RE.match(device_id) or not DATABASE_URL):
            return row
        conn = get_db()
        conn.autocommit = False
        try:
            with conn.cursor() as cur:
                cur.execute("SELECT user_id FROM trial_devices WHERE device_id = %s", (device_id,))
                seen = cur.fetchone()
                given = bool(seen and seen["user_id"] == row["id"] and _reward_referral(cur, row["id"]))
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()
        if given:
            row = dict(row)
            row["referral_rewarded"] = True
    except Exception as e:
        logger.warning(f"Couldn't look at the invite of user {row.get('id')}: {e}")
    return row


def _pass_expired(row):
    exp = row.get("plan_expires_at") if hasattr(row, "get") else None
    return bool(exp and row.get("plan") != "free" and exp <= datetime.now(timezone.utc))


def effective_plan(row):
    """The plan an account really has now: a Pay-per-Use pass whose 30 days
    have run out counts as Free, even before the database is corrected."""
    return "free" if _pass_expired(row) else row["plan"]


def _points_state(row):
    """What the account has: (points in all, pass points, bought points, trial points left, points still allowed
    today (never more than it has), the day's cap or None). Computed, never written."""
    plan = effective_plan(row)
    bought = int(row.get("remy_credits") or 0) if hasattr(row, "get") else 0
    if plan != "free":
        if (row.get("plan_source") in ("store", "promo")) and not (row.get("pass_points") or 0):
            used, limit, _ = _sage_status(row)
            left = max(0, limit - used)
            return left + bought, left, bought, 0, None, None
        pp = max(0, int(row.get("pass_points") or 0))
        today = datetime.now(timezone.utc).date()
        used_today = int(row.get("day_points_used") or 0) if row.get("day_points_at") == today else 0
        today_left = max(0, SCHOLAR_DAY_CAP - used_today)
        return pp + bought, pp, bought, 0, min(today_left, pp + bought), SCHOLAR_DAY_CAP
    trial_left = _free_state(row)[0]
    return trial_left + bought, 0, bought, trial_left, None, None


def _month_start(now=None):
    d = (now or datetime.now(timezone.utc)).date()
    return d.replace(day=1)


def _next_month_start(now=None):
    m = _month_start(now)
    return (m.replace(year=m.year + 1, month=1) if m.month == 12 else m.replace(month=m.month + 1))


def _free_state(row):
    """Free's month, computed, never written: (points left, research writings left, the day they come back).
    A new month starts at zero even before the database row is reset by the next use. A phone whose free points
    belong to another account (trial_blocked) has no free points, but may still use bought points for research
    writing within the month's research limit."""
    get = row.get if hasattr(row, "get") else (lambda k, d=None: d)
    month = _month_start()
    same = get("free_month") == month
    used = int(get("trial_points_used") or 0) if same else 0
    research_used = int(get("free_research_used") or 0) if same else 0
    points_left = 0 if get("trial_blocked") else max(0, FREE_MONTH_POINTS - used)
    return points_left, max(0, FREE_MONTH_RESEARCH - research_used), _next_month_start()


def _research_full(row):
    """Research is open (no monthly count of research writings): every Scholar pass (5.0, the owner on 10 Nov 2026).
    Free: FREE_MONTH_RESEARCH a month. How MANY points a pass may spend on research is _research_left."""
    try:
        return effective_plan(_with_promo(row)) in RESEARCH_PLANS
    except Exception:
        return False


def _bonus_research(row):
    """Research points that never end (from inviting friends): used after a pass's own research points, or after
    Free's research writings of the month."""
    try:
        return max(0, int(row.get("bonus_research") or 0))
    except Exception:
        return 0


def _research_left(row, own=False):
    """How many points may still go on research on the running pass: a number on Month and Half year (5.0) — the
    pass's own research points plus those that never end (own=True: the pass's own only) — and None when there is no
    research limit (Year, a pass bought before 5.0, the promotion) and on Free (which has its own rule:
    FREE_MONTH_RESEARCH research writings a month, then the research points that never end)."""
    try:
        if effective_plan(row) == "free" or row.get("plan_source") in ("store", "promo"):
            return None
        v = row.get("research_points")
        return None if v is None else max(0, int(v)) + (0 if own else _bonus_research(row))
    except Exception:
        return None


def _research_open(row):
    """Research may be asked for: a pass, a research writing left this month on Free, or research points that never end."""
    return _research_full(row) or _free_state(row)[1] > 0 or _bonus_research(row) > 0


class _Charged(str):
    """Where a spend came from ('pass', 'credit', 'trial', 'allowance'), with how many research points that never end
    it took (bonus), so that giving it back returns each to its place."""
    bonus = 0


def user_row_to_dict(row):
    row = _with_promo(row)
    exp = row.get("plan_expires_at") if hasattr(row, "get") else None
    msgs_used, msgs_limit, msgs_reset = _sage_status(row)
    points, pass_pts, bought_pts, trial_left, today_left, day_cap = _points_state(row)
    return {
        # 4.0: points — the unit of AI work — and the plan's name as the app shows it
        "points": points,
        "pass_points": pass_pts,
        "bought_points": bought_pts,
        "trial_points_left": trial_left,
        "trial_blocked": bool(row.get("trial_blocked")) if hasattr(row, "get") else False,
        # 4.5: Free's month — 5 points, at most 2 research writings, back on the first day of next month
        "free_points_month": FREE_MONTH_POINTS,
        "free_research_left": _free_state(row)[1],
        "free_research_month": FREE_MONTH_RESEARCH,
        # 4.7: research writing as often as the points allow (Half a year, Year); otherwise free_research_left a month
        "research_full": _research_full(row),
        # 5.0: on Month and Half year, how many of the pass's points may still go on research (None: no research limit)
        "research_points_left": _research_left(row),
        # 5.0: research points that never end (from inviting friends), and what an invite gives
        "bonus_research_left": _bonus_research(row),
        "invite": {"points": INVITE_POINTS, "research": INVITE_RESEARCH, "month_max": INVITE_MONTH_MAX},
        # whether a friend's code has already counted for this account (it counts once; until then one can be entered)
        "friend_code_counted": bool(row.get("referral_rewarded")) if hasattr(row, "get") else False,
        "free_resets_on": _free_state(row)[2].isoformat(),
        "points_today_left": today_left,
        "points_day_cap": day_cap,
        "plan_name": PLAN_NAMES.get(effective_plan(row), "Scholar"),
        "scholar": effective_plan(row) in PRO_PLANS,
        "plans": {k: {"amount_ghs": v[0] / 100, "days": v[1], "points": v[2], "research_points": PASS_RESEARCH_POINTS.get(k),
                      "name": PLAN_NAMES.get(k, "Points")} for k, v in PAYSTACK_PLANS.items() if k in PAYSTACK_ON_SALE},
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
        # The RamzAI message meter: this window's allowance, used so far, and
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
                    cur.execute("UPDATE users SET plan = 'free', plan_expires_at = NULL, pass_points = 0, research_points = NULL WHERE id = %s AND plan_expires_at <= now()", (user_id,))
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
                    cur.execute("UPDATE users SET plan = 'month', plan_expires_at = now() + interval '30 days', plan_source = 'bonus', pass_points = pass_points + %s, "
                                "research_points = %s, "
                                "bonus_credit_months = bonus_credit_months - 1 WHERE id = %s AND plan = 'free' AND bonus_credit_months > 0 RETURNING *",
                                (BONUS_MONTH_POINTS, PASS_RESEARCH_POINTS["month"], user_id))
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


# ---------- 4.9 (10 Nov 2026): a request sent again is answered once ----------
# On a slow or dropped connection the same request can reach the server twice: the app sends it again after an answer
# that never arrived (Retry, or its own wait), and the first one may have done the work after all. The app adds
# ?once=<key> to the requests that take points or make rows; the first to arrive does the work and its answer is kept,
# and the same key again gets that same answer: no second charge, no second meeting, invitation, payment or history row.
# A request still being worked on answers 409 {"working": true}; GET /api/once/<key> tells the app when it's ready.
# Answers are kept less than an hour (the privacy notice says up to an hour), only the account's own, and go with the account.
ONCE_KEEP = "50 minutes"           # deleted by the tidy-up every 5 minutes: never kept an hour
ONCE_MAX_BODY = 4 * 1024 * 1024
_ONCE_KEY = re.compile(r"^[A-Za-z0-9_-]{16,80}$")
_ONCE_STUCK_S = 600          # "working" longer than this (the longest request takes about 5 minutes): it died with its worker, the next one does the work


def _once_key():
    k = (request.args.get("once") or request.headers.get("Idempotency-Key") or "").strip()
    return k if _ONCE_KEY.match(k) else ""


def _once_replay(row):
    r = Response(bytes(row["body"]) if row["body"] is not None else b"", status=row["status"] or 200,
                 mimetype=row["mimetype"] or "application/json")
    r.headers["X-Once"] = "replay"
    return r


def _once_working():
    r = jsonify({"error": "Still working on this. It will be ready in a moment.", "working": True, "retry_in": 3})
    r.status_code = 409
    return r


def _once_age(row):
    return (datetime.now(timezone.utc) - row["created_at"]).total_seconds()


def _once_settle(uid, key, resp):
    """Keeps a good answer to send again; anything else is forgotten, so the same request runs again (whatever it
    charged was given back with the error)."""
    try:
        conn = get_db()
        try:
            with conn.cursor() as cur:
                good = (resp is not None and 200 <= resp.status_code < 300 and not resp.is_streamed and not resp.direct_passthrough
                        and not resp.headers.get("Content-Encoding"))
                data = resp.get_data() if good else None
                if good and len(data) <= ONCE_MAX_BODY:
                    cur.execute("UPDATE request_once SET state = 'done', status = %s, body = %s, mimetype = %s WHERE user_id = %s AND key = %s",
                                (resp.status_code, psycopg2.Binary(data), resp.mimetype, uid, key))
                else:
                    cur.execute("DELETE FROM request_once WHERE user_id = %s AND key = %s", (uid, key))
        finally:
            conn.close()
    except Exception as e:
        logger.warning(f"once: couldn't keep an answer: {e}")


@app.after_request
def _once_keep(resp):
    """The answer is kept as it was sent. This hook is registered before the others, so it runs after them (the AI's
    words already finished, see _finish_ai_answers) and before the gzip step (Compress, registered first of all)."""
    pending = g.pop("_once_pending", None)
    if pending:
        _once_settle(pending[0], pending[1], resp)
    return resp


@app.teardown_request
def _once_teardown(_exc):
    """A request that ended without an answer (an error while answering) is forgotten, so it can run again."""
    pending = g.pop("_once_pending", None)
    if pending:
        _once_settle(pending[0], pending[1], None)


def _once(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        key = _once_key() if request.method == "POST" else ""
        user = getattr(request, "current_user", None)
        if not key or not user or not DATABASE_URL:
            return fn(*args, **kwargs)
        uid, route = user["id"], request.path
        try:
            conn = get_db()
            try:
                with conn.cursor() as cur:
                    if random.random() < 0.05:
                        cur.execute(f"DELETE FROM request_once WHERE created_at < now() - interval '{ONCE_KEEP}'")
                    cur.execute("INSERT INTO request_once (user_id, key, route, state) VALUES (%s, %s, %s, 'working') "
                                "ON CONFLICT (user_id, key) DO NOTHING RETURNING key", (uid, key, route))
                    if cur.fetchone() is None:
                        cur.execute("SELECT route, state, status, body, mimetype, created_at FROM request_once WHERE user_id = %s AND key = %s", (uid, key))
                        row = cur.fetchone()
                        if row and row["route"] != route:
                            return jsonify({"error": "Something went wrong with this request. Please try again."}), 409
                        if row and row["state"] == "done":
                            logger.info(f"once: sent a kept answer again ({route})")
                            return _once_replay(row)
                        if row and _once_age(row) < _ONCE_STUCK_S:
                            return _once_working()
                        cur.execute("INSERT INTO request_once (user_id, key, route, state) VALUES (%s, %s, %s, 'working') "
                                    "ON CONFLICT (user_id, key) DO UPDATE SET state = 'working', status = NULL, body = NULL, created_at = now()",
                                    (uid, key, route))
            finally:
                conn.close()
        except Exception as e:
            logger.warning(f"once: {e}")              # the table isn't there yet: work as before
            return fn(*args, **kwargs)
        g._once_pending = (uid, key)       # kept (or forgotten) by _once_keep when the answer leaves
        try:
            return fn(*args, **kwargs)
        except Exception:
            g.pop("_once_pending", None)
            _once_settle(uid, key, None)
            raise

    return wrapper


def _once_logged(cur, user_id):
    """The conversion history: a conversion sent again with the same key is recorded once. A batch request records
    each of its files under its own sub-key."""
    key = _once_key()
    if not key:
        return False
    n = g.get("_once_n", 0)
    g._once_n = n + 1
    cur.execute("INSERT INTO request_once (user_id, key, route, state, status) VALUES (%s, %s, %s, 'logged', 200) "
                "ON CONFLICT (user_id, key) DO NOTHING RETURNING key", (user_id, f"{key}.{n}", request.path))
    return cur.fetchone() is None


# What an account that hasn't confirmed its email yet may still do.
_UNCONFIRMED_OK = {"me", "verify_email", "resend_code", "change_email", "delete_account_endpoint", "update_profile"}


# 3.2: what each AI call used, logged as counts with an estimated cost (USD per million tokens; web search per
# search) so the real cost of a RamzAI answer can be checked — never anyone's words or account.
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
AI_MSG_UNAVAILABLE = "RamzAI isn\u2019t available right now. Please try again a little later."
AI_MSG_BUSY = "RamzAI is busy right now. Please try again in a minute."
AI_MSG_SLOW = "RamzAI took too long to answer. Please try again."
AI_MSG_TOO_BIG = "That\u2019s more than RamzAI can read at once. Please try a shorter message or file."
AI_MSG_FAILED = "RamzAI couldn\u2019t answer just now. Please try again."
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
    style = _house_style()                       # 3.15: how the words must read
    if style and style not in (system_prompt or ""):
        system_prompt = (system_prompt or "") + style
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
        logger.warning("search_web_sources: request failed (a %d-character topic): %s", len(topic or ""), e.__class__.__name__)
        return [], f"web search request failed: {e}"
    if resp.status_code != 200:
        logger.warning(
            "search_web_sources: non-200 (%s) for a %d-character topic — body: %s",
            resp.status_code, len(topic or ""), resp.text[:300],
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
    logger.info("search_web_sources: a %d-character topic -> %d source(s)", len(topic or ""), len(sources))
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
            logger.warning("search_semantic_scholar: request failed (a %d-character query): %s", len(query or ""), e.__class__.__name__)
            return [], f"Semantic Scholar request failed: {e}"
        if resp.status_code == 429 and attempt == 0:
            logger.warning("search_semantic_scholar: 429 on first attempt, retrying after backoff")
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
        logger.warning("search_semantic_scholar: non-JSON response")
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
    logger.info("search_semantic_scholar: a %d-character query -> %d source(s) (raw hits: %d)", len(query or ""), len(sources), raw_count)
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
        logger.warning("search_for_sources: ZERO total sources found (a %d-character topic) — %s", len(topic or ""), combined_diag)
        return [], combined_diag
    logger.info("search_for_sources: a %d-character topic -> %d total source(s)", len(topic or ""), len(sources))
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
_PAPER_UA = {"User-Agent": "Academic-Home/3.4 (study app; https://acadhome.vercel.app)"}
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


# A retraction notice, a withdrawn paper or an expression of concern is never offered as evidence
_RETRACTED_TITLE = re.compile(r"^\W*(?:(?:retraction\s+notice|notice\s+of\s+(?:retraction|withdrawal)|withdrawal\s+notice|"
                              r"expression\s+of\s+concern)\b|(?:retracted|retraction|withdrawn)(?:\s+article)?\s*[:.\]\-–—])", re.I)
# Preprint servers, by name or DOI: a preprint isn't peer reviewed, and its reference says so
_PREPRINT_VENUE = re.compile(r"\b(?:arxiv|biorxiv|medrxiv|chemrxiv|psyarxiv|socarxiv|edarxiv|eartharxiv|techrxiv|ssrn|research\s*square|"
                             r"preprints?(?:\.org)?|osf\s+preprints|authorea)\b", re.I)
_PREPRINT_DOI = re.compile(r"^10\.(?:48550/|1101/\d{4}\.\d{2}\.\d{2}\.|2139/ssrn|21203/rs\.|20944/preprints|26434/chemrxiv|"
                           r"31234/osf\.io|31235/osf\.io|35542/osf\.io|31219/osf\.io|22541/au\.|36227/techrxiv|31223/)", re.I)


def _is_preprint(venue, doi):
    return bool((venue and _PREPRINT_VENUE.search(str(venue))) or (doi and _PREPRINT_DOI.match(str(doi))))


def _paper(title, authors, year, venue, doi, pdf, landing, abstract, found_in, cites=0, biblio=None):
    title = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", str(title or ""))).strip().rstrip(".")
    if not title or not pdf or not str(pdf).startswith("http") or _RETRACTED_TITLE.match(title):
        return None
    try:
        year = int(year) if year else None
    except (TypeError, ValueError):
        year = None
    authors = [re.sub(r"\s+", " ", str(a)).strip() for a in (authors or []) if str(a or "").strip()][:100]
    abstract = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", str(abstract or ""))).strip() or None
    doi = _norm_doi(doi)
    return {"title": title[:400], "authors": authors or None, "year": year, "venue": (str(venue).strip()[:200] or None) if venue else None,
            "doi": doi, "pdf": str(pdf)[:600], "landing": (str(landing)[:600] if landing else None),
            "url": f"https://doi.org/{doi}" if doi else str(pdf)[:600],
            "abstract": abstract[:2000] if abstract else None, "abstract_full": abstract[:9000] if abstract else None,
            "found_in": found_in, "cites": int(cites or 0), "biblio": biblio or {},
            "preprint": _is_preprint(venue, doi)}


# 4.6: what kind of work a paper is (Research lets people choose): a journal article, a review of many studies, a book or a
# chapter, a preprint (not yet peer reviewed), or something else (an editorial, a dataset, a thesis), which Research leaves out
_OA_KIND = {"article": "article", "proceedings-article": "article", "review": "review", "book": "book", "book-chapter": "book",
            "monograph": "book", "book-section": "book", "edited-book": "book", "reference-book": "book", "preprint": "preprint",
            "posted-content": "preprint"}
_OA_TYPES = {"article": "article", "review": "review", "book": "book|book-chapter", "preprint": "preprint",
             # 5.0 (the owner: "remove that limitation by asking users how they want it"): more kinds a person may ask for
             "conference": "article", "thesis": "dissertation", "report": "report"}
# 5.0: the kinds a piece of writing may draw on when the person didn't say (everything but editorials, letters, datasets…)
WRITE_KINDS = ("article", "review", "book", "conference", "thesis", "report", "preprint")


def _oa_kind(w):
    """OpenAlex's type (and its source's type) as one of the kinds people choose from."""
    t = str(w.get("type") or "").lower()
    src = (((w.get("primary_location") or {}).get("source")) or {})
    if t in ("article", "proceedings-article") and str(src.get("type") or "").lower() == "conference":
        return "conference"
    if t == "dissertation":
        return "thesis"
    if t == "report":
        return "report"
    return _OA_KIND.get(t, "other")


def _kind_s2(types, venue=None):
    """Semantic Scholar's publicationTypes: an editorial, a letter, a comment, news or a dataset is never taken for an
    article, even when it is also marked JournalArticle; a record with no type counts as an article only when a journal
    or a venue is named."""
    t = {str(x).lower() for x in (types or [])}
    if t & {"editorial", "lettersandcomments", "news", "dataset"}:
        return "other"
    if t & {"review", "metaanalysis"}:
        return "review"
    if t & {"book", "booksection"}:
        return "book"
    if "conference" in t and not (t & {"journalarticle"}):
        return "conference"
    if t & {"journalarticle", "study", "clinicaltrial", "casereport", "conference"}:
        return "article"
    return "article" if (not t and venue) else "other"


def _kind_epmc(types):
    t = " ".join(str(x).lower() for x in (types or []))
    if "preprint" in t:
        return "preprint"
    if re.search(r"\b(?:thesis|dissertation)\b", t) and not re.search(r"\b(?:retract\w*|erratum|correction)\b", t):
        return "thesis"
    if re.search(r"\b(?:retract\w*|erratum|errata|correction|corrigendum|editorial|comment|letter|news|expression of concern|"
                 r"protocol|guideline)\b", t):
        return "other"
    if re.search(r"\b(?:review|meta-analysis|systematic)\b", t):
        return "review"
    return "article"


def _openalex_papers(query, n, since=None, kinds=None, until=None):
    flt = ("is_oa:true,is_retracted:false" + (f",publication_year:>{since - 1}" if since else "")
           + (f",publication_year:<{until + 1}" if until else ""))
    if kinds and not set(_OA_TYPES) <= set(kinds):
        flt += ",type:" + "|".join(sorted({_OA_TYPES[k] for k in kinds if k in _OA_TYPES}))
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
            p["preprint"] = p["preprint"] or str(w.get("type") or "").lower() == "preprint"
            p["kind"] = "preprint" if p["preprint"] else _oa_kind(w)
            if p["kind"] == "thesis":       # a thesis names the university that gave the degree (its author's, then)
                inst = next((((i or {}).get("display_name")) for a in (w.get("authorships") or [])[:1] for i in ((a or {}).get("institutions") or [])
                             if (i or {}).get("display_name")), None)
                if inst:
                    p["institution"] = str(inst)[:160]
            out.append(p)
    return out, None if out else "OpenAlex found no open papers with a PDF"


def _s2_papers(query, n, since=None, kinds=None, until=None):
    headers = dict(_PAPER_UA)
    if SEMANTIC_SCHOLAR_API_KEY:
        headers["x-api-key"] = SEMANTIC_SCHOLAR_API_KEY
    params = {"query": query[:300], "limit": max(n * 3, 15), "openAccessPdf": "",
              "fields": "title,url,year,authors,abstract,venue,openAccessPdf,externalIds,citationCount,journal,publicationTypes"}
    if since or until:
        params["year"] = f"{since or ''}-{until or ''}"
    if kinds and "article" not in kinds:          # many records carry no type: filter there only when articles aren't wanted
        types = [t for k in kinds for t in {"review": ["Review", "MetaAnalysis"], "book": ["Book", "BookSection"], "conference": ["Conference"]}.get(k, [])]
        if types:
            params["publicationTypes"] = ",".join(types)
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
            p["kind"] = "preprint" if p["preprint"] else _kind_s2(x.get("publicationTypes"), x.get("venue") or (x.get("journal") or {}).get("name"))
            out.append(p)
    return out, None if out else "Semantic Scholar found no open papers with a PDF"


def _epmc_papers(query, n, since=None, kinds=None, until=None):
    q = (f"({query[:250]}) AND OPEN_ACCESS:y AND HAS_PDF:y AND NOT PUB_TYPE:\"retracted publication\""
         + (f" AND PUB_YEAR:[{since or 1000} TO {until or 3000}]" if (since or until) else ""))
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
            p["preprint"] = p["preprint"] or str(x.get("source") or "").upper() == "PPR"
            p["kind"] = "preprint" if p["preprint"] else _kind_epmc(((x.get("pubTypeList") or {}).get("pubType")) or [])
            out.append(p)
    return out, None if out else "Europe PMC found no open papers with a PDF"


def _title_key(t):
    return re.sub(r"[^a-z0-9]+", "", str(t or "").lower())[:120]


def _first_surname(p):
    a = (p.get("authors") or [None])[0]
    if not a:
        return ""
    a = str(a).strip()
    return re.sub(r"[^a-z]", "", (a.split(",")[0] if "," in a else a.split()[-1]).lower())


_GENERIC_TITLES = {"introduction", "editorial", "preface", "foreword", "bookreview", "bookreviews", "commentary", "reply",
                   "response", "erratum", "corrigendum", "correction", "letter", "lettertotheeditor", "abstracts", "conclusion",
                   "conclusions", "discussion", "index", "contents", "frontmatter", "backmatter", "acknowledgements",
                   "acknowledgments", "notes", "review", "overview", "summary", "abstract"}


def _same_work(a, b):
    """Two index records are the same paper: the same DOI, or the same title with years at most one apart and the
    same first author's surname (two different papers that share a title stay apart). When one record lacks its
    authors or its year, the same long, specific title and whatever else both give is enough."""
    if a.get("doi") and b.get("doi") and a["doi"] == b["doi"]:
        return True
    ta = _title_key(a.get("title"))
    if not ta or ta != _title_key(b.get("title")):
        return False
    try:
        ya, yb = int(a.get("year") or 0), int(b.get("year") or 0)
    except (TypeError, ValueError):
        ya = yb = 0
    if ya and yb and abs(ya - yb) > 1:
        return False
    sa, sb = _first_surname(a), _first_surname(b)
    if sa and sb:
        return sa == sb
    return len(ta) >= 25 and ta not in _GENERIC_TITLES


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


def search_papers(query, max_results=8, since=None, check_links=True, kinds=None, stats=None, until=None):
    """Real, downloadable papers for a topic: [{title, authors, year, venue, doi, pdf, url,
    abstract, ...}] most relevant first, and a reason when there are none. kinds (4.6, Research): the kinds of work
    asked for, passed to the indexes that can filter by them. stats (a dict, optional) gets how many different papers
    the indexes found ("found") and which indexes found any ("from"). since / until (5.0): the first and the last
    year of publication wanted."""
    query = re.sub(r"\s+", " ", str(query or "")).strip()
    if not query:
        return [], "no topic"
    order = (("OpenAlex", _openalex_papers), ("Semantic Scholar", _s2_papers), ("Europe PMC", _epmc_papers))
    got, why = {}, []
    with _futures.ThreadPoolExecutor(max_workers=3) as pool:
        more = dict(({"kinds": list(kinds)} if kinds else {}), **({"until": until} if until else {}))
        jobs = {name: pool.submit(fn, query, max_results, since, **more) for name, fn in order}
        for name, _ in order:
            try:
                papers, diag = jobs[name].result(timeout=_PAPER_TIMEOUT * 2 + 3)
            except Exception as e:
                papers, diag = [], f"{name} failed: {e.__class__.__name__}"
            for i, p in enumerate(papers):
                p["_rank"] = i          # its place in that index's own relevance order
            got[name] = papers
            if diag:
                why.append(diag)
    found = [p for name, _ in order for p in got.get(name, [])]
    merged = []
    for p in found:
        twin = next((m for m in merged if _same_work(m, p)), None)
        if twin is None:
            merged.append(p)
            continue
        # the same paper from another index: keep whatever this copy adds. Another version of it (a different DOI —
        # a preprint and its journal article) lends only its abstract, so a reference never mixes two versions.
        if twin.get("doi") and p.get("doi"):
            same_record = twin.get("doi") == p.get("doi")
        else:   # a DOI is missing: one record only when the year, the kind and the journal agree (a preprint is another version)
            vn = lambda x: re.sub(r"[^a-z0-9]+", " ", str(x.get("venue") or "").lower()).strip()
            same_record = bool(twin.get("year") and twin.get("year") == p.get("year") and bool(twin.get("preprint")) == bool(p.get("preprint"))
                               and vn(twin) and vn(twin) == vn(p))
        if same_record and twin.get("abstract_lent") and p.get("abstract"):
            twin["abstract"], twin["abstract_full"], twin["abstract_lent"] = p["abstract"], p.get("abstract_full") or p["abstract"], False
        for k in (("abstract", "venue", "doi", "authors", "year") if same_record else ("abstract",)):
            if not twin.get(k) and p.get(k):
                twin[k] = p[k]
                if k == "abstract":
                    twin["abstract_full"] = p.get("abstract_full") or p[k]
                    # 4.6: another version's abstract (a preprint's, for its journal article) may differ in its figures:
                    # Research never quotes it under this paper's reference
                    twin["abstract_lent"] = not same_record
        if same_record and not twin.get("biblio") and p.get("biblio"):
            twin["biblio"] = p["biblio"]
        twin["preprint"] = bool(twin.get("preprint") or (same_record and p.get("preprint")))
        if twin.get("doi"):
            twin["url"] = f"https://doi.org/{twin['doi']}"
    if stats is not None:
        stats["found"] = len(merged)
        stats["from"] = [name for name, _ in order if got.get(name)]
    if until:          # an index that couldn't filter by the last year (or a record dated later) is filtered here
        merged = [p for p in merged if not p.get("year") or int(p["year"]) <= until]
    if kinds:          # 5.0: an index that couldn't filter by kind is filtered here (a record with no kind counts as an article)
        merged = [p for p in merged if (p.get("kind") or "article") in kinds]
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
    logger.info("search_papers: a %d-character query -> %d paper(s) (%s)", len(query or ""), len(papers), "; ".join(why) or "ok")
    return papers, (None if papers else ("; ".join(why) or "no open papers with a PDF were found"))


def paper_topic_search(text, max_results=8, since=None, until=None, kinds=None):
    """search_papers for what a person wrote: by its topic, and not at all for small talk."""
    query = _search_topic(text)
    if not _worth_searching(query):
        return [], "not a topic"
    return search_papers(query, max_results=max_results, since=since, **({"until": until} if until else {}), **({"kinds": kinds} if kinds else {}))


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


# ------------------------------------------------ 4.5: references and in-text citations, as the style guides say
# Every part comes from the record a paper index gave (authors, year, title, journal, volume, issue, pages, DOI);
# nothing is invented. Checked against: APA 7 (8.17–8.19, 9.8–9.11, 9.44–9.49), Cite Them Right Harvard, MLA 9 and
# Chicago author-date (17th).
_NAME_PARTICLES = {"van", "von", "de", "del", "della", "der", "den", "di", "da", "du", "la", "le", "dos", "das", "do", "ter",
                   "ten", "bin", "ibn", "al", "el", "zu", "vom", "af"}
_NAME_SUFFIX = re.compile(r"^(?:jr|sr|ii|iii|iv)\.?$", re.I)
_ORG_WORDS = re.compile(r"\b(?:Organi[sz]ations?|Collaborat(?:ors?|ions?|ives?)|Groups?|Consorti(?:um|a)|Committees?|Institutes?|Institutions?|"
                        r"Associations?|Councils?|Societ(?:y|ies)|Networks?|Agenc(?:y|ies)|Ministr(?:y|ies)|Departments?|Universit(?:y|ies)|"
                        r"Foundations?|Cent(?:re|er)s?|Commissions?|Teams?|Initiatives?|Program(?:me)?s?|Boards?|Offices?|Authorit(?:y|ies)|"
                        r"Alliances?|Federations?|Confederations?|Unions?|Bureaus?|Corporations?|Compan(?:y|ies)|Inc|Ltd|LLC|Units?|Services?|"
                        r"Panels?|Trusts?|Projects?|Investigators|Working Party|Task Force|Taskforce|Writing Committee|Bank|Fund|Forum|"
                        r"Academ(?:y|ies)|Coalition|Partnership|Secretariat|Assembly|Parliament|Government|Cabinet|Tribunal|Hospitals?|"
                        r"Clinics?|College|Library|Museum|Laborator(?:y|ies)|Observatory|Administration|Corps|League|Cooperative|"
                        r"Co-operative|Nations|Statistics|Survey|Collective)\b")
_SMALL_TITLE = {"a", "an", "the", "and", "but", "or", "nor", "for", "so", "yet", "as", "at", "by", "in", "of", "off", "on", "per", "to",
                "up", "via", "vs", "with", "from", "into", "onto", "over", "upon", "than", "versus",
                # 5.0: MLA 9 and Chicago 17 lower-case prepositions of any length (those that are never anything else)
                "among", "amongst", "amid", "between", "through", "throughout", "during", "without", "within", "against", "across",
                "behind", "below", "beneath", "beside", "besides", "beyond", "despite", "except", "inside", "outside", "toward",
                "towards", "under", "underneath", "unlike"}


def _name_parts(full):
    """One author as {"org": True, "name"} or {"surname", "given": [...], "suffix"}; particles stay with the surname
    ("van den Berg"), Jr./Sr./II stay apart, and a group author ("World Health Organization") stays whole."""
    full = re.sub(r"\s+", " ", str(full or "")).strip().strip(",")
    if not full:
        return None
    if (re.search(r"\d", full) or (_ORG_WORDS.search(full) and not _person_like(full))
            or (full.isupper() and len(full) > 3 and " " not in full)
            or (" " in full and "," not in full and re.search(r"\s(?:the|of|for|on|and|&)\s", full))):   # "Save the Children"
        return {"org": True, "name": full}
    if "," in full:                                   # "Mensah, Kofi A."
        sur, _, giv = full.partition(",")
        given = giv.split()
        suffix = ""
        if given and _NAME_SUFFIX.match(given[-1]):
            suffix = given.pop()
        return {"org": False, "surname": sur.strip(), "given": given, "suffix": suffix}
    toks = full.split(" ")
    suffix = ""
    if len(toks) > 1 and _NAME_SUFFIX.match(toks[-1]):
        suffix = toks.pop()
    if len(toks) == 1:
        return {"org": False, "surname": toks[0], "given": [], "suffix": suffix}
    i = j = len(toks) - 1                              # the surname, with any particles just before it
    while j - 1 >= 1 and toks[j - 1].lower().strip(".") in _NAME_PARTICLES:
        j -= 1
    run = toks[j:i]
    if run and (len(run) >= 2 or any(t[:1].islower() for t in run) or _is_particle(run[-1], toks[i])):
        i = j
    return {"org": False, "surname": " ".join(toks[i:]), "given": toks[:i], "suffix": suffix}


# Surnames from East and South-East Asia that a capitalised "Van", "Da", "Le"… in front of is a given or middle name
# ("Hung Van Tran", "Wei Da Chen"), not a particle ("Martin Van Buren", "Maria Da Silva").
_ASIAN_SURNAMES = set("""nguyen tran le pham hoang huynh phan vu vo dang bui do ho ngo duong ly truong dinh mai trinh lam ha
luong cao doan ta quach thai wang li zhang liu chen yang huang zhao wu zhou xu sun ma zhu hu guo he lin gao luo zheng liang
xie song tang han feng deng peng zeng xiao tian dong pan yuan cai jiang yu du ye cheng wei su lu ding ren shen yao cui zhong
tan fan jin shi liao jia xia fu fang bai zou meng xiong qin qiu hou yin xue yan duan lei long tao mao hao gu gong shao wan
qian lai hong kong chang chan cheung chow lau leung ng wong lee kim park choi jung kang cho yoon jang lim oh seo shin kwon
hwang ahn yoo jeon ko moon son bae baek heo nam shim noh kwak sung cha joo woo koo min ryu na""".split())


def _is_particle(tok, after):
    """"van", "de", "da"… before a surname are part of it; capitalised, they are only when the surname isn't one of the
    East or South-East Asian ones where "Van" or "Da" is a given name."""
    low = tok.lower().strip(".")
    if low not in _NAME_PARTICLES:
        return False
    if tok[:1].islower():
        return True
    return re.sub(r"[^a-z]", "", str(after or "").lower()) not in _ASIAN_SURNAMES


# the first word of a two-word organisation ("World Bank", "Reserve Bank"); "Robert Service" and "Jean Bureau" are people
_ORG_FIRST = re.compile(r"^(?:World|National|International|Global|Central|Federal|Royal|United|European|African|American|Asian|"
                        r"Commonwealth|State|Regional|Local|City|County|Public|Private|Reserve|Development|Education|Health|"
                        r"Statistical|Statistics|Research|Medical|Science|Scientific|Civil|Community|Ghana|Nigeria|Kenya|Africa)$")


def _person_like(full):
    """A two-word name whose org word is only its last ("Robert Service", "Mikkel Bank") names a person."""
    if re.fullmatch(r"[A-ZÀ-ɏ][\w'’\-]+,\s*(?:[A-ZÀ-ɏ]\.\s*-?\s*){1,3}(?:,?\s*(?:Jr|Sr|II|III|IV)\.?)?", full):
        return True                                   # "Bank, A.", "Service, R. W." — a surname, then initials
    toks = full.split()
    if "," in full or len(toks) != 2 or _ORG_WORDS.search(toks[0]) or not _ORG_WORDS.fullmatch(toks[1]):
        return False
    return not (_ORG_FIRST.match(toks[0]) or _PLACES_RX.search(toks[0].lower()))


def _initials(given):
    """["Jean-Pierre", "K."] -> "J.-P. K." (lower-case particles in a given name are left out)."""
    out = []
    for g in given:
        g = g.strip()
        if not g or g[0].islower():
            continue
        if re.fullmatch(r"(?:[A-ZÀ-ɏ]\.?){2,3}", g) and g.replace(".", "").isupper():   # "JR" or "J.R." already initials
            out += [c + "." for c in g.replace(".", "")]
            continue
        out.append("-".join(p[0].upper() + "." for p in g.split("-") if p))
    return " ".join(out)


def _au_surname(a):
    return a["name"] if a.get("org") else a["surname"]


def _au_inverted_apa(a):
    if a.get("org"):
        return a["name"]
    ini = _initials(a["given"])
    return a["surname"] + (", " + ini if ini else "") + (", " + a["suffix"].rstrip(".") + "." if a.get("suffix") else "")


def _au_full(a):
    if a.get("org"):
        return a["name"]
    return " ".join(a["given"] + [a["surname"]]) + (", " + a["suffix"] if a.get("suffix") else "")


def _au_inverted_full(a):
    if a.get("org"):
        return a["name"]
    return a["surname"] + (", " + " ".join(a["given"]) if a["given"] else "") + (", " + a["suffix"] if a.get("suffix") else "")


def _end_dot(s):
    """s with one full stop at its end ("Does it work?" keeps its question mark)."""
    s = str(s or "").rstrip()
    return s if not s or s[-1] in ".?!" else s + "."


def _title_case(t):
    """MLA and Chicago titles: principal words capitalised; words with capitals or digits inside are left as they are."""
    words = t.split(" ")
    out = []
    for i, w in enumerate(words):
        core = w.strip("“”\"'‘’()[]")
        first = i == 0 or (out and out[-1].endswith((":", "?", "!", "—")))
        last = i == len(words) - 1
        if not core or any(c.isdigit() for c in core) or any(c.isupper() for c in core[1:]) or core.startswith(("http", "www")):
            out.append(w)
        elif core.lower() in _SMALL_TITLE and not first and not last:
            out.append(w.replace(core, core.lower(), 1))
        else:
            out.append(w.replace(core, "-".join(p[:1].upper() + p[1:] for p in core.split("-")), 1))
    return " ".join(out)


def _tidy_title(t):
    """A title as published, minus the "Effects Of Peer Tutoring On" quirk of some indexes: a small word inside a title
    that capitalises every word is lower-cased (safe: no proper name is changed)."""
    t = re.sub(r"\s+", " ", str(t or "")).strip().rstrip(".")
    words = t.split(" ")
    caps = [w for w in words if len(w) > 3 and w[:1].isalpha()]
    if caps and sum(1 for w in caps if w[:1].isupper()) >= 0.8 * len(caps):
        for i in range(1, len(words) - 1):
            if words[i] in {w.capitalize() for w in _SMALL_TITLE} and not words[i - 1].endswith((":", "?", "!")):
                words[i] = words[i].lower()
    return " ".join(words)


def _short_title(t, n=4):
    w = re.sub(r"[.:;,?!]+$", "", " ".join(str(t or "").split()[:n]))
    return w


def _sort_title(t):
    return re.sub(r"^(?:a|an|the)\s+", "", str(t or "").lower()).strip()


def _group_norm(s):
    """A name for comparing: accents off, lower case, letters and digits only ("van den Berg" -> "vandenberg")."""
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", "", s)


def _names_group(names, title):
    """What two works must share before their in-text citations can affect each other (initials, more names, a/b):
    the first author's surname (a group's name), or for a work with no author, its short title."""
    if names:   # cut where a citation's own text would be cut ("Department of Health and Social Care" -> "departmentofhealth")
        return _group_norm(re.split(r"\s*(?:,|&|\band\b|\bet al\.?)\s*", _au_surname(names[0]))[0])
    t = _group_norm(_short_title(_tidy_title(title)))
    return "t:" + t if t else ""


def _paper_group(p):
    return _names_group([a for a in (_name_parts(x) for x in (p.get("authors") or [])) if a], p.get("title") or "")


_KEY_YEAR = re.compile(r",?\s+(?:(?:1[5-9]|20)\d{2}[a-z]?|n\.\s?d\.(?:-[a-z])?|no date)(?:\s*,\s*(?:pp?|paras?|chap|ch|sec)\.?\s*[\w–\-]+)?\s*$")


def _key_group(key):
    """The same, worked out from an in-text citation alone: "(J. Lee, 2020)" -> "lee", "(Mensah et al., 2021)" ->
    "mensah", "(“Short Title,” 2021)" -> "t:shorttitle"."""
    s = str(key or "").strip()
    if s.startswith("(") and s.endswith(")"):
        s = s[1:-1]
    s = s.strip()
    if s[:1] == "*":                                  # Harvard sets a no-author title in italics
        return "t:" + _group_norm(s[1:].split("*")[0])
    if s[:1] in "“\"":
        return "t:" + _group_norm(re.split(r"[”\"]", s[1:])[0])
    s = _KEY_YEAR.sub("", s)
    s = re.sub(r"^(?:[A-ZÀ-ɏ]\.\s*(?:-?[A-ZÀ-ɏ]\.\s*)*)", "", s).strip()
    first = re.split(r"\s*(?:,|&|\band\b|\bet al\.?)\s*", s)[0]
    return _group_norm(first)


def _domain_org_name(url):
    """A readable site name (e.g. 'Healthline') for a web page with no author."""
    try:
        host = urlparse(url).netloc.lower()
        if host.startswith("www."):
            host = host[4:]
        base = host.split(".")[0] if host else "Source"
        return base.replace("-", " ").title() or "Source"
    except Exception:
        return "Source"


def _format_author_apa(full_name):
    a = _name_parts(full_name)
    return _au_inverted_apa(a) if a else str(full_name or "")


def _biblio_parts(e):
    """(volume, issue, pages) as the index gave them; pages as "100–110"."""
    b = e.get("biblio") or {}
    vol, iss = str(b.get("volume") or "").strip(), str(b.get("issue") or "").strip()
    first, last = str(b.get("first_page") or "").strip(), str(b.get("last_page") or "").strip()
    pages = f"{first}–{last}" if first and last and first != last else first
    return vol[:20], iss[:20], pages[:30]


CITE_STYLES = {
    "apa": {"name": "APA 7", "heading": "References", "example": "(Smith, 2023)"},
    "harvard": {"name": "Harvard", "heading": "Reference list", "example": "(Smith, 2023)"},
    "mla": {"name": "MLA 9", "heading": "Works Cited", "example": "(Smith)"},
    "chicago": {"name": "Chicago (author-date)", "heading": "References", "example": "(Smith 2023)"},
}


def _intext_names(names, style, n_full=None):
    """The authors as an in-text citation names them. n_full: how many to name before "et al." (APA's way of telling
    apart two works whose "et al." forms would be the same)."""
    sur = [_au_surname(a) for a in names]
    amp = " & " if style == "apa" else " and "
    if n_full and n_full < len(sur):
        return ", ".join(sur[:n_full]) + ", et al."
    if n_full and n_full >= len(sur) and len(sur) > 2:
        return ", ".join(sur[:-1]) + ("," if style in ("apa", "chicago", "mla") else "") + amp + sur[-1]
    if len(sur) == 1:
        return sur[0]
    if len(sur) == 2:
        return sur[0] + amp + sur[1]
    if len(sur) == 3 and style in ("harvard", "chicago"):
        return f"{sur[0]}, {sur[1]}" + ("," if style == "chicago" else "") + f" and {sur[2]}"
    return sur[0] + " et al."


def _it_dot(t):
    """A title in italics, with one full stop after it ("*Does it work?*" keeps its question mark, inside)."""
    t = str(t or "").strip().rstrip(".")
    return f"*{t}*" + ("" if t.endswith(("?", "!")) else ".")


def _no_date(style):
    return "no date" if style == "harvard" else "n.d."


def enrich_sources(sources, style="apa"):
    """Each source with its in-text citation (citation_key) and its reference entry (ref_text) in the chosen style,
    told apart from the others as the style asks (more names, or a/b after the year)."""
    style = style if style in CITE_STYLES else "apa"
    out = []
    for s in sources or []:
        e = dict(s)
        e["_names"] = [a for a in (_name_parts(x) for x in (s.get("authors") or [])) if a]
        e["_title"] = _tidy_title(s.get("title") or "")
        e["_year"] = str(s.get("year")) if s.get("year") else ""
        out.append(e)
    # the in-text names, before telling works apart
    for e in out:
        e["_intext"] = _intext_names(e["_names"], style) if e["_names"] else None
    # APA 8.19 / MLA: two first authors with the same surname: their initials go in front
    firsts = {}
    for e in out:
        if e["_names"] and not e["_names"][0].get("org"):
            firsts.setdefault(e["_names"][0]["surname"].lower(), set()).add(_initials(e["_names"][0]["given"]))
    for e in out:
        if e["_names"] and not e["_names"][0].get("org") and len(firsts.get(e["_names"][0]["surname"].lower(), ())) > 1:
            ini = _initials(e["_names"][0]["given"])
            if ini and style in ("apa", "mla", "chicago", "harvard"):
                e["_intext"] = (ini.split(" ")[0] if style == "mla" else ini) + " " + e["_intext"]
    # APA 8.18: works whose "et al." forms collide name as many authors as it takes to tell them apart
    if style == "apa":
        groups = {}
        for e in out:
            if e["_names"] and len(e["_names"]) >= 3:
                groups.setdefault((e["_intext"], e["_year"]), []).append(e)
        for g in groups.values():
            lists = {tuple(_au_surname(a).lower() for a in e["_names"]) for e in g}
            if len(g) < 2 or len(lists) < 2:
                continue
            for e in g:
                sur = [_au_surname(a).lower() for a in e["_names"]]
                n = 1
                while n < len(sur) and sum(1 for o in g if [_au_surname(a).lower() for a in o["_names"]][:n] == sur[:n]) > 1:
                    n += 1
                e["_intext"] = _intext_names(e["_names"], style, n_full=n)
    # the same authors in the same year: a, b, c after the year, in title order (APA 8.19, Harvard, Chicago)
    letters = {}
    for e in out:
        who = tuple(_au_surname(a).lower() + "|" + _initials(a.get("given") or []) for a in e["_names"]) if e["_names"] else ("t:" + _sort_title(e["_title"]),)
        e["_who"] = who
        letters.setdefault((who, e["_year"]), []).append(e)
    for g in letters.values():
        if len(g) > 1:
            g.sort(key=lambda e: _sort_title(e["_title"]))
            for i, e in enumerate(g):
                e["_letter"] = "abcdefghijklmnopqrstuvwxyz"[i] if i < 26 else str(i)
    for e in out:
        L = e.get("_letter", "")
        yd = (e["_year"] + L) if e["_year"] else (_no_date(style) + ("-" + L if L and style != "harvard" else (L if L else "")))
        e["year_display"] = yd
        title, venue = e["_title"], (e.get("venue") or "").strip()
        vol, iss, pages = _biblio_parts(e)
        pre = " [Preprint]" if e.get("preprint") else ""
        doi_url = f"https://doi.org/{e['doi']}" if e.get("doi") else ""
        if style == "apa":
            if e["_names"]:
                refs = [_au_inverted_apa(a) for a in e["_names"]]
                if len(refs) >= 21:
                    auth = ", ".join(refs[:19]) + ", . . . " + refs[-1]
                elif len(refs) == 1:
                    auth = refs[0]
                else:
                    auth = ", ".join(refs[:-1]) + ", & " + refs[-1]
                key = f"({e['_intext']}, {yd})"
                head = f"{_end_dot(auth)} ({yd}). "
                body = (f"*{title}*{pre}" + ("" if (not pre and title.endswith(("?", "!"))) else ".")) if (not venue or e.get("preprint")) else _end_dot(title)
            else:                                       # no author: the title takes the author's place (APA 9.12)
                key = f"(“{_title_case(_short_title(title))},” {yd})"
                head = ""
                body = _end_dot(title) + f" ({yd})."
            tail = ""
            if venue and not e.get("preprint"):
                tail = f" *{venue}*" + (f", *{vol}*" if vol else "") + (f"({iss})" if vol and iss else "") + (f", {pages}" if pages else "") + "."
            elif venue:
                tail = f" {venue}."
            # 5.0: a paper in published proceedings (APA 10.5), a thesis (10.6) and a report (10.4)
            kind, inst = e.get("kind"), (e.get("institution") or "").strip()
            if kind == "conference" and venue and not e.get("preprint"):
                body, tail = _end_dot(title), f" In *{venue}*" + (f" (pp. {pages})" if pages else "") + "."
            elif kind == "thesis" and not e.get("preprint"):
                body, tail = f"*{title}*" + (f" [Thesis, {inst}]." if inst else " [Thesis]."), (f" {venue}." if venue and venue != inst else "")
            elif kind == "report" and not e.get("preprint"):
                body, tail = _it_dot(title), (f" {venue}." if venue else "")
            if kind in ("conference", "thesis", "report") and not e["_names"] and not e.get("preprint"):
                head, body = "", body + f" ({yd})."
            e["ref_text"] = (head + body + tail).strip()
        elif style == "harvard":
            nm = e["_names"]
            avail = " Available at:" if (doi_url or e.get("url")) else ""
            if e.get("url") and not doi_url:          # Cite Them Right: a link with no DOI is followed by when it was read
                e["ref_after"] = f"(Accessed: {_today_words()})."
            if venue and not e.get("preprint"):
                src = f"*{venue}*" + (f", {vol}" if vol else "") + (f"({iss})" if vol and iss else "") + (f", pp. {pages}" if pages else "") + "."
            else:
                src = ""
            kind, inst = e.get("kind"), (e.get("institution") or "").strip()
            special = None       # 5.0: Cite Them Right's conference paper, thesis and report
            if kind == "conference" and venue and not e.get("preprint"):
                special = f"‘{title}’, in *{venue}*" + (f", pp. {pages}" if pages else "") + "."
            elif kind == "thesis" and not e.get("preprint"):
                special = _it_dot(title) + " Thesis." + (f" {inst}." if inst else "")
            elif kind == "report" and not e.get("preprint"):
                special = _it_dot(title) + (f" {venue}." if venue else "")
            if nm:
                parts = [_au_inverted_apa(a) for a in nm]
                auth = parts[0] + " et al." if len(parts) > 3 else parts[0] if len(parts) == 1 else ", ".join(parts[:-1]) + " and " + parts[-1]
                key = f"({e['_intext']}, {yd})"
                if special:
                    body = special
                elif src:
                    body = f"‘{title}’, {src}"
                else:
                    body = f"*{title}*" + (f" [Preprint]. {venue}." if e.get("preprint") and venue else ("" if title.endswith(("?", "!")) else "."))
                e["ref_text"] = f"{auth} ({yd}) {body}{avail}"
            else:
                key = f"(*{_short_title(title)}*, {yd})"
                if special:
                    e["ref_text"] = re.sub(r"^(‘[^’]*’|\*[^*]*\*\.?)", lambda m: m.group(1).rstrip(".") + f" ({yd})", special, count=1) + avail
                else:
                    e["ref_text"] = (f"‘{title}’ ({yd}) {src}" if src else f"*{title}* ({yd})" + (f" [Preprint]. {venue}." if e.get("preprint") and venue else ".")) + avail
        elif style == "chicago":
            nm = e["_names"]
            full = [_au_full(a) for a in nm]
            if not nm:
                auth = ""
            elif len(full) == 1:
                auth = _au_inverted_full(nm[0])
            elif len(full) > 10:
                auth = _au_inverted_full(nm[0]) + ", " + ", ".join(full[1:7]) + ", et al."
            elif len(full) == 2:
                auth = f"{_au_inverted_full(nm[0])}, and {full[1]}"
            else:
                auth = _au_inverted_full(nm[0]) + ", " + ", ".join(full[1:-1]) + ", and " + full[-1]
            ttl = _title_case(title)
            key = f"({e['_intext']} {yd})" if nm else f"(“{_short_title(ttl)}” {yd})"
            src = (f" *{venue}*" + (f" {vol}" if vol else "") + (f" ({iss})" if vol and iss else "") + (f": {pages}" if pages else "") + ".") if venue and not e.get("preprint") else (f" {venue}." if venue else "")
            lead = (_end_dot(auth) + " " + _end_dot(yd) + " ") if nm else ""
            ttl_part = f"“{_end_dot(ttl)}”" + (" Preprint." if e.get("preprint") else "")
            # 5.0: a paper in proceedings, a thesis and a report (Chicago 17, author-date)
            kind, inst = e.get("kind"), (e.get("institution") or "").strip()
            if kind == "conference" and venue and not e.get("preprint"):
                src = f" In *{venue}*" + (f", {pages}" if pages else "") + "."
            elif kind == "thesis" and not e.get("preprint"):
                src = " Thesis" + (f", {inst}" if inst else "") + "."
            elif kind == "report" and not e.get("preprint"):
                ttl_part, src = _it_dot(ttl), (f" {venue}." if venue else "")
            e["ref_text"] = (lead + ttl_part + (" " + _end_dot(yd) if not nm else "") + src).strip()
        else:   # MLA 9
            nm = e["_names"]
            ttl = _title_case(title)
            if nm:
                auth = _au_inverted_full(nm[0]) if len(nm) == 1 else (f"{_au_inverted_full(nm[0])}, and {_au_full(nm[1])}" if len(nm) == 2 else f"{_au_inverted_full(nm[0])}, et al.")
                key = f"({e['_intext']})"
                lead = _end_dot(auth) + " "
            else:
                key = f"(“{_short_title(ttl)}”)"
                lead = ""
            kind, inst = e.get("kind"), (e.get("institution") or "").strip()
            bits = []
            if venue:
                bits.append(f"*{_title_case(venue)}*" if not e.get("preprint") else venue)
            if vol:
                bits.append(f"vol. {vol}")
            if iss:
                bits.append(f"no. {iss}")
            if e["_year"]:
                bits.append(e["_year"])
            if pages:
                bits.append(f"pp. {pages}")
            e["ref_text"] = (lead + f"“{_end_dot(ttl)}” " + (", ".join(bits) + "." if bits else "")).strip()
            # 5.0: a thesis (MLA 9: title in italics, the year, the institution and the kind of work) and a report
            if kind == "thesis" and not e.get("preprint"):
                e["ref_text"] = (lead + _it_dot(ttl) + (f" {e['_year']}." if e["_year"] else "") + (f" {inst}, Thesis." if inst else " Thesis.")).strip()
            elif kind == "report" and not e.get("preprint"):
                e["ref_text"] = (lead + _it_dot(ttl) + " " + ", ".join(x for x in (_title_case(venue) if venue else "", e["_year"]) if x) + ".").strip().replace(" .", ".")
        e["citation_key"] = key
        e["author_intext"] = e["_intext"] or _short_title(title)
        e["_sort"] = (tuple((_au_surname(a).lower(), _initials(a.get("given") or []).lower()) for a in e["_names"]) if e["_names"] else ((_sort_title(title), ""),),
                      -1 if not e["_year"] else int(e["_year"]) if e["_year"].isdigit() else 9999, _sort_title(title))
    if style == "mla":     # two works by one author: MLA tells them apart by a short title
        counts = Counter(e["citation_key"] for e in out)
        for e in out:
            if counts[e["citation_key"]] > 1 and e["_names"]:
                e["citation_key"] = e["citation_key"][:-1] + f", “{_short_title(_title_case(e['_title']), 3)}”)"
    for e in out:
        e["group"] = _names_group(e["_names"], e["_title"])
        for k in ("_names", "_title", "_year", "_intext", "_letter", "_who"):
            e.pop(k, None)
    return out


def enrich_sources_for_apa(sources):
    return enrich_sources(sources, "apa")


def _today_words():
    d = datetime.now(timezone.utc)
    return f"{d.day} {d.strftime('%B')} {d.year}"


def build_sources_block(sources):
    """Source list for the writing prompt, keyed by the exact citation (in the chosen style) the model must copy
    verbatim, with a short excerpt of each abstract."""
    lines = []
    for s in sources:
        line = f"{s['citation_key']} {s['title']}"
        if s.get("venue") or s.get("year"):
            line += f" — {s.get('venue') or 'paper'}, {s.get('year') or 'n.d.'}"
        if s.get("abstract"):
            line += f"\n    Abstract: {s['abstract'][:700]}"
        lines.append(line)
    return "\n".join(lines)


def build_references_list(sources, cited=None):
    """The reference list: only the sources the text cites (all of them when cited is None), in the order the style
    asks (by authors' surnames, then year with n.d. first, then title). Each entry: its text, its link (DOI, or the
    PDF), the PDF itself, the DOI, and its in-text key."""
    picked = [s for i, s in enumerate(sources) if cited is None or i in cited]
    ordered = sorted(picked, key=lambda s: s.get("_sort") or (((str(s.get("author_intext") or "").lower(), ""),), 0, ""))
    return [{"text": s["ref_text"], "url": s["url"], "pdf": s.get("pdf") or None, "doi": s.get("doi") or None,
             "key": s.get("citation_key") or "", "group": s.get("group") or _key_group(s.get("citation_key")),
             "sort": _sort_json(s.get("_sort")), "after": s.get("ref_after") or ""} for s in ordered]


def _sort_json(k):
    """A reference's place in the list, as plain data the app can send back (to continue a piece of writing)."""
    try:
        names, year, title = k
        return [[[str(a)[:120], str(b)[:40]] for a, b in names][:30], int(year), str(title)[:300]]
    except Exception:
        return None


def _sort_back(v, text=""):
    """The place sent back by the app, as a sort key again; for an older list without one, worked out from the text."""
    try:
        names, year, title = v
        return (tuple((str(a)[:120], str(b)[:40]) for a, b in names[:30]), int(year), str(title)[:300])
    except Exception:
        pass
    t = re.sub(r"[*_]", "", str(text or "")).strip()
    m = re.search(r"(?:\(|\.\s+)(?:(1[5-9]\d{2}|20\d{2})[a-z]?|n\.\s?d\.|no date)", t)
    year = int(m.group(1)) if m and m.group(1) else (-1 if m else 9999)
    lead = t[:m.start() + (1 if m and t[m.start()] == "." else 0)] if m else t
    names = []
    for part in re.split(r",?\s*&\s*|,?\s+and\s+|(?<=\.),\s+(?=[A-ZÀ-ɏ])", lead.strip()):
        sur, _, ini = part.partition(",")
        if sur.strip():
            names.append((sur.strip().lower(), re.sub(r"\s+", " ", ini.strip()).lower()))
    return (tuple(names) or ((t.lower(), ""),), year, t.lower())


# An in-text citation's shape: a capitalised name, then the year (or n.d.), maybe a page.
_CITE_SHAPE = re.compile(r"^[A-ZÀ-ɏ“\"*][^;()]{0,120}?,?\s+(?:(?:1[5-9]|20)\d{2}[a-z]?|n\.\s?d\.(?:-[a-z])?|no date)"
                         r"(?:\s*,\s*(?:p|pp|para|chap|ch|sec)\.?\s*[\w–\-]+)?$")
# MLA's: a name, maybe "and Name" or "et al.", maybe a page
_MLA_SHAPE = re.compile(r"^[A-ZÀ-ɏ][\w'’\-]+(?:,?\s+(?:and|&)\s+[A-ZÀ-ɏ][\w'’\-]+|\s+et al\.)?(?:,?\s+\d+(?:[–\-]\d+)?)?$")
_CITE_LEAD = re.compile(r"^(?:see also|see|e\.g\.,?|cf\.|for example,?|as cited in)\s+", re.I)
_CITE_LOCATOR = re.compile(r"(?:,\s*|\s+)(?:(?:pp?|paras?|chap|ch|sec)\.\s*)?\d+(?:\s*[–\-]\s*\d+)?\s*$", re.I)
_NARRATIVE = re.compile(r"([A-ZÀ-ɏ][\w'’\-]+(?:\s+(?:and|&)\s+[A-ZÀ-ɏ][\w'’\-]+|,?\s+et al\.)?)(?:['’]s)?\s*"
                        r"\(((?:1[5-9]|20)\d{2}[a-z]?|n\.\s?d\.(?:-[a-z])?)(?:,\s*(?:pp?|paras?)\.\s*[\d–\-]+)?\)")
# a page or paragraph given with "p."/"para.": a key the app made never has one, a person's own may
_CITE_PAGE = re.compile(r",?\s*(?:pp?|paras?|chap|ch|sec)\.\s*[\w–\-]+(?:\s*[–\-]\s*\w+)?\s*$", re.I)
# the word before "(2008)" names a law, an event, a survey or a report, or a month, not an author: such a phrase is left
# to the sentence check (which keeps it only when a source supports it)
_NARR_NOT_AUTHOR = re.compile(r"^(?:Acts?|Laws?|Bills?|Polic(?:y|ies)|Constitution|Amendments?|Agreements?|Accords?|Treat(?:y|ies)|"
                              r"Conventions?|Declarations?|Charters?|Protocols?|Census|Reforms?|Pandemic|Epidemic|Outbreak|"
                              r"Elections?|War|Crisis|Summit|Conference|Games|Olympics|Budget|Plan|Strategy|Framework|Agenda|"
                              r"Programmes?|Programs?|Initiatives?|Projects?|Campaigns?|Goals|Surveys?|Data|Dataset|Database|"
                              r"Index|Reports?|Stud(?:y|ies)|Cohort|Atlas|Yearbook|Statistics|Assessments?|"
                              r"January|February|April|July|August|September|October|November|December)$")
_PLACE_SURNAMES = {"england", "ireland", "wales", "france", "london", "berlin", "madrid", "colombo", "lagos", "kano", "cairo", "mali",
                   "brazil", "canada", "geneva", "amman", "vancouver", "dublin", "athens", "rome", "lisbon", "amsterdam", "zurich",
                   "manila", "durban", "rabat", "tripoli", "juba", "blantyre", "india", "china", "kenya", "turkey", "chile", "cuba"}
# Countries, regions and large cities (names that are also common surnames, like Jordan or Washington, left out): a
# place before "(2012)" ("Literacy rose in Rwanda (2012)") or in brackets ("(Accra)") is not an author
_PLACES_RX = re.compile(r"(?:^|[^a-z'])(?:" + "|".join(sorted({re.escape(x.strip()) for x in """
afghanistan|albania|algeria|andorra|angola|antigua and barbuda|argentina|armenia|australia|austria|azerbaijan|bahamas|bahrain|
bangladesh|barbados|belarus|belgium|belize|benin|bhutan|bolivia|bosnia and herzegovina|botswana|brazil|brunei|bulgaria|
burkina faso|burundi|cabo verde|cape verde|cambodia|cameroon|canada|central african republic|chile|china|colombia|comoros|congo|
costa rica|cote d'ivoire|côte d'ivoire|ivory coast|croatia|cuba|cyprus|czechia|czech republic|denmark|djibouti|dominica|
dominican republic|drc|ecuador|egypt|el salvador|equatorial guinea|eritrea|estonia|eswatini|swaziland|ethiopia|fiji|finland|
france|gabon|gambia|the gambia|germany|ghana|greece|grenada|guatemala|guinea|guinea-bissau|guyana|haiti|honduras|hungary|iceland|
india|indonesia|iran|iraq|ireland|italy|jamaica|japan|kazakhstan|kenya|kiribati|kosovo|kuwait|kyrgyzstan|laos|latvia|lebanon|
lesotho|liberia|libya|liechtenstein|lithuania|luxembourg|madagascar|malawi|malaysia|maldives|mali|malta|mauritania|mauritius|
mexico|micronesia|moldova|monaco|mongolia|montenegro|morocco|mozambique|myanmar|burma|namibia|nauru|nepal|netherlands|
new zealand|nicaragua|niger|nigeria|north korea|north macedonia|norway|oman|pakistan|palau|palestine|panama|papua new guinea|
paraguay|peru|philippines|poland|portugal|qatar|romania|russia|rwanda|saint lucia|samoa|san marino|saudi arabia|senegal|serbia|
seychelles|sierra leone|singapore|slovakia|slovenia|solomon islands|somalia|somaliland|south africa|south korea|korea|
south sudan|spain|sri lanka|sudan|suriname|sweden|switzerland|syria|taiwan|tajikistan|tanzania|thailand|timor-leste|togo|tonga|
trinidad and tobago|tunisia|turkey|türkiye|turkmenistan|tuvalu|uganda|ukraine|united arab emirates|uae|united kingdom|uk|
britain|great britain|england|scotland|wales|northern ireland|united states|usa|us|the us|the uk|uruguay|uzbekistan|vanuatu|
venezuela|vietnam|viet nam|yemen|zambia|zimbabwe|hong kong|macau|puerto rico|greenland|
africa|asia|europe|oceania|antarctica|north america|south america|latin america|central america|caribbean|the caribbean|
middle east|sub-saharan africa|west africa|east africa|north africa|southern africa|central africa|south asia|east asia|
southeast asia|south-east asia|central asia|western europe|eastern europe|scandinavia|sahel|the sahel|
accra|kumasi|tamale|cape coast|takoradi|lagos|abuja|kano|ibadan|nairobi|mombasa|kisumu|kampala|kigali|dar es salaam|dodoma|
addis ababa|cairo|johannesburg|cape town|pretoria|durban|lusaka|harare|bulawayo|maputo|luanda|kinshasa|dakar|abidjan|bamako|
ouagadougou|niamey|lomé|lome|cotonou|freetown|monrovia|conakry|banjul|nouakchott|rabat|casablanca|algiers|tunis|tripoli|
khartoum|juba|mogadishu|asmara|antananarivo|lilongwe|blantyre|gaborone|windhoek|maseru|mbabane|yaoundé|yaounde|douala|libreville|
brazzaville|bangui|n'djamena|london|berlin|madrid|rome|lisbon|dublin|amsterdam|brussels|vienna|zurich|geneva|stockholm|oslo|
copenhagen|helsinki|warsaw|prague|budapest|athens|moscow|kyiv|kiev|istanbul|ankara|tehran|baghdad|riyadh|dubai|doha|beirut|
jerusalem|amman|delhi|new delhi|mumbai|kolkata|chennai|bangalore|bengaluru|karachi|lahore|islamabad|dhaka|kathmandu|colombo|
beijing|shanghai|tokyo|seoul|manila|jakarta|bangkok|hanoi|ho chi minh city|kuala lumpur|melbourne|auckland|toronto|montreal|
vancouver|ottawa|new york|los angeles|mexico city|bogotá|bogota|buenos aires|são paulo|sao paulo|rio de janeiro|brasília|
brasilia|caracas|havana""".replace("\n", "").split("|") if x.strip()}, key=len, reverse=True)) + r")$")


def _cite_norm(s):
    s = str(s or "").replace("&", " and ").replace("’", "'").replace("‘", "'").replace("“", '"').replace("”", '"').replace("*", "").lower()
    s = re.sub(r"\s+", " ", s)
    return s.strip(" ,.;")


def _narr_key(s):
    """A narrative citation as compared with the person's own: "Adams' (2019)" and "d'Adams (2019)" are Adams (2019)."""
    return re.sub(r"\s+", " ", re.sub(r"'s?\s*(?=\()", " ", _cite_norm(s))).strip()


def _narr_tail(s):
    """A narrative citation's words before the year, compared: "Mensah's", "Adams'" are Mensah, Adams."""
    return {s, re.sub(r"'s$", "", s), re.sub(r"'$", "", s)}


def keep_known_citations(text, sources, style="apa", locators=False, allow=None):
    """Keeps the in-text citations that exactly match a real source and takes out any other, so every citation left
    leads to the right paper. A page given after a real key ("p. 45") is taken off unless locators is set (the person's
    own text): the app has a paper's abstract, never its pages, so a page it wrote would be made up. A sentence that
    names a study in the narrative way ("Johnson (2019) found…") that is not one of the sources goes too; a known one
    is recognised by its full name ("The World Health Organization (2023)", "van den Berg (2022)", "J. Lee (2020)").
    allow: narrative citations (as _cite_norm gives them) to keep as they are, e.g. those in the person's own text.
    Returns (text, cited): the indices of the sources the text cites."""
    if not text:
        return text, set()
    inners, narr = {}, {}
    for i, s in enumerate(sources or []):
        key = str(s.get("citation_key") or "")
        if key.startswith("(") and key.endswith(")"):
            inner = key[1:-1]
            plain = _CITE_PAGE.sub("", inner)
            inners.setdefault(_cite_norm(inner), i)
            inners.setdefault(_cite_norm(plain), i)
            m = re.match(r"^(.*?)(?:,\s*|\s+)((?:1[5-9]|20)\d{2}[a-z]?|n\.\s?d\.(?:-[a-z])?|no date)$", plain)
            if m:
                narr.setdefault((_cite_norm(m.group(1)), _cite_norm(m.group(2))), i)
        if style == "mla" and s.get("author_intext") and s.get("year"):
            narr.setdefault((_cite_norm(s["author_intext"]), _cite_norm(str(s["year"]))), i)
    cited = set()
    words = set(re.findall(r"[a-zà-ɏ'’\-]+", " ".join(f"{s.get('title') or ''} {s.get('abstract') or ''} {s.get('venue') or ''}"
                                                        for s in sources or []).lower()))
    stems = {w[:5] for w in words if len(w) >= 4}

    def known(part):
        part = part.strip()
        lead = _CITE_LEAD.match(part)
        p = _CITE_LEAD.sub("", part)
        if _cite_norm(p) in inners:
            return inners[_cite_norm(p)], part
        q = _CITE_LOCATOR.sub("", p).rstrip(" ,")
        if _cite_norm(q) in inners:
            return inners[_cite_norm(q)], (part if locators else (lead.group(0) if lead else "") + q)
        return None, part

    def shaped(part):
        p = _CITE_LEAD.sub("", part.strip())
        if re.match(r"^(?:Figure|Fig|Table|Chapter|Section|Appendix|Equation|Eq|Step|Part|Item|Question|Page|Note|Box|Panel|Model|Study|"
                    r"Phase|Grade|Class|Level|Unit|Week|Day|Room|Article|Rule|Act|Year|Form|Stage|Wave|Session|Lesson|Module|Round|"
                    r"Cohort|Experiment|Trial|Hypothesis|Objective|Aim|Goal|Standard|Primary|Basic|Volume|Vol|Issue|No|Number|"
                    r"Version|Type|Tier|Sample|Group|Arm|Site|School|Region|District|Zone|Month|Quarter|Semester|Term|Period|Theme|"
                    r"Cycle|Participant|Respondent|Interviewee|Patient|Case|Condition|Scenario|Task|Visit|Dose|Factor|Component|"
                    r"Dimension|Domain|Category|Code|Station|Batch|Teacher|Student|Pupil|Child|Subject|Lab|Ward|Block|Plot)\b", p):
            return False                # "(Figure 2)", "(Table 1)", "(Year 12)": not citations
        if style == "mla":
            return bool(_MLA_SHAPE.match(p)) and bool(re.search(r"\d|\band\b|&|et al", p))
        # an MLA-shaped "(Johnson 12)" in another style is a citation too; an acronym, or a word the sources use
        # ("(Theme 3)", "(Participant 12)"), isn't
        if _CITE_SHAPE.match(p):
            return True
        if not (_MLA_SHAPE.match(p) and re.search(r"\d", p)) or p.split()[0].isupper():
            return False
        head = re.sub(r",?\s+\d.*$", "", p)
        return not any(w.lower()[:5] in stems for w in re.findall(r"[A-Za-zÀ-ɏ]+", head) if len(w) >= 4)

    def bare_name(part, m):
        # MLA cites by a surname alone, "(Johnson)": at the end of a sentence, a name that no source has anywhere is one
        if style != "mla" or not words:
            return False
        p = _CITE_LEAD.sub("", part.strip())
        if not re.fullmatch(r"[A-ZÀ-ɏ][a-zà-ɏ'’\-]+(?:\s+[A-ZÀ-ɏ][a-zà-ɏ'’\-]+)?", p):
            return False
        if not re.match(r"\s*(?:[.!?;,:]|$)", m.string[m.end():m.end() + 3]):
            return False
        if _PLACES_RX.search(_cite_norm(p)) and _cite_norm(p) not in _PLACE_SURNAMES:
            return False
        return not any(w.lower() in words for w in p.split())

    def fix(m):
        parts = [p.strip() for p in m.group(1).split(";")]
        kept, dropped, changed = [], False, False
        for p in parts:
            i, clean = known(p)
            if i is not None:
                cited.add(i)
                changed = changed or clean != p
                kept.append(clean)
            elif shaped(p) or bare_name(p, m):
                dropped = True          # looks like a citation, but no such source: out
            else:
                kept.append(p)
        if not dropped and not changed:
            return m.group(0)
        return "(" + "; ".join(kept) + ")" if kept else ""

    out = re.sub(r"\(([^()]{2,400})\)", fix, text)

    allow = {_narr_key(x) for x in (allow or set())}

    def ends(text, name):
        return text == name or (text.endswith(name) and not text[-len(name) - 1].isalnum())

    def narr_known(m):
        # the source's whole name ends right before "(year)", covers every name the citation gives ("Park and Lee
        # (2020)" is not Lee (2020)), and is not the end of a longer list ("Park, Mensah and Boateng (2022)")
        yr = _cite_norm(m.group(2))
        own = _narr_tail(_cite_norm(m.group(1)))
        before = _cite_norm(m.string[max(0, m.start(1) - 300):m.end(1)])
        for b in _narr_tail(before):
            for (nm, yy), i in narr.items():
                if yy != yr or not nm or not ends(b, nm) or not any(ends(nm, g) for g in own):
                    continue
                if re.search(r"(?<!\))(?:,|\s+and)$", b[:len(b) - len(nm)].rstrip()):
                    continue
                return i
        return None

    # narrative citations: "Mensah and Boateng (2022) found…" — a known one counts (a page the app gave it is taken
    # off); an unknown one takes its sentence out
    def narr_fix(m):
        if locators or "," not in m.group(0)[m.group(0).rfind("("):] or narr_known(m) is None:
            return m.group(0)
        return m.group(0)[:m.group(0).rfind("(")] + "(" + m.group(2) + ")"

    out = _NARRATIVE.sub(narr_fix, out)
    bad = []
    for m in _NARRATIVE.finditer(out):
        if _narr_key(m.group(0)) in allow:
            continue
        i = narr_known(m)
        if i is not None:
            cited.add(i)
            continue
        pre = out[max(0, m.start(1) - 120):m.start(1)]
        after_prep = bool(re.search(r"\b(?i:in|across|throughout|during|from|of|at|within|outside|to|into|for|on|since|until|by)\s+"
                                    r"(?:the\s+)?(?:[A-ZÀ-ɏ][\w'’\-]*\s+){0,3}$", pre))
        after_the = bool(re.search(r"\b[Tt]he\s+(?:[A-ZÀ-ɏ][\w'’\-]*\s+){0,6}$", pre))
        last = m.group(1).split()[-1].strip(".,")
        if " and " not in m.group(1) and " & " not in m.group(1) and (
                re.search(r"\d", m.group(1))
                or (_NARR_NOT_AUTHOR.match(last) and (after_the or after_prep))
                or (after_prep and _PLACES_RX.search(_cite_norm(out[max(0, m.start(1) - 60):m.end(1)])))):
            continue
        bad.append(m.start())
    if bad:
        spans = _c34_sentences(out)
        for a, b in sorted({sp for sp in spans for x in bad if sp[0] <= x < sp[1]}, key=lambda sp: -sp[0]):
            out = _c34_cut(out, a, b)
        out = _c34_tidy(out) if "\n" in out else out.strip()
    out = re.sub(r"[ \t]+([.,;:])", r"\1", out)
    out = re.sub(r"[ \t]{2,}", " ", out)
    out = re.sub(r"[ \t]+(?=\n|$)", "", out)
    return out, cited


# ---------- 5.0 citation styles beyond the four (the owner, 10 Nov 2026: "Remove that limitations by asking users how they
# want it"). The writing is always made, and checked, with author-date citations (every check above relies on them); only
# then is it turned into the style the person asked for:
#   - Vancouver and IEEE, built exactly: numbers in order of first citation, and the reference list in that order;
#   - another style: the person says how its citations look in the text, and may give one reference as their guide
#     writes it. RamzAI lays out each reference like that example, and each one is checked against the paper's own data
#     (every name, number and word must be the paper's); one that fails is written in the nearest built-in style, and
#     the person is told.
NUM_STYLES = {"vancouver": {"name": "Vancouver", "heading": "References"}, "ieee": {"name": "IEEE", "heading": "References"}}
NUM_MARKS = ("brackets", "parens", "raised")
OTHER_INTEXT = {"author-comma": "harvard", "author-space": "chicago", "author-only": "mla",
                "brackets": "numbers", "parens": "numbers", "raised": "numbers"}
_SUP = str.maketrans("0123456789,-", "⁰¹²³⁴⁵⁶⁷⁸⁹˒⁻")


def _cite_plan(data):
    """What a request asks for: (the style the writing is made and checked in, what it is turned into). The second is
    None for the four built-in styles, else {"style", "name", "marks", "refs", "example"}."""
    style = str(data.get("style") or "apa").lower()
    if style in CITE_STYLES:
        return style, None
    if style in NUM_STYLES:
        marks = "ieee" if style == "ieee" else (str(data.get("cite_marks") or "parens") if str(data.get("cite_marks") or "") in NUM_MARKS else "parens")
        return "apa", {"style": style, "name": NUM_STYLES[style]["name"], "marks": marks, "refs": style, "example": ""}
    if style == "other":
        o = data.get("cite_other") if isinstance(data.get("cite_other"), dict) else {}
        name = re.sub(r"\s+", " ", str(o.get("name") or "")).strip()[:60] or "Your style"
        shape = str(o.get("intext") or "author-comma")
        shape = shape if shape in OTHER_INTEXT else "author-comma"
        example = re.sub(r"[ \t]+", " ", str(o.get("example") or "")).strip()[:600]
        if OTHER_INTEXT[shape] == "numbers":
            return "apa", {"style": "other", "name": name, "marks": shape, "refs": "vancouver", "example": example}
        base = OTHER_INTEXT[shape]
        return base, {"style": "other", "name": name, "marks": None, "refs": base, "example": example}
    return "apa", None


def _van_author(a):
    """Vancouver (NLM): the surname, then the initials with no full stops or spaces ("Mensah KA")."""
    if a.get("org"):
        return a["name"]
    ini = ""
    for g in a.get("given") or []:
        g = g.strip()
        if not g or g[0].islower():
            continue
        if re.fullmatch(r"(?:[A-ZÀ-ɏ]\.?){1,3}", g) and g.replace(".", "").isupper():
            ini += g.replace(".", "")
        else:
            ini += "".join(p[0].upper() for p in re.split(r"[-\s.]+", g) if p)
    return a["surname"] + (" " + ini if ini else "") + (" " + a["suffix"].rstrip(".") if a.get("suffix") else "")


def _ieee_author(a):
    """IEEE: the initials first, with full stops ("K. A. Mensah")."""
    if a.get("org"):
        return a["name"]
    ini = _initials(a.get("given") or [])
    return (ini + " " if ini else "") + a["surname"] + (", " + a["suffix"] if a.get("suffix") else "")


def _nlm_pages(pages):
    """NLM's page ranges: the repeated leading digits of the last page are left out ("123–128" -> "123-8")."""
    a, _, b = str(pages or "").replace("–", "-").partition("-")
    a, b = a.strip(), b.strip()
    if a.isdigit() and b.isdigit() and len(a) == len(b) and a != b:
        i = 0
        while i < len(a) - 1 and a[i] == b[i]:
            i += 1
        return f"{a}-{b[i:]}"
    return f"{a}-{b}" if a and b and a != b else a


def numbered_ref(src, style):
    """One source's reference in Vancouver (NLM, as most universities teach it: up to six authors) or IEEE."""
    names = [a for a in (_name_parts(x) for x in (src.get("authors") or [])) if a]
    title = _tidy_title(src.get("title") or "")
    venue = (src.get("venue") or "").strip()
    year = str(src.get("year") or "")
    vol, iss, pages = _biblio_parts(src)
    doi = _norm_doi(src.get("doi"))
    url = str(src.get("url") or "").strip()
    kind = "preprint" if src.get("preprint") else (src.get("kind") or "article")
    inst = (src.get("institution") or "").strip()
    if style == "vancouver":
        au = [_van_author(a) for a in names]
        auth = (", ".join(au[:6]) + ", et al." if len(au) > 6 else ", ".join(au) + ".") if au else ""
        yr = year or "[date unknown]"
        t = _end_dot(title)
        link = f" doi:{doi}" if doi else (f" Available from: {url}" if url else "")
        if kind == "preprint":
            body = f"{title} [Preprint]." + (f" {venue}." if venue else "") + f" {yr}."
        elif kind == "thesis":
            body = f"{title} [thesis]." + (f" {inst};" if inst else "") + f" {yr}."
        elif kind == "conference" and venue:
            body = f"{t} In: {venue}; {yr}." + (f" p. {_nlm_pages(pages)}." if pages else "")
        elif kind in ("book", "report") or not venue:
            body = f"{t}" + (f" {venue};" if venue else "") + f" {yr}."
        else:
            body = f"{t} {venue}. {yr}" + (f";{vol}" if vol else "") + (f"({iss})" if iss else "") + (f":{_nlm_pages(pages)}" if pages else "") + "."
        return re.sub(r"\s+", " ", f"{auth} {body}{link}").strip()
    # IEEE
    au = [_ieee_author(a) for a in names]
    if len(au) > 6:
        auth = au[0] + " *et al.*"
    elif len(au) > 2:
        auth = ", ".join(au[:-1]) + ", and " + au[-1]
    else:
        auth = " and ".join(au)
    q = title.rstrip(".")
    qt = f"“{q}”" if q.endswith(("?", "!")) else f"“{q},”"
    pp = f"pp. {pages}" if pages and ("-" in pages or "–" in pages) else (f"p. {pages}" if pages else "")
    tail = (f", doi: {doi}." if doi else (f". [Online]. Available: {url}" if url else "."))
    if kind == "book" or (not venue and kind not in ("thesis", "report", "preprint")):
        out = f"*{q}*." + (f" {venue}," if venue else "") + f" {year}" + (f". doi: {doi}." if doi else ".")
        return re.sub(r"\s+", " ", ((auth + ", ") if auth else "") + out).strip()
    if kind == "conference":
        parts = [qt, f"in *{venue}*", year, pp]
    elif kind == "thesis":
        parts = [qt, "Thesis", inst, year]
    elif kind == "report":
        parts = [qt, venue, year]
    elif kind == "preprint":
        parts = [qt, (venue + " (preprint)") if venue else "Preprint", year]
    else:
        parts = [qt, f"*{venue}*", f"vol. {vol}" if vol else "", f"no. {iss}" if iss else "", pp, year]
    first, rest = parts[0], [x for x in parts[1:] if x]
    out = first + (" " + ", ".join(rest) if rest else "")
    return re.sub(r"\s+", " ", ((auth + ", ") if auth else "") + out.rstrip(",") + tail).strip()


def _mark(nums, marks):
    """The numbers of one citation, in the look the person chose: IEEE "[1], [2]" and "[1]–[3]"; "[1,2]", "(1–3)", or
    raised "¹˒²" (for Vancouver and the person's own numbered style)."""
    nums = sorted(set(int(n) for n in nums))
    runs = []
    for n in nums:
        if runs and n == runs[-1][1] + 1:
            runs[-1][1] = n
        else:
            runs.append([n, n])
    if marks == "ieee":
        out = []
        for a, b in runs:
            out += [f"[{a}]"] if a == b else ([f"[{a}]", f"[{b}]"] if b == a + 1 else [f"[{a}]–[{b}]"])
        return ", ".join(out)
    items = []
    for a, b in runs:
        items += [str(a)] if a == b else ([str(a), str(b)] if b == a + 1 else [f"{a}–{b}"])
    inner = ",".join(items)
    if marks == "raised":
        return inner.replace("–", "-").translate(_SUP)
    return f"[{inner}]" if marks == "brackets" else f"({inner})"


def to_numbered(text, sources, marks, given=None):
    """The text's author-date citations (exactly the sources' keys, as keep_known_citations left them) as numbers, in
    order of first citation: given (source index -> number) keeps the numbers an earlier part gave. Returns
    (text, {source index: number})."""
    if not text:
        return text, dict(given or {})
    inners, narr = {}, {}
    for i, s in enumerate(sources or []):
        key = str(s.get("citation_key") or "")
        if key.startswith("(") and key.endswith(")"):
            inner = key[1:-1]
            inners.setdefault(_cite_norm(inner), i)
            m = re.match(r"^(.*?)(?:,\s*|\s+)((?:1[5-9]|20)\d{2}[a-z]?|n\.\s?d\.(?:-[a-z])?|no date)$", inner)
            if m:
                narr.setdefault((_cite_norm(m.group(1)), _cite_norm(m.group(2))), i)
    hits = []          # (start, end, [source indices], kind, extra)
    for m in re.finditer(r"\(([^()]{2,400})\)", text):
        known, other = [], []
        for part in m.group(1).split(";"):
            p = _CITE_LEAD.sub("", part.strip())
            q = _CITE_LOCATOR.sub("", p).rstrip(" ,")
            i = inners.get(_cite_norm(p), inners.get(_cite_norm(q)))
            (known if i is not None else other).append(i if i is not None else part.strip())
        if known:
            hits.append((m.start(), m.end(), known, "paren", other))
    taken = [(a, b) for a, b, _, _, _ in hits]
    for m in _NARRATIVE.finditer(text):
        a = m.end(1)
        b = m.end()
        if any(x <= a < y for x, y in taken):
            continue
        yr = _cite_norm(m.group(2))
        before = _cite_norm(text[max(0, m.start(1) - 300):m.end(1)])
        found = None
        for (nm, yy), i in narr.items():
            if yy == yr and nm and (before == nm or before.endswith(" " + nm) or before.endswith(nm)):
                found = i
                break
        if found is not None:
            paren = text.rfind("(", a, b)
            hits.append((paren, b, [found], "narr", []))
    hits.sort(key=lambda h: h[0])
    nums = dict(given or {})
    nxt = max(nums.values(), default=0) + 1
    for _, _, idx, _, _ in hits:
        for i in idx:
            if i not in nums:
                nums[i] = nxt
                nxt += 1
    out = text
    for a, b, idx, kind, other in sorted(hits, key=lambda h: -h[0]):
        mk = _mark([nums[i] for i in idx], marks)
        if kind == "paren" and other:
            rep = "(" + "; ".join(other) + ")" + ("" if marks == "raised" else " ") + mk
        else:
            rep = mk
        lead = out[:a]
        if marks == "raised":
            lead = lead.rstrip(" \t")
        else:
            lead = lead.rstrip(" \t") + " "
        out = lead + rep + out[b:]
    if marks == "raised":
        # raised numbers sit outside a full stop or a comma, inside a colon or a semicolon (AMA, as most guides teach)
        out = re.sub(r"([⁰¹²³⁴⁵⁶⁷⁸⁹˒⁻]+)([.,])", r"\2\1", out)
    out = re.sub(r"[ \t]+([.,;:])", r"\1", out)
    return out, nums


def numbered_refs(sources, nums, style, marks):
    """The reference list in number order; each with its number's mark (key), its number (n) and its author-date key
    (akey: the next part of a long piece goes on from it)."""
    out = []
    for i, n in sorted(nums.items(), key=lambda x: x[1]):
        s = sources[i]
        out.append({"text": numbered_ref(s, style), "url": s.get("url"), "pdf": s.get("pdf") or None, "doi": s.get("doi") or None,
                    "key": _mark([n], marks), "n": n, "akey": s.get("citation_key") or "", "group": s.get("group") or _key_group(s.get("citation_key")),
                    "sort": None, "after": ""})
    return out


_EXAMPLE_WORDS = set("""a an and the of in on for to at by from with et al ed eds edn edition editor editors vol volume no
number issue iss pp p pages page doi https http www org com net html pdf available online internet cited accessed retrieved
viewed preprint thesis dissertation phd master masters doctoral report rep technical proceedings proc conference conf
journal university press publisher abstract article review book chapter series accessed date unknown nd n d s sd no
january february march april may june july august september october november december jan feb mar apr jun jul aug sep sept
oct nov dec""".split())


def _example_check(text, src):
    """A reference RamzAI laid out like the person's example is kept only when everything in it is the paper's own:
    every number is the year, the volume, the issue, a page, or part of its DOI or link; every word comes from its
    authors, title, journal, institution or link, or is a word references use ("In", "Available", "vol"…); and the
    first author's surname, the year and the title are all there."""
    plain = re.sub(r"[*_]", "", str(text or ""))
    if not plain or len(plain) > 1500:
        return False
    names = [a for a in (_name_parts(x) for x in (src.get("authors") or [])) if a]
    vol, iss, pages = _biblio_parts(src)
    data = " ".join(str(x or "") for x in (src.get("title"), src.get("venue"), src.get("institution"), src.get("doi"), src.get("url"),
                                           src.get("year"), vol, iss, pages) + tuple(src.get("authors") or []))
    asc = lambda x: unicodedata.normalize("NFKD", re.sub(r"[\u2010-\u2015\u2212]", " ", x)).encode("ascii", "ignore").decode().lower()
    low = asc(data)
    words = set(re.findall(r"[a-z]+", low))
    nums = set(re.findall(r"\d+", low))
    pl = asc(plain)
    for n in re.findall(r"\d+", pl):
        # a page range may be written short ("123-8"): a number that ends one of the paper's numbers is the paper's
        if n not in nums and not any(x.endswith(n) for x in nums):
            return False
    for w in re.findall(r"[a-z]+", pl):
        if len(w) > 1 and w not in words and w not in _EXAMPLE_WORDS:
            # initials of the authors' given names are the paper's too
            if not (len(w) <= 3 and all(any(g[:1].lower() == c for a in names if not a.get("org") for g in a.get("given") or []) for c in w)):
                return False
    if names:
        sur = unicodedata.normalize("NFKD", _au_surname(names[0])).encode("ascii", "ignore").decode().lower()
        if re.sub(r"[^a-z]", "", sur) not in re.sub(r"[^a-z]", "", pl):
            return False
    if src.get("year") and str(src["year"]) not in pl:
        return False
    tw = [w for w in re.findall(r"[a-z]{4,}", unicodedata.normalize("NFKD", str(src.get("title") or "")).encode("ascii", "ignore").decode().lower())]
    if tw and sum(1 for w in tw if w in pl) < 0.8 * len(tw):
        return False
    return True


def refs_like_example(entries, example, name):
    """Each source's reference in the person's own style, laid out like their example. entries: [(source, the reference
    in the nearest built-in style)]. Returns a list of (text, kept): kept is False where the check failed and the
    built-in reference stands."""
    if not example or not entries:
        return [(t, False) for _, t in entries]
    data = []
    for src, _ in entries:
        vol, iss, pages = _biblio_parts(src)
        data.append({k: v for k, v in (("authors", src.get("authors") or []), ("year", src.get("year")), ("title", _tidy_title(src.get("title") or "")),
                                        ("journal_or_source", src.get("venue")), ("volume", vol), ("issue", iss), ("pages", pages),
                                        ("doi", src.get("doi")), ("link", src.get("url")), ("kind", src.get("kind") or "article"),
                                        ("institution", src.get("institution"))) if v})
    system = ("You lay out references in a citation style that a person showed you by example. Reply with one JSON object only, "
              'no code fences: {"refs": ["…", "…"]}, one reference for each item, in the same order. Copy the example\'s layout '
              "exactly: the order of the parts, the punctuation, the capitalisation, how authors' names and initials are written, "
              "abbreviations such as vol. or pp., and where italics go (mark italics with *asterisks*). Use ONLY the data given for "
              "each item: never add, guess or change a name, a number, a word of a title, a journal, a place or a publisher. Leave "
              "out a part the data doesn't have. The style is called: " + name[:60])
    user = "THE PERSON'S EXAMPLE:\n<<<\n" + example[:600] + "\n>>>\n\nTHE ITEMS:\n" + json.dumps(data, ensure_ascii=False)
    try:
        raw = call_claude(system, user, max_tokens=2500, feature="cite-style")
        obj = _guide_json(raw) or {}
        got = obj.get("refs") if isinstance(obj.get("refs"), list) else []
    except Exception as e:
        logger.warning(f"Couldn't lay out references like the example: {e}")
        got = []
    out = []
    for k, (src, built) in enumerate(entries):
        t = re.sub(r"\s+", " ", str(got[k] if k < len(got) else "")).strip()
        ok = bool(t) and _example_check(t, src)
        out.append((t if ok else built, ok))
    return out


# 5.0: the reference list's heading in the language the piece is written in (the references themselves stay as published)
REF_TITLES = {"french": "Références", "français": "Références", "francais": "Références", "spanish": "Referencias", "español": "Referencias",
              "espanol": "Referencias", "portuguese": "Referências", "português": "Referências", "portugues": "Referências",
              "german": "Literaturverzeichnis", "deutsch": "Literaturverzeichnis", "arabic": "المراجع", "العربية": "المراجع",
              "kiswahili": "Marejeleo", "swahili": "Marejeleo"}


def ref_title_for(language):
    return REF_TITLES.get(re.sub(r"\s+", " ", str(language or "")).strip().lower())


def restyle_result(result, sources, cited, plan, given=None, language=None):
    """Turns a finished piece (made with author-date citations, all checked) into the style the person asked for.
    result: the reply about to be sent (text or sections, references). cited: the indices of the sources it cites."""
    if not plan:
        if ref_title_for(language):
            result["references_title"] = ref_title_for(language)
        return result
    cited = set(cited or ())
    if plan["marks"]:
        nums = dict(given or {})
        if result.get("sections"):
            for sec in result["sections"]:
                sec["text"], nums = to_numbered(sec.get("text") or "", sources, plan["marks"], nums)
        else:
            result["text"], nums = to_numbered(result.get("text") or "", sources, plan["marks"], nums)
        # a source cited only in an earlier part keeps its number; one cited nowhere isn't listed
        nums = {i: n for i, n in nums.items() if i in cited or (given and i in given)}
        refs = numbered_refs(sources, nums, plan["refs"], plan["marks"])
        result["references"] = refs
        result["references_title"] = "References"
        result["cite_marks"] = plan["marks"]
    else:
        result["cite_marks"] = None
        refs = result.get("references") or []
    result["style"] = plan["style"]
    result["style_name"] = plan["name"]
    if plan["style"] == "other" and plan.get("example") and refs:
        by_key = {str(s.get("citation_key") or ""): s for s in sources}
        entries = [(by_key.get(r.get("akey") or r.get("key"), {}), r["text"]) for r in refs]
        laid = refs_like_example([(src, t) for src, t in entries if src] , plan["example"], plan["name"]) if any(src for src, _ in entries) else []
        k = 0
        failed = 0
        for r, (src, _) in zip(refs, entries):
            if not src:
                continue
            r["text"], ok = laid[k]
            r["after"] = "" if ok else r.get("after", "")
            failed += 0 if ok else 1
            k += 1
        if failed:
            note = (f"{failed} of the {len(refs)} references couldn't be laid out like your example and checked, so "
                    + ("they follow Vancouver." if plan["marks"] else f"they follow {CITE_STYLES[plan['refs']]['name']}."))
            result["source_note"] = ((result.get("source_note") or "") + " " + note).strip()
    elif plan["style"] == "other" and not plan.get("example"):
        note = ("You didn't give an example reference, so the references follow "
                + ("Vancouver" if plan["marks"] else CITE_STYLES[plan["refs"]]["name"]) + ".")
        result["source_note"] = ((result.get("source_note") or "") + " " + note).strip()
    if ref_title_for(language):
        result["references_title"] = ref_title_for(language)
    return result


_MARK_GROUPS = {
    "ieee": re.compile(r"\[\d{1,3}\](?:\s*(?:,|–|-)\s*\[\d{1,3}\])*"),
    "brackets": re.compile(r"\[\d{1,3}(?:\s*[,–\-]\s*\d{1,3})*\]"),
    "parens": re.compile(r"\(\d{1,3}(?:\s*[,–\-]\s*\d{1,3})*\)"),
    "raised": re.compile(r"[⁰¹²³⁴⁵⁶⁷⁸⁹]+(?:[˒⁻][⁰¹²³⁴⁵⁶⁷⁸⁹]+)*"),
}
_SUP_BACK = str.maketrans("⁰¹²³⁴⁵⁶⁷⁸⁹˒⁻", "0123456789,-")


def _mark_numbers(mark):
    """The numbers one mark stands for ("[1]–[3]" and "(1–3)" are 1, 2, 3)."""
    t = mark.translate(_SUP_BACK)
    out = []
    for m in re.finditer(r"(\d{1,3})(?:\]?\s*[–\-]\s*\[?(\d{1,3}))?", t):
        a, b = int(m.group(1)), int(m.group(2)) if m.group(2) else None
        out += list(range(a, b + 1)) if b and b >= a and b - a < 60 else [a]
    return out


def marks_back(text, marks, n_to_key):
    """An earlier part's numbers as the author-date citations they stand for, so the writing goes on from its words (a
    mark with a number that isn't known stays as it is)."""
    rx = _MARK_GROUPS.get(marks)
    if not rx or not text:
        return text

    def back(m):
        keys = [n_to_key.get(n) for n in _mark_numbers(m.group(0))]
        if not keys or not all(keys):
            return m.group(0)
        return " (" + "; ".join(k[1:-1] for k in keys) + ")"
    out = rx.sub(back, text)
    return re.sub(r"[ \t]{2,}", " ", out)


def _strip_invented(text):
    """Numbers in square brackets, or raised, that the model wrote itself: never citations (they'd lead nowhere)."""
    out = _MARK_GROUPS["brackets"].sub("", str(text or ""))
    out = _MARK_GROUPS["raised"].sub("", out)
    return re.sub(r"[ \t]+([.,;:])", r"\1", out)


# 5.0: the language a piece is written in (asked before research starts). Citations stay exactly as given, numbers are
# written in digits (so every figure can still be checked against the papers), and a quotation stays in its own language.
LANGUAGES_RTL = {"arabic", "hebrew", "urdu", "persian", "farsi"}


def language_rule(language):
    lang = re.sub(r"[^\w\s\-’'()]", "", str(language or "")).strip()[:40]
    if not lang or lang.lower() in ("english", "en", "british english", "american english", "us english", "uk english"):
        return "" if not lang or lang.lower() in ("english", "en") else (f"\n\nLANGUAGE: write in {lang}.")
    return (f"\n\nLANGUAGE: write everything you write in {lang}, in the formal register of academic writing in {lang}. Keep every "
            "citation key exactly as given (its names, its \"&\", \"and\" or \"et al.\" and its year are never translated), and always "
            "place citations in brackets after the claim (never as \"Name (Year)\" inside a sentence). Write every number in digits. "
            "A paper's title or a quotation from it stays in its own language.")


@app.after_request
def add_cors_headers(resp):
    origin = request.headers.get("Origin", "")
    if _origin_is_allowed(origin):
        resp.headers["Access-Control-Allow-Origin"] = origin
        resp.headers["Vary"] = "Origin"
    resp.headers["Access-Control-Allow-Methods"] = "POST, GET, DELETE, OPTIONS"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization, X-Device"      # 4.9: X-Device (sent with /api/auth/me since 4.1) was refused by browsers
    resp.headers["Access-Control-Expose-Headers"] = "X-Batch-Summary, X-Docente-Note, X-Size-Info"
    # 3.11: the browser may remember a preflight's answer for two hours (it asked before almost every request)
    if request.method == "OPTIONS":
        resp.headers["Access-Control-Max-Age"] = "7200"
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
    # token '<'") instead of an explanation. RamzAI's monthly allowance returns
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


# ======================= 3.15: moving to acadhome.vercel.app =======================
MOVE_DIR = os.path.join(tempfile.gettempdir(), "ah-moves")
MOVE_PART_MAX = 8 * 1024 * 1024 + 4096    # one locked part (the lock adds 28 bytes)
MOVE_PER_PERSON = 112 * 1024 * 1024       # what one person (or one network, signed out) may have waiting at once
MOVE_ALL_MAX = 640 * 1024 * 1024          # what everyone together may have waiting
MOVE_DISK_SPARE = 1024 * 1024 * 1024      # the disk always keeps this much free (conversions need room too)
MOVE_TTL = 3600                           # an hour


def _move_id_ok(mid):
    return bool(re.fullmatch(r"[A-Za-z0-9_-]{32}", mid or ""))


def _move_owner():
    """Who sent a part, without keeping who: a short hash of the account (or the network, signed out)."""
    return hashlib.sha256(("move|" + _account_or_ip_key()).encode()).hexdigest()[:12]


def _move_sweep():
    """Parts older than an hour go. Returns what's still waiting: (bytes in all, {owner: bytes})."""
    total, per = 0, {}
    try:
        names = os.listdir(MOVE_DIR)
    except FileNotFoundError:
        return 0, per
    now = time.time()
    for n in names:
        p = os.path.join(MOVE_DIR, n)
        try:
            st = os.stat(p)
        except FileNotFoundError:
            continue
        if now - st.st_mtime > MOVE_TTL or (n.endswith(".tmp") and now - st.st_mtime > 600):
            try:
                os.remove(p)
            except FileNotFoundError:
                pass
            continue
        if n.endswith(".tmp"):
            continue
        total += st.st_size
        owner = n.split(".", 1)[-1]
        per[owner] = per.get(owner, 0) + st.st_size
    return total, per


def _move_path(mid):
    try:
        for n in os.listdir(MOVE_DIR):
            if n.startswith(mid + ".") and not n.endswith(".tmp"):
                return os.path.join(MOVE_DIR, n)
    except FileNotFoundError:
        pass
    return None


@app.route("/api/move", methods=["POST", "OPTIONS"])
@limiter.limit("60 per hour", key_func=_account_or_ip_key)
def move_put():
    """One locked part of what a phone saved under an old address (only the phone has the key)."""
    if request.method == "OPTIONS":
        return "", 204
    data = request.get_data(cache=False)
    if not data or len(data) < 29:
        return jsonify({"error": "Nothing came through. Please try again."}), 400
    if len(data) > MOVE_PART_MAX:
        return jsonify({"error": "That part is too big."}), 413
    os.makedirs(MOVE_DIR, exist_ok=True)
    total, per = _move_sweep()
    owner = _move_owner()
    if per.get(owner, 0) + len(data) > MOVE_PER_PERSON:
        return jsonify({"error": "That's more than can move at once. Please try again in an hour."}), 429
    try:
        spare = shutil.disk_usage(MOVE_DIR).free
    except OSError:
        spare = 0
    if total + len(data) > MOVE_ALL_MAX or spare - len(data) < MOVE_DISK_SPARE:
        logger.warning(f"Move: no room for a part of {len(data) // 1024} KB ({total // (1024 * 1024)} MB waiting)")
        return jsonify({"error": "Lots of people are moving right now. Please try again in a few minutes."}), 503
    mid = secrets.token_urlsafe(24)
    tmp = os.path.join(MOVE_DIR, f"{mid}.{owner}.tmp")
    with open(tmp, "wb") as f:
        f.write(data)
    os.replace(tmp, os.path.join(MOVE_DIR, f"{mid}.{owner}"))
    logger.info(f"Move: a part of {len(data) // 1024} KB is waiting")
    return jsonify({"id": mid})


@app.route("/api/move/<mid>", methods=["GET", "DELETE", "OPTIONS"])
@limiter.limit("240 per hour")
def move_get(mid):
    if request.method == "OPTIONS":
        return "", 204
    if not _move_id_ok(mid):
        return jsonify({"error": "This moving link isn't right."}), 404
    path = _move_path(mid)
    if request.method == "DELETE":
        if path:
            try:
                os.remove(path)
            except FileNotFoundError:
                pass
        return jsonify({"ok": True})
    try:
        fresh = path and time.time() - os.stat(path).st_mtime <= MOVE_TTL
    except FileNotFoundError:
        fresh = False
    if not fresh:
        return jsonify({"error": "This moving link has expired or was used already."}), 404
    try:
        with open(path, "rb") as f:
            data = f.read()
    except FileNotFoundError:
        return jsonify({"error": "This moving link has expired or was used already."}), 404
    resp = app.response_class(data, mimetype="application/octet-stream")
    resp.headers["Cache-Control"] = "no-store"
    return resp


@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "email": _email_ready(),
                    "promo_until": _promo_until().isoformat() if _promo_active() else None})


@app.route("/api/promo", methods=["GET"])
def promo_status():
    """For the app's banner — no account needed."""
    return jsonify({"active": _promo_active(), "until": _promo_until().isoformat() if _promo_active() else None})


# ---------- 3.15: study music: songs by mood ----------
# The list is songs.json beside this file: edit it on GitHub to change the songs (the app picks the new list up within a
# day; a broken file just means the app keeps the list it ships with). Each song: v (the YouTube video's id), t, a, m.
_SONGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "songs.json")
_SONG_ID = re.compile(r"^[A-Za-z0-9_-]{11}$")
_SONG_LIST = re.compile(r"^[A-Za-z0-9_-]{10,64}$")
_SONG_MOOD = re.compile(r"^[a-z]{2,20}$")
_songs_cache = {"mtime": None, "data": None}
_song_about = {}


def _songs_load():
    try:
        mtime = os.stat(_SONGS_FILE).st_mtime
    except OSError:
        return None
    if _songs_cache["mtime"] == mtime:
        return _songs_cache["data"]
    data = None
    try:
        with open(_SONGS_FILE, encoding="utf-8") as f:
            raw = json.load(f)
        moods = [{"id": m["id"], "name": str(m["name"])[:24], "lead": str(m["lead"])[:140]}
                 for m in raw.get("moods", []) if isinstance(m, dict) and _SONG_MOOD.match(str(m.get("id", "")))
                 and m.get("id") != "mine" and m.get("name") and m.get("lead")][:12]
        ids = {m["id"] for m in moods}
        songs, seen = [], set()
        for s in raw.get("songs", []):
            if not isinstance(s, dict) or not _SONG_ID.match(str(s.get("v", ""))) or s.get("m") not in ids or not s.get("t"):
                continue
            if s["v"] in seen:
                continue
            seen.add(s["v"])
            one = {"v": s["v"], "t": str(s["t"])[:120], "a": str(s.get("a", ""))[:80], "m": s["m"]}
            if s.get("live"):
                one["live"] = True
            songs.append(one)
        if len(moods) >= 2 and len(songs) >= 6:
            data = {"updated": str(raw.get("updated", ""))[:20], "moods": moods, "songs": songs[:400]}
    except (OSError, ValueError, TypeError, KeyError, AttributeError):
        logging.getLogger(__name__).warning("songs.json could not be read; the app keeps its own list")
    _songs_cache.update(mtime=mtime, data=data)
    return data


@app.route("/api/music/songs", methods=["GET", "OPTIONS"])
def music_songs():
    """The songs for Study music, by mood. No account needed."""
    if request.method == "OPTIONS":
        return "", 204
    data = _songs_load()
    if not data:
        return jsonify({"error": "No list of songs here."}), 404
    resp = jsonify(data)
    resp.headers["Cache-Control"] = "public, max-age=3600"
    return resp


@app.route("/api/music/about", methods=["GET", "OPTIONS"])
@limiter.limit("60 per hour")
def music_about():
    """A YouTube song's or playlist's name, for Your songs: YouTube's oEmbed is asked with only the link."""
    if request.method == "OPTIONS":
        return "", 204
    v, lst = request.args.get("v", ""), request.args.get("list", "")
    if v and _SONG_ID.match(v):
        url, key = "https://www.youtube.com/watch?v=" + v, "v:" + v
    elif lst and _SONG_LIST.match(lst):
        url, key = "https://www.youtube.com/playlist?list=" + lst, "l:" + lst
    else:
        return jsonify({"ok": False, "reason": "bad_link"}), 400
    hit = _song_about.get(key)
    if hit and time.time() - hit[0] < 86400:
        return jsonify(hit[1])
    try:
        r = requests.get("https://www.youtube.com/oembed", params={"url": url, "format": "json"}, timeout=6)
    except requests.RequestException:
        return jsonify({"ok": False, "reason": "unreachable"}), 502
    if r.status_code == 200:
        try:
            d = r.json()
        except ValueError:
            d = {}
        out = {"ok": True, "title": str(d.get("title", ""))[:140], "author": str(d.get("author_name", ""))[:80]}
    elif r.status_code in (401, 403):
        out = {"ok": False, "reason": "not_embeddable"}     # its owner doesn't let it play in other apps (or it's private)
    elif r.status_code in (400, 404):
        out = {"ok": False, "reason": "not_found"}
    else:
        return jsonify({"ok": False, "reason": "unreachable"}), 502
    if len(_song_about) > 2000:
        _song_about.clear()
    _song_about[key] = (time.time(), out)
    return jsonify(out)


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
            cur.execute("SELECT id, name, token_version FROM users WHERE email = %s", (email,))
            row = cur.fetchone()
    finally:
        conn.close()

    # Always the same response whether or not the email has an account —
    # never confirm or deny that in the response itself.
    if row:
        token = make_reset_token(row["id"], email, row.get("token_version") or 0)
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
            logger.warning(f"Couldn't send a password reset email: {e.__class__.__name__}")
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
    link_tv = _reset_token_version(token)

    conn = get_db()
    try:
        with conn.cursor() as cur:
            if link_tv is not None:
                cur.execute("SELECT token_version FROM users WHERE id = %s", (user_id,))
                cur_row = cur.fetchone()
                if cur_row and (cur_row.get("token_version") or 0) != link_tv:
                    return jsonify({"error": "This reset link has already been used — request a new one"}), 400
            # Receiving the link proves the address it was sent to: if that's still the account's
            # address, it counts as confirmed (and a Google link made before it was confirmed is
            # dropped). Every other signed-in device is signed out.
            cur.execute(
                """UPDATE users SET password_hash = %s, token_version = token_version + 1,
                       google_sub = CASE WHEN email_verified OR lower(email) <> %s THEN google_sub ELSE NULL END,
                       email_verified = email_verified OR lower(email) = %s,
                       email_proven = email_proven OR lower(email) = %s
                   WHERE id = %s AND (%s IS NULL OR token_version = %s)""",
                (generate_password_hash(new_password), link_email or "", link_email or "", link_email or "", user_id, link_tv, link_tv),
            )
            if cur.rowcount == 0 and link_tv is not None:
                return jsonify({"error": "This reset link has already been used — request a new one"}), 400
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


_DEVICE_RE = re.compile(r"^[A-Za-z0-9_-]{16,80}$")


def _device_seen(row, device_id):
    """The free trial is once per phone: the first account on a phone keeps it; another account on the same phone
    gets none. Returns the row (updated when the trial was just closed)."""
    if not device_id or not _DEVICE_RE.match(device_id) or not DATABASE_URL:
        return row
    try:
        conn = get_db()
        try:
            with conn.cursor() as cur:
                cur.execute("SELECT user_id FROM trial_devices WHERE device_id = %s", (device_id,))
                seen = cur.fetchone()
                if not seen:
                    cur.execute("INSERT INTO trial_devices (device_id, user_id) VALUES (%s, %s) ON CONFLICT (device_id) DO NOTHING", (device_id, row["id"]))
                elif seen["user_id"] != row["id"] and not row.get("trial_blocked") and effective_plan(row) == "free" and not (row.get("trial_points_used") or 0):
                    cur.execute("UPDATE users SET trial_blocked = TRUE WHERE id = %s RETURNING *", (row["id"],))
                    updated = cur.fetchone()
                    if updated:
                        row = _with_promo(updated)
                        logger.info(f"Trial closed for user {row['id']}: the phone already had its trial")
            conn.commit()
        finally:
            conn.close()
    except Exception as e:
        logger.warning(f"Couldn't record the device for user {row.get('id')}: {e}")
    return row


@app.route("/api/auth/me", methods=["GET", "OPTIONS"])
@auth_required
def me():
    row = _device_seen(request.current_user, request.headers.get("X-Device", ""))
    row = _invite_joined(row, request.headers.get("X-Device", ""))
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
                # One full use of Pro: RamzAI messages that never expire, on top
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
# On the web (acadhome.vercel.app) Pro is bought with Paystack — card or
# mobile money, charged in Ghana cedis — as a one-time pass: 30 or 365 days of
# Pro, or 30 RamzAI messages. The App Store and Google Play apps must use the
# stores' own purchases (their rules), so requests from the store apps are
# refused here and the app hides Paystack there too. The secret key is read
# only from the server's environment — never from code or the app.
PAYSTACK_SECRET_KEY = os.environ.get("PAYSTACK_SECRET_KEY", "").strip()
PAYSTACK_PUBLIC_KEY = os.environ.get("PAYSTACK_PUBLIC_KEY", "").strip()
PAYSTACK_API = "https://api.paystack.co"
PAYSTACK_PLANS = {
    # plan: (price in pesewas, days of Scholar, points). A pass's points end with the pass; "points" never end.
    # 4.7 (the owner, 9 Nov 2026): Month GH₵160 = 200 points; Half a year GH₵480 = 620; Year GH₵1,200 = 1,560
    # 5.0 (the owner, 10 Nov 2026): Month = 100 points, 10 of them for research; Half year "stays 620" (his answer at
    # 10:05), 60 of them for research (six months of Month's 10); Year unchanged, with no research limit
    # (PASS_RESEARCH_POINTS)
    "month": (16000, 30, 100),
    "half": (48000, 182, 620),
    "year": (120000, 365, 1560),
    # 4.1 to 4.6: no longer sold, kept so a pass bought then still applies
    "day": (2000, 1, 40),
    "week": (6000, 7, 150),
    "semester": (48000, 120, 1400),
    "points": (4000, 0, 100),
    # earlier plans: no longer sold, kept so an older payment still applies
    "monthly": (10000, 30, 400),
    "halfyear": (40000, 182, 1600),
    "payperuse": (3000, 0, 30),
    "yearly": (30000, 365, 3000),
}
PAYSTACK_ON_SALE = ("month", "half", "year")
# a payment started before 4.7 has no points of its own written down: it gets what was offered then
PAYSTACK_POINTS_BEFORE_47 = {"month": 450, "year": 3500}


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
            if pay.get("points") is not None:
                credits = int(pay["points"])                 # what was offered when the payment started
            elif pay["plan"] in PAYSTACK_POINTS_BEFORE_47:
                credits = PAYSTACK_POINTS_BEFORE_47[pay["plan"]]   # started before 4.7
            # 5.0: the research points this pass came with (None: no research limit — Year, or a payment started before 5.0)
            research = pay.get("research_points")
            if days:
                # Bought during the free-Pro promotion: the paid days start when it ends.
                start = _promo_until() if _promo_active() else datetime.now(timezone.utc)
                # A shorter pass bought on top of a longer one that is still running keeps the longer one's name;
                # the days and points are added either way.
                cur.execute("SELECT plan, plan_expires_at, research_points FROM users WHERE id = %s FOR UPDATE", (pay["user_id"],))
                cur_row = cur.fetchone() or {}
                new_plan = pay["plan"]
                cur_plan, cur_exp = cur_row.get("plan"), cur_row.get("plan_expires_at")
                running = bool(cur_plan in PAYSTACK_PLANS and cur_exp and cur_exp > datetime.now(timezone.utc))
                if running and PAYSTACK_PLANS[cur_plan][1] > days:
                    new_plan = cur_plan
                # research points are added like the points themselves; no limit wins (this pass has none, or the pass
                # still running has none), and a pass that had already ended leaves nothing behind
                if research is None or (running and cur_row.get("research_points") is None):
                    new_research = None
                else:
                    new_research = (max(0, int(cur_row.get("research_points") or 0)) if running else 0) + int(research)
                cur.execute(
                    "UPDATE users SET plan = %s, plan_expires_at = GREATEST(COALESCE(plan_expires_at, now()), now(), %s::timestamptz) + make_interval(days => %s), "
                    "plan_source = 'paystack', pass_points = pass_points + %s, research_points = %s WHERE id = %s",
                    (new_plan, start, days, credits, new_research, pay["user_id"]))
                if pay["plan"] in ("month", "half", "semester", "year", "monthly", "halfyear", "yearly") and _reward_referral(cur, pay["user_id"], pay.get("referral_code")):
                    cur.execute("UPDATE paystack_payments SET referral_applied = TRUE WHERE reference = %s", (reference,))
            elif credits:
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
        "plans": {k: {"amount_ghs": v[0] / 100, "days": v[1], "credits": v[2], "points": v[2], "research_points": PASS_RESEARCH_POINTS.get(k),
                      "name": PLAN_NAMES.get(k, "Points")} for k, v in PAYSTACK_PLANS.items() if k in PAYSTACK_ON_SALE},
        "trial_points": TRIAL_POINTS, "day_cap": SCHOLAR_DAY_CAP,
        "free_points": FREE_MONTH_POINTS, "free_research": FREE_MONTH_RESEARCH,
    })


@app.route("/api/paystack/initialize", methods=["POST", "OPTIONS"])
@limiter.limit("20 per hour", key_func=_account_or_ip_key)
@auth_required
@_once
def paystack_initialize():
    if not _paystack_on():
        return jsonify({"error": "Payments aren't set up yet — please try again later."}), 503
    if _from_store_app():
        return jsonify({"error": "In the Academic-Home app from a store, Scholar is bought through the App Store or Google Play."}), 403
    plan = str((request.get_json(silent=True) or {}).get("plan") or "")
    if plan not in PAYSTACK_ON_SALE:
        return jsonify({"error": "Choose a Scholar pass: Month, Half year or Year."}), 400
    user = request.current_user
    if plan in PRO_PLANS and user.get("plan") in PRO_PLANS and _owns_store_plan(user["id"]):
        return jsonify({"error": "You already have Scholar through the App Store or Google Play; manage it there."}), 409
    friend_code = _clean_code((request.get_json(silent=True) or {}).get("referral_code"))
    amount = PAYSTACK_PLANS[plan][0]
    reference = f"dct_{user['id']}_{secrets.token_hex(6)}"
    conn = get_db()
    try:
        with conn.cursor() as cur:
            if friend_code and plan in PRO_PLANS:
                if not _find_referrer(cur, friend_code, user["id"]):
                    return jsonify({"error": "That referral code wasn't found (or it's your own) — check it, or leave it empty.", "field": "referral"}), 400
            cur.execute("INSERT INTO paystack_payments (reference, user_id, plan, amount, referral_code, points, research_points) VALUES (%s, %s, %s, %s, %s, %s, %s)",
                        (reference, user["id"], plan, amount, friend_code or None, PAYSTACK_PLANS[plan][2], PASS_RESEARCH_POINTS.get(plan)))
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


# ------------------------------------------------------------ RamzAI's memory
# What RamzAI remembers about a person is theirs to see and control: list it,
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
                    return jsonify({"error": "Write something for RamzAI to remember."}), 400
                cur.execute("SELECT note FROM user_memory WHERE user_id = %s", (user["id"],))
                existing = [r["note"].strip().lower() for r in cur.fetchall()]
                if len(existing) >= MEMORY_LIMIT:
                    return jsonify({"error": f"RamzAI can remember up to {MEMORY_LIMIT} notes — delete one first."}), 400
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
# friend joins with it, the person who invited gets INVITE_POINTS points (5.0;
# before: both got a bonus month when the friend subscribed) — the same reward as
# a personal referral code. Kept in promo_codes (created_by = the referrer).
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
            uid = request.current_user["id"]
            state = _referral_state(_latest_referral(cur, uid))
            # 5.0: the friends who have counted, and what an invite gives
            cur.execute("SELECT count(*) AS n, COALESCE(sum(points), 0) AS p, count(*) FILTER (WHERE created_at >= date_trunc('month', now())) AS m "
                        "FROM invite_rewards WHERE inviter_id = %s", (uid,))
            r = cur.fetchone() or {}
            state.update({"invited": int(r.get("n") or 0), "invited_points": int(r.get("p") or 0), "invited_this_month": int(r.get("m") or 0),
                          "invite": {"points": INVITE_POINTS, "research": INVITE_RESEARCH, "month_max": INVITE_MONTH_MAX}})
            return jsonify(state)
    finally:
        conn.close()


def _wrong_code(resp):
    """5.0 (the owner: "Protect users' invite code"): only wrong codes count toward the stricter limit, so guessing other
    people's codes is stopped after a few tries while a person typing their friend's code is never held up."""
    try:
        j = resp.get_json(silent=True) or {}
        return j.get("valid") is False or j.get("field") == "code"
    except Exception:
        return False


def _wrong_code_limit():
    """Signed in: 12 wrong codes an hour for the account. Signed out (the sign-up page), the count is per network, and many
    people share one network (a school, a mobile provider), so it is higher there."""
    return "12 per hour" if str(_account_or_ip_key()).startswith("user:") else "30 per hour"


@app.route("/api/referral/check", methods=["POST", "OPTIONS"])
@limiter.limit("40 per hour")
@limiter.limit(_wrong_code_limit, deduct_when=_wrong_code, key_func=_account_or_ip_key)
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
                    return jsonify({"valid": False, "error": "A friend’s code has already counted for this account. It counts once."})
    finally:
        conn.close()
    first = (str(ref.get("name") or "").strip().split() or ["a friend"])[0][:30]
    return jsonify({"valid": True, "from": first})


@app.route("/api/referral/use", methods=["POST", "OPTIONS"])
@limiter.limit("10 per hour", key_func=_account_or_ip_key)
@limiter.limit("6 per hour", deduct_when=_wrong_code, key_func=_account_or_ip_key)
@auth_required
def referral_use():
    """An account that already exists enters the code of the friend who invited it — once, with the same care as a
    new one: its email is confirmed (auth_required), it is the first account on this phone, and it isn't its own code.
    Nobody has to sign up again for a friend to get their points."""
    user = request.current_user
    code = _clean_code((request.get_json(silent=True) or {}).get("code"))
    if len(code) < 4:
        return jsonify({"error": "Enter the whole code.", "field": "code"}), 400
    device = request.headers.get("X-Device", "")
    if not device or not _DEVICE_RE.match(device):
        return jsonify({"error": "Open Academic-Home on your phone and try again."}), 400
    user = _device_seen(user, device)
    conn = get_db()
    conn.autocommit = False
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT id, referral_code, referral_rewarded FROM users WHERE id = %s FOR UPDATE", (user["id"],))
            me_row = cur.fetchone()
            if not me_row:
                conn.rollback()
                return jsonify({"error": "Not authenticated"}), 401
            if me_row["referral_rewarded"]:
                conn.rollback()
                return jsonify({"error": "A friend’s code has already counted for this account. It counts once."}), 409
            if _clean_code(me_row.get("referral_code")) == code:
                conn.rollback()
                return jsonify({"error": "That’s your own code. Share it with friends instead.", "field": "code"}), 400
            ref = _find_referrer(cur, code, user["id"])
            if not ref:
                conn.rollback()
                return jsonify({"error": "That code wasn’t found. Check the spelling, or ask your friend for it again.", "field": "code"}), 400
            cur.execute("SELECT user_id FROM trial_devices WHERE device_id = %s", (device,))
            seen = cur.fetchone()
            if not seen or seen["user_id"] != user["id"] or user.get("trial_blocked"):
                conn.rollback()
                return jsonify({"error": "This phone already had another account before this one, so a friend’s code can’t count here."}), 409
            cur.execute("UPDATE users SET referred_code = %s WHERE id = %s", (code, user["id"]))
            given = _reward_referral(cur, user["id"], code)
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
    first = (str(ref.get("name") or "").strip().split() or ["Your friend"])[0][:30]
    # (when this month's friends have already counted, this one counts next month, when this account next opens the app)
    return jsonify({"ok": True, "from": first, "given": bool(given), "points": INVITE_POINTS})


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
            try:
                if _once_logged(cur, row["id"]):      # 4.9: this conversion was sent again: it's in the history already
                    conn.commit()
                    return
            except Exception as e:
                logger.warning(f"once (conversion history): {e}")
                conn.rollback()
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


def _refund_sage_message(user_row, source="allowance", points=1, research=False, in_research=False):
    """Gives back what a request spent when it never got a reply: pass points, bought points, Free's points (and,
    for research writing, that month's research writing) or a message of the old monthly allowance; and, on a pass
    with research points (5.0), the research points it took. Best-effort."""
    points = max(1, int(points or 1))
    conn = None
    try:
        conn = get_db()
        conn.autocommit = False
        with conn.cursor() as cur:
            if source == "credit":
                cur.execute("UPDATE users SET remy_credits = remy_credits + %s WHERE id = %s", (points, user_row["id"]))
            elif source == "pass":
                cur.execute("UPDATE users SET pass_points = pass_points + %s, day_points_used = GREATEST(day_points_used - %s, 0) WHERE id = %s", (points, points, user_row["id"]))
            elif source == "trial":
                cur.execute("UPDATE users SET trial_points_used = GREATEST(trial_points_used - %s, 0) WHERE id = %s", (points, user_row["id"]))
            else:
                cur.execute("UPDATE users SET sage_messages_used = GREATEST(sage_messages_used - 1, 0) WHERE id = %s", (user_row["id"],))
            # a research writing that never arrived doesn't count against the month's research writings: on Free, and
            # (4.7) on a pass that isn't Half a year or Year
            plan = effective_plan(user_row)
            bonus = max(0, min(points, int(getattr(source, "bonus", 0) or 0)))     # research points that never end, taken by this spend
            if research and not bonus and ((plan == "free" and source in ("trial", "credit")) or (plan != "free" and plan not in RESEARCH_PLANS)):
                cur.execute("UPDATE users SET free_research_used = GREATEST(free_research_used - 1, 0) WHERE id = %s", (user_row["id"],))
            # 5.0: research that never arrived gives its research points back where they came from: the pass's own (a pass
            # without a research limit has none written down, and stays that way), and the ones that never end
            if bonus:
                cur.execute("UPDATE users SET bonus_research = bonus_research + %s WHERE id = %s", (bonus, user_row["id"]))
            if (research or in_research) and plan != "free" and source in ("pass", "credit") and points - bonus > 0:
                cur.execute("UPDATE users SET research_points = research_points + %s WHERE id = %s AND research_points IS NOT NULL", (points - bonus, user_row["id"]))
        conn.commit()
    except Exception as e:
        logger.error("Couldn't refund points for user %s: %s", user_row.get("id"), e)
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


class PointsError(ConversionError):
    """The account has no points for this piece of AI work; the app offers Scholar or points."""
    extra = {"points_needed": True}


def _points_reply(e):
    return jsonify(dict({"error": str(e)}, **getattr(e, "extra", {}))), 429


def _no_points_error(row, points):
    plan = effective_plan(row)
    month = PAYSTACK_PLANS["month"]
    offer = f"Scholar Month is GH₵{month[0] // 100:.0f} for {month[2]} points"
    if plan == "free":
        back = _free_back_words()
        if row.get("trial_blocked"):
            return PointsError(f"This phone’s free points belong to another account. {offer}.")
        return PointsError(f"You’ve used this month’s {FREE_MONTH_POINTS} free points. They come back on {back}. {offer}.")
    return PointsError(f"Your pass has no points left. A new pass adds its points on top: {offer}.")


def _research_used_error():
    return PointsError(f"You’ve used this month’s {FREE_MONTH_RESEARCH} research writings. They come back on "
                       f"{_free_back_words()}. {SCHOLAR_MORE_RESEARCH}")


class ResearchPointsError(PointsError):
    """A Month or Half year pass has used the part of its points that may go on research (5.0). The rest of its
    points are untouched; the app says so and doesn't offer points, only the plans."""
    extra = {"points_needed": True, "research_limit": True}


def _research_points_error(left, points):
    left = max(0, int(left or 0))
    rest = "Your other points still work on everything else in RamzAI. The Year pass has no research limit."
    if left <= 0:
        e = ResearchPointsError("This pass’s research points are used. " + rest)
    else:
        e = ResearchPointsError(f"This needs {points} research points, and you have {left} left. " + rest)
    e.extra = dict(ResearchPointsError.extra, research_points_left=left)
    return e


def _free_back_words():
    d = _next_month_start()
    return f"{d.day} {d.strftime('%B')}"


def _log_sage_message_and_consume(user_row, points=1, research=False, in_research=False):
    """Spend points for one piece of AI work, under a row lock so two requests at once can't both slip past.
    A pass's points first, then points bought on their own. On Free: the month's free points first, then bought
    points, and research writing at most FREE_MONTH_RESEARCH times a month. Any pass spends at most SCHOLAR_DAY_CAP
    a day. On a pass with research points (5.0: Month, Half year), work inside a research feature (research=True, or
    in_research=True for the plan of a manual research writing, which is not one of Free's research writings) may use
    at most those; every other feature is not limited by them. Store and promo accounts keep the earlier monthly
    allowance. Returns where the points came from: 'pass', 'credit', 'trial' (Free's month) or 'allowance'."""
    points = max(1, int(points or 1))
    conn = get_db()
    conn.autocommit = False
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM users WHERE id = %s FOR UPDATE", (user_row["id"],))
            row = _with_promo(cur.fetchone())
            plan = effective_plan(row)
            old_way = plan != "free" and row.get("plan_source") in ("store", "promo") and not (row.get("pass_points") or 0)
            month = _month_start()
            same_month = row.get("free_month") == month
            # 4.7: on a pass that isn't Half a year or Year, research writing counts against the month's few, as on Free
            counted = bool(research and plan != "free" and plan not in RESEARCH_PLANS)
            if counted and _free_state(row)[1] <= 0:
                conn.rollback()
                raise _research_used_error()
            if counted:
                cur.execute("UPDATE users SET free_research_used = %s, free_month = %s WHERE id = %s",
                            ((int(row.get("free_research_used") or 0) if same_month else 0) + 1, month, row["id"]))
                if not same_month:
                    cur.execute("UPDATE users SET trial_points_used = 0 WHERE id = %s", (row["id"],))
            if not old_way:
                bought = int(row.get("remy_credits") or 0)
                if plan != "free":
                    today = datetime.now(timezone.utc).date()
                    used_today = int(row.get("day_points_used") or 0) if row.get("day_points_at") == today else 0
                    if used_today + points > SCHOLAR_DAY_CAP:
                        conn.rollback()
                        raise PointsError(f"That's today's limit ({SCHOLAR_DAY_CAP} points). It starts again tomorrow. "
                                              "Everything that runs on your phone keeps working.")
                    pp = int(row.get("pass_points") or 0)
                    if pp + bought < points:
                        conn.rollback()
                        raise _no_points_error(row, points)
                    # 5.0: research uses the pass's research points (Month, Half year), then the research points that
                    # never end (from inviting friends); None = no research limit
                    r_own = _research_left(row, own=True) if (research or in_research) else None
                    from_bonus = 0
                    if r_own is not None:
                        if r_own + _bonus_research(row) < points:
                            conn.rollback()
                            raise _research_points_error(r_own + _bonus_research(row), points)
                        from_bonus = points - min(r_own, points)
                    from_pass = min(pp, points)
                    cur.execute("UPDATE users SET pass_points = pass_points - %s, remy_credits = remy_credits - %s, day_points_used = %s, day_points_at = %s WHERE id = %s",
                                (from_pass, points - from_pass, used_today + points, today, row["id"]))
                    if r_own is not None:
                        cur.execute("UPDATE users SET research_points = %s, bonus_research = bonus_research - %s WHERE id = %s",
                                    (r_own - (points - from_bonus), from_bonus, row["id"]))
                    conn.commit()
                    charged = _Charged("pass" if from_pass else "credit")
                    charged.bonus = from_bonus
                    return charged
                trial_left, research_left, _ = _free_state(row)
                month = _month_start()
                same = row.get("free_month") == month
                # 5.0: when the month's research writings are used, research points that never end (from inviting
                # friends) still pay for research
                from_bonus = 0
                if research and research_left <= 0:
                    if _bonus_research(row) < points:
                        conn.rollback()
                        raise PointsError(f"You’ve used this month’s {FREE_MONTH_RESEARCH} free research writings. They come back on "
                                          f"{_free_back_words()}. {SCHOLAR_MORE_RESEARCH}")
                    from_bonus = points
                if trial_left + bought < points:
                    conn.rollback()
                    raise _no_points_error(row, points)
                from_trial = min(trial_left, points)
                used = (int(row.get("trial_points_used") or 0) if same else 0) + from_trial
                research_used = (int(row.get("free_research_used") or 0) if same else 0) + (1 if (research and not from_bonus) else 0)
                cur.execute("UPDATE users SET trial_points_used = %s, free_research_used = %s, free_month = %s, "
                            "remy_credits = remy_credits - %s, bonus_research = bonus_research - %s WHERE id = %s",
                            (used, research_used, month, points - from_trial, from_bonus, row["id"]))
                conn.commit()
                charged = _Charged("trial" if from_trial else "credit")
                charged.bonus = from_bonus
                return charged
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
                            f"You've used your {SAGE_FREE_MESSAGES_LIMIT} free messages with RamzAI this month. "
                            f"A Scholar pass adds points for RamzAI and every other AI feature."
                        )
                    raise ConversionError(
                        f"You've reached RamzAI's monthly limit ({SAGE_PAID_MESSAGES_LIMIT} messages) for this "
                        f"account. It resets next month, or a Scholar pass adds points now."
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
    # 3.15: a file's ending by any of its names (.jpeg is a JPG, .htm a web page, .tif a TIFF)
    ext, from_fmt, to_fmt = _any.ALIASES.get(ext, ext), _any.ALIASES.get(from_fmt, from_fmt), _any.ALIASES.get(to_fmt, to_fmt)
    # 3.13: every kind of PowerPoint (.ppt, .pps, .ppsx, .pptm, .pot, .potx, .odp…) is a PowerPoint here, whichever it's
    # called (presenting in a meeting says .pps); a kind with its own way (.ppt or .odp into another kind) goes that way
    own = ext == from_fmt and ((from_fmt, to_fmt) in _OLD_PAIRS or _any.can(from_fmt, to_fmt))
    ppt_in = ext in PPT_EXTS and from_fmt in PPT_EXTS and (from_fmt == "pptx" or not own)
    pair_from = "pptx" if ppt_in else from_fmt
    if (pair_from, to_fmt) not in _OLD_PAIRS and not _any.can(pair_from, to_fmt):
        return jsonify({"error": "Those two kinds of file can't be converted into each other here."}), 400
    if ext != from_fmt and not ppt_in:
        return jsonify({"error": f"File extension .{ext} doesn't match declared source format '{from_fmt}'"}), 400
    if ppt_in:
        from_fmt = "pptx"

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
            if ppt_in:
                src_path = _pptx_for(src_path, to_fmt, work_dir)
            base_name = safe_download_name(file.filename).rsplit(".", 1)[0]
            if studio_template:
                deck = parse_pptx_to_deck(src_path)
                deck["template"] = studio_template
                result_path = build_template_deck(deck, work_dir, "fixed.pptx")
            elif (from_fmt, to_fmt) in _OLD_PAIRS:
                result_path = convert(src_path, from_fmt, to_fmt, work_dir, style=style, pptx_options=pptx_options)
            else:
                result_path = _any.convert_any(src_path, from_fmt, to_fmt, os.path.join(work_dir, "out"), stem=base_name)
        except (ConversionError, StudioError) as e:
            return jsonify({"error": str(e)}), 422
        except FileNotFoundError as e:
            return jsonify({"error": f"Required conversion tool missing on server: {e}"}), 500

        if to_fmt == "pptx" and result_path.endswith(".pptx"):
            try:
                _colour_not_bold(result_path)
            except Exception as e:   # a finished deck is never lost to this tidy-up
                logger.warning(f"Colour-not-bold pass skipped: {e}")

        zipped = result_path.lower().endswith(".zip") and to_fmt != "zip"
        download_name = f"{base_name}.{'zip' if zipped else to_fmt}"

        return send_file(
            result_path,
            mimetype="application/zip" if zipped else (MIME_TYPES.get(to_fmt) or _any.MIME.get(to_fmt) or "application/octet-stream"),
            as_attachment=True,
            download_name=download_name,
        )
    except Exception as e:
        logger.error(f"Conversion failed unexpectedly: {e}", exc_info=True)
        return jsonify({"error": "Conversion failed. Please try again."}), 500
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


def _pptx_for(src_path, to_fmt, work_dir):
    """3.13: a PowerPoint of any kind, ready for a conversion. A PDF is drawn by LibreOffice straight from the file as it
    is (it reads them all); everything else is made from a plain .pptx."""
    kind = presentation_kind(src_path)
    if to_fmt == "pdf" and kind in ("binary", "odp", "strict", "relabel"):
        return src_path
    return as_pptx(src_path, work_dir)


@app.route("/api/convert/formats", methods=["GET", "OPTIONS"])
def convert_formats_endpoint():
    """3.15: every format Convert takes, and what each turns into."""
    return jsonify(_any.catalog())


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
    if ext not in ("pdf", "docx", "xlsx", "doc", "xls", "odt", "rtf", "ods", "csv", "txt", "html") + PPT_EXTS:   # 3.15: more kinds
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
    from_fmt, to_fmt = _any.ALIASES.get(from_fmt, from_fmt), _any.ALIASES.get(to_fmt, to_fmt)      # 3.15
    ppt_batch = from_fmt in PPT_EXTS          # 3.13: any kind of PowerPoint, each read as a .pptx
    if ppt_batch:
        from_fmt = "pptx"
    if (from_fmt, to_fmt) not in _OLD_PAIRS and not _any.can(from_fmt, to_fmt):
        return jsonify({"error": "Those two kinds of file can't be converted into each other here."}), 400
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
            ext = _any.ALIASES.get(ext, ext)
            if ext != from_fmt and not (ppt_batch and ext in PPT_EXTS):
                details.append({"filename": filename, "success": False, "error": "Not a PowerPoint file" if ppt_batch else f"Not a .{from_fmt} file"})
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
                if ppt_batch:
                    src_path = _pptx_for(src_path, to_fmt, file_dir)
                if studio_template:
                    deck = parse_pptx_to_deck(src_path)
                    deck["template"] = studio_template
                    result_path = build_template_deck(deck, file_dir, "fixed.pptx")
                else:
                    if (from_fmt, to_fmt) in _OLD_PAIRS:
                        result_path = convert(src_path, from_fmt, to_fmt, file_dir, style=style, pptx_options=pptx_options)
                    else:
                        result_path = _any.convert_any(src_path, from_fmt, to_fmt, os.path.join(file_dir, "out"),
                                                       stem=safe_download_name(filename).rsplit(".", 1)[0])
                if to_fmt == "pptx":
                    try:
                        _colour_not_bold(result_path)
                    except Exception as e:   # a finished deck is never lost to this tidy-up
                        logger.warning(f"Colour-not-bold pass skipped: {e}")
                base_name = safe_download_name(filename).rsplit(".", 1)[0]
                converted.append((result_path, f"{base_name}.{'zip' if result_path.lower().endswith('.zip') else to_fmt}"))
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


def _scholar_gate(user, feature, research=False):
    """A JSON answer (402) when a Scholar feature is asked for and the account can't have it now; None when it may go
    on. Research writing (research=True): as often as the points allow on Half a year and Year; on Free, Month and the
    other passes while the month's research writings last (4.7)."""
    month = PAYSTACK_PLANS["month"][0] // 100
    if research:
        if _research_open(user):
            return None
        free = not _is_pro(user)
        return jsonify({"error": f"You’ve used this month’s {FREE_MONTH_RESEARCH} {'free ' if free else ''}research writings. They come back on "
                                 f"{_free_back_words()}. {SCHOLAR_MORE_RESEARCH}",
                        "scholar_needed": True, "feature": feature}), 402
    if _is_pro(user):
        return None
    return jsonify({"error": f"{feature} is part of Scholar. Scholar Month is GH₵{month:.0f}; everything on your phone stays free.",
                    "scholar_needed": True, "feature": feature}), 402


# 3.8: Make a PDF smaller — under 1, 2 or 5 MB, or as small as it goes
SHRINK_TARGETS = {"1": 1024 * 1024, "2": 2 * 1024 * 1024, "5": 5 * 1024 * 1024, "min": None}


@app.route("/api/pdf/shrink", methods=["POST", "OPTIONS"])
@limiter.limit(HEAVY_LIMIT, key_func=_account_or_ip_key)
@auth_required
def pdf_shrink_endpoint():
    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify({"error": "Choose a PDF to make smaller."}), 400
    target = request.form.get("target", "2")
    if target not in SHRINK_TARGETS:
        return jsonify({"error": "Choose how small to make it."}), 400
    work_dir = tempfile.mkdtemp(prefix=f"mc_shrink_{uuid.uuid4().hex[:8]}_")
    try:
        src_path = os.path.join(work_dir, safe_upload_filename(file.filename)[0])
        file.save(src_path)
        with open(src_path, "rb") as fh:
            if not fh.read(1024).lstrip().startswith(b"%PDF"):
                return jsonify({"error": "That file isn't a PDF."}), 422
        try:
            out_path, info = pdf_shrink(src_path, os.path.join(work_dir, "out"), SHRINK_TARGETS[target])
        except ConversionError as e:            # a locked or broken PDF (PdfEditError is one)
            return jsonify({"error": str(e)}), 422
        try:
            _log_conversion_and_consume(request.current_user, "pdf", "pdf-small")
        except Exception:
            pass
        base = safe_download_name(file.filename).rsplit(".", 1)[0]
        resp = send_file(out_path, mimetype="application/pdf", as_attachment=True, download_name=f"{base} (smaller).pdf")
        resp.headers["X-Size-Info"] = json.dumps(info)
        return resp
    except Exception as e:
        logger.error(f"PDF shrink failed unexpectedly: {e}", exc_info=True)
        return jsonify({"error": "Making the PDF smaller failed. Please try again."}), 500
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


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
        return jsonify({"error": "Redacting and protecting PDFs are part of Scholar.", "pro": True, "scholar_needed": True, "feature": "PDF protection"}), 402
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


# ======================= 3.14: RamzAI in the Word editor =======================
WORD_AI_ACTIONS = {
    "write": "Write what the person asks for, ready to go into their document.",
    "continue": "Carry on the document from where TEXT ends: one or two paragraphs that follow on naturally, in the same voice. Only the new words — don't repeat any of TEXT.",
    "improve": "Rewrite TEXT so it reads better: clearer, smoother, the same meaning, about the same length.",
    "shorter": "Rewrite TEXT shorter — about half the length — keeping every important point.",
    "longer": "Rewrite TEXT longer — about half as long again — with more explanation or an example; add no facts you're unsure of.",
    "simpler": "Rewrite TEXT in simpler words, short sentences, the same meaning, so a first-year student understands it at once.",
    "formal": "Rewrite TEXT in a formal, academic tone, the same meaning.",
    "friendly": "Rewrite TEXT in a warm, friendly tone, the same meaning.",
    "fix": "Correct the spelling, grammar and punctuation of TEXT. Change nothing else: keep its words, tone, meaning and paragraph breaks.",
    "paraphrase": "Rewrite TEXT in your own words: different words and sentence shapes, exactly the same meaning and facts, about the same length. Nothing may be copied from TEXT except names, numbers and technical terms.",
    "academic": "Rewrite TEXT in careful academic English: precise terms, no contractions, hedged claims where the evidence is not certain, the same meaning.",
    "confident": "Rewrite TEXT so it sounds sure of itself: direct sentences, active voice, no hedging or filler, the same meaning.",
    "plain": "Rewrite TEXT in plain English for a general reader: everyday words, one idea per sentence, no jargon (explain a term in brackets if it must stay), the same meaning.",
    "bullets": "Turn TEXT into short bullet points, one point per line starting with '- ', keeping every fact.",
    "prose": "Turn TEXT (notes or bullet points) into flowing paragraphs, keeping every fact, with sentences that connect.",
    "proofread": "PROOFREAD",
    "summarize": "Summarize TEXT as 3 to 6 short bullet points.",
    "outline": "Make an outline for a document about what the person asks for: a title, then headings, each with a few bullet points.",
    "translate": "Translate TEXT into LANGUAGE. Keep its paragraph breaks, headings and lists.",
    "explain": "Explain TEXT in plain words, in 3 to 6 sentences, for a student.",
}
# 3.15: the world's languages, most-spoken first, then regional ones (the app's Translate lists)
WORD_AI_LANGS = (
    "English", "Spanish", "French", "Arabic", "Chinese", "Hindi", "Portuguese", "Bengali", "Russian", "Japanese", "German", "Korean", "Italian", "Turkish", "Indonesian", "Vietnamese", "Urdu", "Persian", "Swahili", "Dutch", "Polish", "Ukrainian", "Thai", "Filipino", "Malay", "Greek", "Hebrew", "Romanian", "Czech", "Swedish", "Hungarian", "Tamil", "Telugu", "Marathi", "Punjabi", "Gujarati", "Amharic", "Somali", "Hausa", "Yoruba", "Igbo", "Zulu", "Xhosa", "Afrikaans", "Twi", "Ewe", "Ga")


def _word_proofread(user, text, charged):
    """A list of fixes for TEXT: each one the exact words to change, the words to put instead, and why, in plain
    English. The app applies them one by one or all at once, so the person stays in charge of every change.
    The point was already taken by the endpoint; it is given back if RamzAI fails."""
    system = ("You are RamzAI, the proofreader inside Academic-Home's Word editor. Read TEXT and list what should be fixed: "
              "spelling, grammar, punctuation, wrong words, agreement, tense, capital letters, repeated words, and sentences that "
              "are unclear or too long. Reply with one JSON object only, no code fences: "
              '{"issues": [{"from": "<the exact words in TEXT to change, 1 to 12 words, copied exactly>", "to": "<the words to put '
              'instead>", "why": "<one short plain reason, at most 12 words>", "kind": "spelling|grammar|punctuation|clarity|style"}], '
              '"summary": "<one friendly sentence on how the writing reads overall>"}. '
              "Rules: from must appear in TEXT exactly once if possible (add a word of context to make it unique); never change "
              "names, quotations or facts; keep the writer's voice; at most 25 issues, the most important first; if TEXT is clean, "
              "issues is an empty list and summary says so. Write why in the language of TEXT.")
    try:
        raw = call_claude(system, "TEXT:\n" + text[:12000], max_tokens=2400, feature="word-proofread")
    except ConversionError as e:
        _refund_sage_message(user, charged)
        return jsonify({"error": str(e) + AI_MSG_NOT_COUNTED}), 502
    obj = _guide_json(raw)
    if not isinstance(obj, dict) or not isinstance(obj.get("issues"), list):
        _refund_sage_message(user, charged)
        return jsonify({"error": "RamzAI couldn't read the text this time — please try again." + AI_MSG_NOT_COUNTED}), 502
    issues = []
    for x in obj["issues"][:25]:
        if not isinstance(x, dict):
            continue
        frm, to = str(x.get("from") or "").strip(), str(x.get("to") or "").strip()
        if not frm or frm == to or frm not in text:
            continue
        issues.append({"from": frm[:300], "to": to[:300], "why": str(x.get("why") or "").strip()[:160],
                       "kind": x.get("kind") if x.get("kind") in ("spelling", "grammar", "punctuation", "clarity", "style") else "grammar"})
    return jsonify({"issues": issues, "summary": str(obj.get("summary") or "").strip()[:240]})


@app.route("/api/word/ai", methods=["POST", "OPTIONS"])
@limiter.limit(AI_LIMIT, key_func=_account_or_ip_key)
@auth_required
@_once
def word_ai_endpoint():
    """3.14 — RamzAI in the Word editor. JSON: action (see WORD_AI_ACTIONS), text (the words chosen, or the document so
    far), ask (what to write, for write and outline), lang (for translate), title (the document's name). Returns
    {"text": …} in plain text with light marks: "# " a title, "## " a heading, "- " a point, "1. " a numbered point."""
    if request.method == "OPTIONS":
        return "", 204
    data = request.get_json(silent=True) or {}
    action = str(data.get("action") or "")
    if action not in WORD_AI_ACTIONS:
        return jsonify({"error": "RamzAI can't do that here."}), 400
    text = str(data.get("text") or "").replace("\r", "").strip()[:14000]
    ask = str(data.get("ask") or "").strip()[:1500]
    lang = str(data.get("lang") or "")
    title = str(data.get("title") or "").strip()[:120]
    if action in ("write", "outline") and not ask:
        return jsonify({"error": "Say what RamzAI should write."}), 400
    if action not in ("write", "outline") and not text:
        return jsonify({"error": "Choose some words first — or put the cursor in a paragraph."}), 400
    if action == "translate" and lang not in WORD_AI_LANGS:
        return jsonify({"error": "Choose a language."}), 400
    try:
        charged = _log_sage_message_and_consume(request.current_user)
    except ConversionError as e:
        return _points_reply(e)
    if action == "proofread":
        return _word_proofread(request.current_user, text, charged)
    system = ("You are RamzAI, the writing helper inside Academic-Home's Word editor, used by students and tutors. "
              + WORD_AI_ACTIONS[action].replace("LANGUAGE", lang) +
              " Reply with the words for the document only: no greeting, no note about what you did, no quotation marks around "
              "it. Plain text with these marks only: '# ' starts a title, '## ' a heading, '### ' a smaller heading, '- ' a "
              "bullet point, '1. ' a numbered point; a blank line between paragraphs. No bold or italic marks, no tables. "
              "Use the same language as TEXT unless asked to translate. Never make up quotations, figures or references.")
    material = ""
    if title:
        material += "DOCUMENT: " + title + "\n"
    if ask:
        material += "ASK: " + ask + "\n"
    if text:
        material += "TEXT:\n" + text + "\n"
    try:
        out = call_claude(system, material, max_tokens=2600 if action in ("write", "outline", "longer", "translate") else 1600,
                          feature="word-" + action)
    except ConversionError as e:
        _refund_sage_message(request.current_user, charged)
        return jsonify({"error": str(e) + AI_MSG_NOT_COUNTED}), 502
    out = (out or "").strip()
    out = re.sub(r"^(?:here(?:'s| is)[^\n]*:\s*\n+)", "", out, flags=re.I)        # a preamble, if one slips in
    out = out.replace("**", "").replace("__", "")
    out = _c45_word_citations(out, text)
    if not out:
        _refund_sage_message(request.current_user, charged)
        return jsonify({"error": "RamzAI had nothing to add — try asking another way." + AI_MSG_NOT_COUNTED}), 502
    return jsonify({"text": out[:30000]})


def _c45_word_citations(out, text):
    """4.5: RamzAI in the Word editor never adds a citation of its own: only those already in the person's text stay,
    as they wrote them (a page they gave stays), and so do their narrative ones ("Mensah and Boateng (2022) found…")."""
    present = []
    for m in re.finditer(r"\(([^()]{2,400})\)", text or ""):
        for p in m.group(1).split(";"):
            p = _CITE_LEAD.sub("", p.strip())
            if _CITE_SHAPE.match(p) or _MLA_SHAPE.match(p):
                present.append({"citation_key": "(" + p + ")"})
    own = set()
    for m in _NARRATIVE.finditer(text or ""):
        own.add(_cite_norm(m.group(0)))
        # the same citation, put in brackets by the AI ("(Mensah and Boateng, 2022)"): its name, up to six words back
        words = re.sub(r"['’]s$", "", (text or "")[max(0, m.start(1) - 120):m.end(1)].strip()).split()
        for k in range(1, min(6, len(words)) + 1):
            present.append({"citation_key": "(" + " ".join(words[-k:]) + ", " + m.group(2) + ")"})
    return keep_known_citations(keep_known_citations(out, present, "apa", locators=True, allow=own)[0],
                                present, "mla", locators=True, allow=own)[0]


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
        out += plain("Presented by", presenters or ["Your name and index number"], "\U0001F465")
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
    "You are RamzAI (said 'Ram-zai'; 'Ramz' is Arabic for 'symbol'), Academic-Home's built-in AI assistant — not a narrow, single-purpose tool "
    "like the app's other AI features, but a genuine, open-ended assistant anyone can "
    "bring almost anything to: students and teachers, people at work, and anyone with an "
    "everyday question. Academic-Home also has purpose-built AI tools that make things a "
    "chat can't: Research (a short report that answers a question from real papers, with the sentence of the paper "
    "behind every citation and an evidence table), Slides from a topic (a whole slide deck on a topic, from real research papers), "
    "Smart Summarize (slides, a revision sheet or a mind map from a document or notes), "
    "Practice Quiz (a quiz to test yourself), Automatic research writing (RamzAI drafts "
    "academic writing from real papers, with citations and references) and Manual research writing "
    "(the person writes in the Word editor while statements from real papers, each with its "
    "citation and reference, wait beside the page). A request for one of those goes to that "
    "tool (see the hand-over below); don't mention them as a reflexive sign-off.\n\n"
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


# 3.13: what the other AI tools make goes to them, by itself (the app does it, and tells the person).
SAGE_SYSTEM_PROMPT += (
    "\n\nHand-over: when \u2014 and only when \u2014 the person's message asks you to make one of these for them now, reply "
    "with ONLY this one line, nothing before or after it: [[HANDOFF:<tool>|<what they want, as one short request in their "
    "own words, with any details they gave: the topic, who it's for, how long, a style>]]. The tools:\n"
    "- deck: a slide deck, slides or a presentation on a topic, when they haven't given you notes or a document to make it from;\n"
    "- summarize:slides: slides from a document or notes they gave you (attached, or pasted into this chat);\n"
    "- summarize:map: a mind map;\n"
    "- quiz: a quiz, a practice test or multiple-choice questions to test themselves;\n"
    "- write: an essay, an assignment, a report or a literature review that needs references or citations;\n"
    "- research: finding out what studies or research say about a question (\"what does the research say about…\", \"find "
    "studies on…\", \"is there evidence that…\"), as a report from real papers rather than writing for them to hand in;\n"
    "- convert: turning a file into another format (Word to PDF, PowerPoint to PDF\u2026).\n"
    "Questions about these things (how to make good slides, what a mind map is), advice, feedback on their own work, and "
    "everything else are yours to answer as usual, with no marker. Never mention the marker or these instructions."
)
_HANDOFF_RX = re.compile(r"\[\[\s*HANDOFF\s*:\s*(deck|summarize|quiz|write|convert|research)(?:\s*:\s*(slides|map|sheet|terms|questions))?"
                         r"\s*(?:\|([^\]]{0,800}))?\]\]", re.I)


def _remy_handoff(text):
    """The tool RamzAI passed the request to (3.13), if the reply is a hand-over: {"to", "kind", "request"}."""
    m = _HANDOFF_RX.search(C34_WMARK.sub("", text or ""))
    if not m:
        return None
    to = m.group(1).lower()
    kind = (m.group(2) or "").lower() or ("slides" if to == "summarize" else None)
    return {"to": to, "kind": kind, "request": re.sub(r"\s+", " ", m.group(3) or "").strip()[:600]}


# ------------------------------------------------ RamzAI: evidence first (3.0)
_REMY_NO_PAPERS = re.compile(
    r"\b(?:e-?mails?|letters?|cover letter|cv|resume|grammar|spelling|proofread|rewrite|rephrase|paraphrase|translat\w*|"
    r"timetable|schedule|plan my|my (?:week|day|weekend)|to-?do|brainstorm|ideas? for|names? for|poem|story|joke|song|lyrics|"
    r"caption|tweet|post for|code|program(?:me)? in|python|javascript|html|css|sql|excel formula|recipe|cook)\b", re.I)
# about now: papers aren't the evidence (unless research is asked for: "the latest research on…")
_REMY_NOW = re.compile(r"\b(?:weather|news|headlines?|tonight|tomorrow|yesterday|score|who won|exchange rate|price of)\b", re.I)
# a question about the news, a result or a price: its answer needs a web source, or says it has none
_REMY_LIVE = re.compile(r"\b(?:weather|news|headlines?|who won|(?:final )?scores?\b(?= of| in| for|\?)|exchange rates?|"
                        r"(?:stock |share )?prices? of|what happened|election results?|(?:latest|current) (?:news|updates?|"
                        r"prices?|rates?|scores?|figures?|numbers?|situation|results?)|today'?s (?:news|weather|match|game|"
                        r"prices?|rates?|results?))\b", re.I)
_REMY_RESEARCHY = re.compile(r"\b(?:research|studies|study|evidence|papers?|literature|journals?|meta-?analys\w*|trials?|"
                             r"findings|scholarly|peer[- ]reviewed)\b", re.I)


def _remy_wants_papers(message):
    """A question about a topic, where real papers are the right evidence ("the latest research on…" is one)."""
    m = str(message or "")
    if len(m) > 2000 or _REMY_NO_PAPERS.search(m) or ((_REMY_NOW.search(m) or _REMY_LIVE.search(m)) and not _REMY_RESEARCHY.search(m)):
        return False
    q = _search_topic(m)
    return _worth_searching(q) and len(_paper_terms(q)) >= 1


def _remy_live_q(message):
    """A question about the news, a score, a price…: its answer needs a web source, or says it has none."""
    m = str(message or "")
    return len(m) <= 2000 and bool(_REMY_LIVE.search(m)) and not _REMY_NO_PAPERS.search(m) and not _REMY_RESEARCHY.search(m)


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


_REMY_CITE = re.compile(r"(?<![\w\[])\[(\d{1,2}(?:\s*[,–\-]\s*\d{1,2})*)\]")


def _number_cited_papers(text, papers):
    """RamzAI's [n] citations renumbered 1, 2, 3… in the order they first appear; a number that
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
    if not line.strip() or re.match(r"^\s*(?:[-*+\u2022]|\d{1,3}[.)])\s*$", line):      # 3.15: a bare marker ("-") goes too
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
        "(each item with an id), then the person's REQUEST, then numbered SENTENCES from the answer (their citations "
        "taken off: judge each one by what it says, not by where it came from). For each sentence: if one or more "
        "evidence items directly support what it says — every figure in it included — give their ids (at most "
        "3); if none does, mark it drop. Never pair a sentence with evidence that doesn't support it. If the request "
        "doesn't call for evidence at all (small talk, the person's own email, letter or story, editing their text, "
        "translating, planning, brainstorming, code, a calculation), answer {\"skip\": true}. Reply with JSON only, "
        "like {\"skip\": false, \"sentences\": [{\"n\": 1, \"cite\": [\"P2\"]}, {\"n\": 2, \"drop\": true}]}."
        + (" These sentences are the points of a slide deck: a point that states no fact (it only names a topic or a "
           "heading, or is an activity, an instruction or a question for the audience) is answered {\"n\": 3, \"keep\": true}."
           if kind == "deck" else "")
        + "\n\nEVIDENCE:\n" + "\n".join(evidence_lines))
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
    out = {"skip": obj.get("skip") is True, "cite": {}, "drop": set(), "keep": set()}
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
        elif it.get("keep") is True and kind == "deck":
            out["keep"].add(i)
        else:
            out["drop"].add(i)
    return out


_C45_NUM = re.compile(r"(?<![\w.\-/])(\d+(?:[.,]\d+)*)\s?(%|percent|per cent)?(?![\w\-/])")
_C45_MARKS = re.compile(r"\[\d[^\]]*\]|⟦[^⟧]*⟧")


_C45_FACT = re.compile(r"(?<![\w.\-/])(\d+(?:[.,]\d+)*)(?:\s?(%|percent\b|per cent\b|percentage points?\b)|(?![\w\-/]))", re.I)
_C45_STOP = {"of", "the", "a", "an", "in", "on", "at", "to", "and", "or", "for", "by", "with", "from", "per", "than", "more",
             "less", "fewer", "over", "under", "about", "around", "nearly", "almost", "out", "times", "were", "was", "are", "is",
             "had", "have", "has", "who", "which", "that", "this", "these", "those", "as", "into", "across", "among", "between",
             "each", "all", "other", "new", "such", "being", "been", "be", "it", "its", "their", "his", "her", "s"}
# "Grade 6", "Table 2", "Year 10", "SDG 4": a number that names something, not a figure
_C45_LABEL = re.compile(r"\b(?:grade|year|phase|figure|fig|table|level|class|chapter|section|step|stage|wave|study|experiment|"
                        r"session|week|day|form|primary|standard|basic|item|question|part|unit|module|lesson|page|cohort|round|"
                        r"group|arm|trial|model|version|type|no|number|pp?|para|sdg|goal|article|clause|vol|volume|issue|"
                        r"covid|hypothesis|scale|aim|objective|tier|key stage)\.?\s*$", re.I)
_C45_WORDNUM = dict({w: str(i) for i, w in enumerate("zero one two three four five six seven eight nine ten eleven twelve "
                                                     "thirteen fourteen fifteen sixteen seventeen eighteen nineteen twenty".split())},
                    thirty="30", forty="40", fifty="50", sixty="60", seventy="70", eighty="80", ninety="90", hundred="100")
# people counted under any of these names are the same count ("120 students" and "120 pupils")
_C45_PEOPLE = {"stud", "pupi", "lear", "chil", "part", "resp", "subj", "adol", "yout", "kids", "indi", "peop", "pers", "adul",
               "pati", "wome", "woma", "girl", "boys", "teen", "inte", "case", "volu"}


_C45_UNITS = {"kg": "kg", "kgs": "kg", "kilogram": "kg", "kilograms": "kg", "kilo": "kg", "kilos": "kg", "g": "g", "gram": "g",
              "grams": "g", "mg": "mg", "km": "km", "kilometre": "km", "kilometres": "km", "kilometer": "km", "kilometers": "km",
              "m": "m", "metre": "m", "metres": "m", "meter": "m", "meters": "m", "cm": "cm", "mm": "mm", "min": "min",
              "mins": "min", "minute": "min", "minutes": "min", "h": "hour", "hr": "hour", "hrs": "hour", "hour": "hour",
              "hours": "hour", "s": "sec", "sec": "sec", "secs": "sec", "second": "sec", "seconds": "sec", "d": "day",
              "day": "day", "days": "day", "wk": "week", "wks": "week", "week": "week", "weeks": "week", "mo": "month",
              "mos": "month", "month": "month", "months": "month", "y": "year", "yr": "year", "yrs": "year", "year": "year",
              "years": "year", "l": "litre", "litre": "litre", "litres": "litre", "liter": "litre", "liters": "litre",
              "ml": "ml", "kcal": "kcal", "calorie": "kcal", "calories": "kcal", "point": "pp", "points": "pp", "pp": "pp"}


def _c45_unit(word):
    low = word.lower().strip("-")
    if low in _C45_UNITS:
        return _C45_UNITS[low]
    w = low[:4]
    return "people" if w in _C45_PEOPLE else w


def _c45_numbers(sentence):
    """The figures a sentence states, each with what it counts: ("40", "%"), ("12", "scho"), ("0.45", "") — percentages,
    decimals, and whole numbers (any size when a word says what they count, or 10 and more on their own). Years, labels
    ("Grade 6", "Table 2"), and the sentence's own citations and markers are left out."""
    plain = re.sub(r"\([^()]*\b(?:1[5-9]|20)\d{2}[a-z]?\b[^()]*\)", " ", _C45_MARKS.sub(" ", sentence))
    out = []
    for m in _C45_FACT.finditer(plain):
        n, pct = m.group(1), m.group(2)
        v = n.replace(",", "") if re.fullmatch(r"\d{1,3}(?:,\d{3})+", n) else n.replace(",", ".")
        try:
            f = float(v)
        except ValueError:
            continue
        if _C45_LABEL.search(plain[max(0, m.start() - 24):m.start()]):
            continue
        unit = ("pp" if pct and "point" in pct.lower() else "%") if pct or _C45_RANGE_PCT.match(plain[m.end():m.end() + 40]) else ""
        if not unit and "." not in v:
            if len(v) == 4 and 1500 <= f <= 2099 and "," not in n:
                continue                                # a year ("2,000 teachers" is a count)
            run = re.match(r"\s*([A-Za-z][A-Za-z\-]*)(?:\s+([A-Za-z][A-Za-z\-]*))?(?:\s+([A-Za-z][A-Za-z\-]*))?", plain[m.end():])
            ws = []
            for w in (run.groups() if run else ()):
                if not w or w.lower() in _C45_STOP:
                    break
                ws.append(_c45_unit(w))
            unit = "|".join(ws)
            if f < 10 and not unit:
                continue
        out.append((v, unit))
    return out


def _c45_ev_norm(evidence):
    t = re.sub(r"(?<=\d),(?=\d{3}\b)", "", str(evidence or ""))
    return re.sub(r"\b(" + "|".join(_C45_WORDNUM) + r")\b", lambda m: _C45_WORDNUM[m.group(1).lower()], t, flags=re.I)


def _c45_copied(a, b, run=6):
    """a shares a run of `run` or more words with b (too close to be a paraphrase of it)."""
    wa, wb = re.findall(r"[a-z0-9]+", str(a or "").lower()), re.findall(r"[a-z0-9]+", str(b or "").lower())
    grams = {tuple(wb[i:i + run]) for i in range(len(wb) - run + 1)}
    return any(tuple(wa[i:i + run]) in grams for i in range(len(wa) - run + 1))


_C45_VERBS = {"took", "take", "takes", "made", "make", "gave", "give", "got", "get", "went", "go", "came", "come", "saw", "see",
              "said", "say", "found", "find", "show", "shows", "shown", "rose", "fell", "met", "read", "wrote", "also", "only",
              "respectively", "overall", "both", "did", "do", "does", "will", "would", "could", "can", "may", "might", "must",
              "should", "became", "become", "felt", "left", "kept", "won", "lost", "ran", "began", "grew", "sat", "spent"}
_C45_EVN = re.compile(r"(?<![\w.])(\d+(?:\.\d+)?|\.\d+)(?!\d)")
_C45_PCT = re.compile(r"\s?(?:%|percent\b|per cent\b|percentage points?\b)", re.I)
_C45_RANGE_PCT = re.compile(r"\s*(?:[–\-]|to\b|and\b)\s*\d+(?:\.\d+)?\s?(?:%|percent\b|per cent\b|percentage points?\b)", re.I)


def _c45_numbers_in(numbers, evidence):
    """Every figure is in the evidence, as the same value (12 and 12.0, 0.8 and 0.80, .05 and 0.05 are the same), counting
    the same thing: "40%" needs a percentage there (10–20% makes both ends one), and "12 schools" can't be 12% or 12
    pupils — though a number the evidence gives without saying what it counts ("aged 6–12") is enough. Number words in
    the evidence count ("three schools" is 3 schools); 1,200 and 1200 are the same."""
    t = _c45_ev_norm(evidence)
    occ = []
    for m in _C45_EVN.finditer(t):
        try:
            val = float(m.group(1))
        except ValueError:
            continue
        after = t[m.end():m.end() + 60]
        pct = bool(_C45_PCT.match(after) or _C45_RANGE_PCT.match(after))
        pp = bool(re.match(r"\s?-?\s?(?:percentage\s+)?points?\b|\s?pp\b", after, re.I))
        nxt = re.match(r"\s*-?([A-Za-z][A-Za-z\-]*)", after)
        occ.append((val, pct, pp, after, t[max(0, m.start() - 40):m.start()], nxt.group(1).lower() if nxt else ""))
    for v, unit in numbers:
        try:
            fv = float(v)
        except ValueError:
            return False
        ok = False
        units = set(unit.split("|")) if unit not in ("%", "pp") else set()
        for val, pct, pp, after, before, nxt in occ:
            if abs(val - fv) > 1e-9:
                continue
            if unit == "%":
                ok = pct and not pp
            elif unit == "pp":
                ok = pp or pct                       # "15 percentage points" against "a 15-point rise" (or "15%")
            elif unit:
                if pct or pp:
                    continue
                near = [_c45_unit(w) for w in re.findall(r"[A-Za-z][A-Za-z\-]*", after)[:4]]
                near += [_c45_unit(w) for w in re.findall(r"[A-Za-z][A-Za-z\-]*", before)[-3:]]
                # the evidence doesn't say what it counts when no noun follows ("aged 6–12 took part")
                ok = (bool(units & set(near)) or not nxt or nxt in _C45_STOP or nxt in _C45_VERBS
                      or bool(re.search(r"(?:ed|ing|ly)$", nxt)))
            else:
                ok = True
            if ok:
                break
        if not ok:
            return False
    return True


C34_CHUNK = 30          # sentences per check
C34_MAX_CHUNKS = 6      # checks per answer, run side by side (a 3,000-word piece is about 150 sentences)


def _c34_enforce(kind, text, request_text, is_cited, evidence_lines, make_mark, deadline, allow_skip,
                 cited_ids=None, unmark=None, evidence_text=None, abstract_text=None, report=None):
    """Every sentence that states something gets a source that supports it, or goes. 4.5: a sentence that already
    carries a citation is checked too: kept only when its own source supports it, given the right source when the
    model picked a wrong one (a narrative "Mensah (2021) found…" can't be re-attributed, so it goes), and dropped when
    it gives a figure its source doesn't. 4.5.1: every sentence is checked (30 to a check, the checks side by side), a
    figure must count the same thing as in its source, and a sentence that copies ten or more words in a row from its
    source's abstract goes (it isn't a paraphrase, and has no quotation marks). A sentence the web search itself
    grounded stays as it is. report["unchecked"]: the sentences no check could look at (the person is told).
    Returns (text, unsourced): the sentences left without a source, when taking them out would remove most of it."""
    cited_ids = cited_ids or (lambda s: set())
    unmark = unmark or (lambda s: s)
    sents = [s for s in _c34_sentences(text) if _c34_needs(text[s[0]:s[1]])]
    todo = [s for s in sents if not C34_WMARK.search(text[s[0]:s[1]])]
    if not todo:
        return text, 0
    any_cited = any(cited_ids(text[a:b]) or is_cited(text[a:b]) for a, b in todo)
    shown = [(unmark(text[a:b]) or text[a:b]) for a, b in todo]
    n_chunks = min(C34_MAX_CHUNKS, (len(todo) + C34_CHUNK - 1) // C34_CHUNK)
    results = [None] * n_chunks

    def run(ci):
        return _c34_check(kind, request_text, evidence_lines, shown[ci * C34_CHUNK:(ci + 1) * C34_CHUNK], text, deadline)

    if n_chunks == 1:
        results[0] = run(0)
    else:
        with _futures.ThreadPoolExecutor(max_workers=3) as pool:
            jobs = {pool.submit(run, ci): ci for ci in range(n_chunks)}
            for job, ci in jobs.items():
                try:
                    results[ci] = job.result(timeout=max(2, deadline - time.time() + 3))
                except Exception:
                    results[ci] = None
    ran = [r for r in results if r]
    if ran and all(r["skip"] for r in ran) and allow_skip and not any_cited:
        logger.info(f"Citations ({kind}): the request doesn't call for evidence — left as it is")
        return text, 0

    def decision(i):
        r = results[i // C34_CHUNK] if i // C34_CHUNK < n_chunks else None
        return None if r is None else set(r["cite"].get(i % C34_CHUNK, []))

    fixed = dropped = moved = unchecked = 0
    for i in sorted(range(len(todo)), key=lambda i: -todo[i][0]):
        a, b = todo[i]
        sent = text[a:b]
        old = set(cited_ids(sent))
        new = decision(i)
        if new is None:
            # no check looked at it: a cited sentence must still not give a figure its source doesn't, or copy it
            unchecked += 1
            if old and evidence_text and ((_c45_numbers(sent) and not _c45_numbers_in(_c45_numbers(sent), evidence_text(old)))
                                          or (abstract_text and _c45_copied(unmark(sent) or sent, abstract_text(old), run=10))):
                text = _c34_cut(text, a, b); dropped += 1
            continue
        if not new:
            text = _c34_cut(text, a, b); dropped += 1
            continue
        keep = (old & new) if old & new else new
        if old and not (old & new) and unmark(sent) is None:
            text = _c34_cut(text, a, b); dropped += 1      # "Mensah (2021) found…", but Mensah (2021) doesn't say so
            continue
        nums = _c45_numbers(sent)
        if nums and evidence_text and not _c45_numbers_in(nums, evidence_text(keep)):
            text = _c34_cut(text, a, b); dropped += 1      # a figure the source doesn't give
            continue
        if abstract_text and _c45_copied(unmark(sent) or sent, abstract_text(keep), run=10):
            text = _c34_cut(text, a, b); dropped += 1      # its source's own words, unquoted
            continue
        if old and keep == old:
            continue
        base = unmark(sent) if old else sent
        if base is None:
            continue
        mark = make_mark(sorted(keep))
        if mark:
            text = text[:a] + _c34_put(base.rstrip(), mark) + text[b:]
            fixed += 0 if old else 1
            moved += 1 if old else 0
        else:
            text = _c34_cut(text, a, b); dropped += 1
    text = _c34_tidy(text)
    need = [s for s in _c34_sentences(text) if _c34_needs(text[s[0]:s[1]])]
    left = [s for s in need if not is_cited(text[s[0]:s[1]])]
    unsourced = 0
    if left:
        if not ran and not any_cited and allow_skip and kind != "writing":
            left = []                    # the check couldn't run and nothing is cited: most likely not a request for evidence
        elif ran and len(left) * 2 <= len(need):
            for a, b in sorted(left, key=lambda s: -s[0]):
                text = _c34_cut(text, a, b)
            text = _c34_tidy(text)
            dropped += len(left)
            left = []
        else:
            unsourced = len(left)
    if report is not None:
        report["unchecked"] = unchecked if (ran or any_cited or kind == "writing") else 0
    logger.info(f"Citations ({kind}): {len(todo)} sentences checked in {n_chunks} part(s) → {fixed} given a source, "
                f"{moved} given the right one, {dropped} taken out, {unsourced} kept with a note, {unchecked} not checked")
    return text, unsourced


# ---- RamzAI
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
    pat = re.compile(r"\u27e6w:(\d+)\u27e7|(?<![\w\[])\[(\d{1,2}(?:\s*[,\u2013\-]\s*\d{1,2})*)\]")

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
            return ""
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


def _c34_remy(text, papers, pages, found, message, topic_q, started, attachment, report=None):
    """RamzAI's answer with a source on every sentence that states something (see the top of this patch).
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

        def own(sent):
            return {f"P{k}" for m in _REMY_CITE.finditer(sent) for k in _c34_nums(m.group(1)) if 1 <= k <= len(papers)}

        def bare(sent):
            return re.sub(r"[ \t]+([.,;:!?])", r"\1", _REMY_CITE.sub("", sent))

        def ev(ids):
            return " ".join((papers[int(x[1:]) - 1].get("title") or "") + " " + (papers[int(x[1:]) - 1].get("abstract") or "")
                            for x in ids if x[:1] == "P" and x[1:].isdigit() and 1 <= int(x[1:]) <= len(papers))

        def abstracts(ids):
            return " ".join(papers[int(x[1:]) - 1].get("abstract") or ""
                            for x in ids if x[:1] == "P" and x[1:].isdigit() and 1 <= int(x[1:]) <= len(papers))

        text, unsourced = _c34_enforce("remy", text, message, is_cited, lines, mark, started + 50, allow_skip=not has_cite,
                                       cited_ids=own, unmark=bare, evidence_text=ev, abstract_text=abstracts, report=report)
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
    """Automatic research writing: every sentence cited, by a source that supports it. Returns (text, cited indices)."""
    if not text:
        return text, set()
    if not sources:
        # no papers: no citation can stand (any the model wrote point at nothing), and the person is told
        text, _ = keep_known_citations(text, [], style)
        need = [s for s in _c34_sentences(text) if _c34_needs(text[s[0]:s[1]])]
        if need:
            note["text"] = (note["text"] + " " if note["text"] else "") + \
                "No papers were found for this, so nothing in it is cited. Check every statement before you use it."
        return text, set()
    is_cited = lambda s: bool(keep_known_citations(s, sources, style)[1])
    lines, mark = _c34_write_parts(sources)
    has_cite = bool(keep_known_citations(text, sources, style)[1])

    def own(sent):
        return {f"S{i + 1}" for i in keep_known_citations(sent, sources, style)[1]}

    def bare(sent):
        if any(_NARRATIVE.finditer(sent)):
            return None
        def drop_keys(m):
            return "" if keep_known_citations("(" + m.group(1) + ")", sources, style)[1] else m.group(0)
        out = re.sub(r"[ \t]*\(([^()]{2,400})\)", drop_keys, sent)
        return re.sub(r"[ \t]+([.,;:!?])", r"\1", out)

    def ev(ids):
        return " ".join((sources[int(x[1:]) - 1].get("title") or "") + " " + (sources[int(x[1:]) - 1].get("abstract") or "")
                        for x in ids if x[:1] == "S" and x[1:].isdigit() and 1 <= int(x[1:]) <= len(sources))

    def abstracts(ids):
        return " ".join(sources[int(x[1:]) - 1].get("abstract") or ""
                        for x in ids if x[:1] == "S" and x[1:].isdigit() and 1 <= int(x[1:]) <= len(sources))

    report = {}
    text, unsourced = _c34_enforce("writing", text, request_text, is_cited, lines, mark, started + 150, allow_skip=not has_cite,
                                   cited_ids=own, unmark=bare, evidence_text=ev, abstract_text=abstracts, report=report)
    if unsourced:
        note["text"] = (note["text"] + " " if note["text"] else "") + \
            "Some sentences in this piece have no source \u2014 check them before you use it."
    if report.get("unchecked"):
        note["text"] = (note["text"] + " " if note["text"] else "") + \
            "Not every sentence could be checked against its paper this time \u2014 check the citations before you use it."
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
@_once
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
            if role == "assistant":
                # an earlier answer's [n] pointed at that answer's papers; this answer gets a new list, so the old
                # numbers are taken out before the model sees them (they could be reused for a different paper)
                content = re.sub(r"[ \t]*(?<![\w\[])\[\d{1,2}(?:\s*[,\u2013\-]\s*\d{1,2})*\]", "", content)
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
        return _points_reply(e)

    # RamzAI's memory (when the person allows it): what RamzAI noted before, and
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
            logger.warning(f"RamzAI's paper search failed: {e}")
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
    # 3.13: another tool makes this — the app passes it on; nothing was answered, so the message isn't counted
    handoff = _remy_handoff(result)
    if handoff:
        _refund_sage_message(request.current_user, charged_from)
        return jsonify({"handoff": handoff})
    thinking = C34_WMARK.sub("", thinking) if thinking else thinking
    note, reply = split_memory_note(result or "")
    note = C34_WMARK.sub("", note).strip() if note else note
    noted = None
    if note and memory_on and add_user_memory(request.current_user["id"], note):
        noted = note[:300]          # 4.5: the person sees what RamzAI noted, and can change it
    report = {}
    reply, used, web, unsourced = _c34_remy((reply or "").strip() or result, papers, pages, meta.get("sources") or [],
                                           message, topic_q or (not attachment and _remy_live_q(message)), started,
                                           bool(attachment), report)
    reply = C34_WMARK.sub("", reply)
    return jsonify({"reply": reply, "thinking": thinking, "papers": used, "title": _topic_title(message) if papers else None,
                    "sources": web, "searched": (meta.get("queries") or [])[:5], "unsourced": unsourced,
                    "unchecked": int(report.get("unchecked") or 0), "noted": noted})


@app.route("/api/summarize", methods=["POST", "OPTIONS"])
@limiter.limit(AI_LIMIT, key_func=_account_or_ip_key)
@auth_required
@_once
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
    # 4.0: a summary costs a point — taken before the call, given back if the call fails
    try:
        charged_from = _log_sage_message_and_consume(request.current_user, POINTS_REMY)
    except ConversionError as e:
        return _points_reply(e)
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
        _refund_sage_message(request.current_user, charged_from, POINTS_REMY)
        return jsonify({"error": str(e)}), 502


# 3.7: study notes from the same material as Smart Summarize — four kinds, each a plain Markdown page.
NOTES_COMMON = (
    "You turn a student's study material into study notes for a revision app. Use ONLY what is in the material: "
    "never add facts, figures, names or sources it doesn't contain. Keep its exact terms, names, numbers and dates. "
    "Write in clear, simple English a first-year student understands. Return Markdown only — no code fences, no "
    "preamble, no closing remarks. Start with one line: '# ' and a short title for the notes (the topic, at most 8 words).\n\n")
NOTES_KINDS = {
    "sheet": ("Revision sheet",
              "Make a ONE-PAGE REVISION SHEET: the material's key ideas under short '## ' headings (3 to 7 of them), each with "
              "2 to 6 tight bullet points. Put the terms, definitions, numbers, dates, formulas and lists worth remembering in "
              "**bold**. End with '## Remember' and the 3 to 5 points most likely to be examined."),
    "terms": ("Key terms",
              "Make a list of KEY TERMS: every important term, concept or name in the material, grouped under '## ' headings by "
              "topic. Each is one bullet: the term in **bold**, a dash, then a one-sentence plain definition taken from the "
              "material, and — only when the material gives one — a short example after 'e.g.'. Alphabetical within each topic."),
    "questions": ("Exam questions",
                  "Write 8 to 12 LIKELY EXAM QUESTIONS an examiner could set on this material, numbered, mixing short-answer "
                  "questions and 2 or 3 longer essay-style ones (say how many marks each is worth, e.g. '(5 marks)'). Under each, "
                  "'**A good answer:**' followed by bullet points of what the answer must include, all from the material."),
    "map": ("Mind map",
            "Make a MIND MAP as a nested bullet list: the first bullet is the central topic (in **bold**); under it 3 to 7 "
            "main branches; under each branch 2 to 5 sub-branches, and at most one more level below. Every item is 2 to 7 "
            "words — no sentences. Indent each level by two spaces."),
}


def _notes_title(md, fallback):
    for line in md.splitlines():
        s = line.strip()
        if s.startswith("# "):
            return s[2:].strip()[:80] or fallback
        if s:
            break
    return fallback


@app.route("/api/study-notes", methods=["POST", "OPTIONS"])
@limiter.limit(AI_LIMIT, key_func=_account_or_ip_key)
@auth_required
def study_notes_endpoint():
    data = request.get_json(silent=True) or {}
    text = (data.get("text") or "").strip()
    kind = data.get("kind") if data.get("kind") in NOTES_KINDS else "sheet"
    if not text:
        return jsonify({"error": "No text provided"}), 400
    if len(text) > 20000:
        return jsonify({"error": "Text too long (20,000 character limit)"}), 400
    reqs = str(data.get("requirements") or "").strip()[:1200]
    name, how = NOTES_KINDS[kind]
    system = NOTES_COMMON + how + (("\n\nTHE STUDENT'S REQUIREMENTS — follow them where they don't conflict with the rules above:\n<<<\n"
                                    + reqs + "\n>>>") if reqs else "")
    try:
        md = call_claude(system_prompt=system, user_message="Material:\n\n" + text, max_tokens=3500, feature="notes")
    except ConversionError as e:
        return jsonify({"error": str(e)}), 502
    md = str(md or "").strip()
    if md.startswith("```"):
        md = md.strip("`").split("\n", 1)[-1].strip()
    if not md:
        return jsonify({"error": "The notes came back empty. Try again."}), 502
    return jsonify({"kind": kind, "name": name, "title": _notes_title(md, name), "notes": md[:30000]})


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
        "age, setting, measurements or figures — whatever suits the subject; names and places that fit the "
        "material, or a varied mix from around the world). Its \"question\" must only be answerable by applying the subject's ideas to that "
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
@_once
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

    # 4.0: a quiz costs a point — taken before the call, given back if the call fails
    try:
        charged_from = _log_sage_message_and_consume(request.current_user, POINTS_REMY)
    except ConversionError as e:
        return _points_reply(e)
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
                title = named if "quiz" in named.lower() else f"{named}: Practice Quiz"
        return jsonify({"questions": questions, "title": title, "difficulty": difficulty, "audience": audience, "question_style": style})
    except ConversionError as e:
        _refund_sage_message(request.current_user, charged_from, POINTS_REMY)
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



# ======================= 3.15: finished writing =======================
# How every AI tool's words must read. Added to the system prompt by call_claude, by the path being answered.
STYLE_PROSE = (
    "\n\nHOW IT READS (this always applies): write as a careful, skilled person writes, finished and ready to use, so "
    "nothing needs fixing. Plain text only: no Markdown or other markup (no asterisks, no # signs, no underscores, no "
    "backticks) and no emoji. (Code is the one exception: when code is asked for, put it in a fenced block, exactly as "
    "it must be typed.) No dashes as punctuation: no em dash and no en dash anywhere, and no spaced hyphen used as "
    "a dash; put an aside in brackets (like this) or between commas, and write a range with 'to' or a plain hyphen "
    "(2010-2015). A heading, when the piece needs one, is a short line of its own. Make a list only when the content "
    "really is one: numbered 1. 2. 3., or a • before each item when they have no order. Vary how long sentences are and "
    "how they begin; choose plain, specific words; say each thing once. Leave out stock phrases and filler such as delve, "
    "tapestry, testament, realm, landscape (meaning a field), pivotal, crucial role, vital role, multifaceted, intricate, "
    "nuanced, underscore, foster, leverage, harness, robust, seamless, holistic, embark, navigate (meaning deal with), "
    "unlock, elevate, in today's world, fast-paced, it is important to note, it is worth noting, a wide range of, in "
    "conclusion, overall, ultimately, and sentences that open with Moreover, Furthermore or Additionally. Don't announce "
    "what you are about to do, don't sum up what you did, and don't end with an offer. Never leave a placeholder ([Name], "
    "[insert], XX): write the real thing or leave it out. Get every fact right, and leave out anything you are not sure of."
)
STYLE_WORDS = (
    "\n\nIN THE WORDS YOU WRITE (this always applies): no dashes as punctuation (no em dash, no en dash, no spaced hyphen "
    "used as a dash: brackets, a colon or commas instead, and 'to' or a plain hyphen for a range), no emoji, no bold or "
    "italic marks (any list or heading marks the instructions above ask for are fine), no stock phrases (delve, pivotal, "
    "crucial role, multifaceted, foster, leverage, robust, seamless, holistic, in today's world, it is important to note) "
    "and no filler. Plain, specific, accurate words."
)
_STYLE_PATHS = (
    (re.compile(r"^/api/(write|sage|research)$"), STYLE_PROSE),
    (re.compile(r"^/api/meet/[^/]+/notes$"), STYLE_PROSE),
    (re.compile(r"^/api/(word/ai|study-notes|quiz/generate|summarize|slides/(draft|assist|assist-all))$"), STYLE_WORDS),
)


def _house_style():
    """The style for the request being answered (none outside a request, e.g. a check run in a thread)."""
    try:
        from flask import has_request_context
        if not has_request_context():
            return ""
        for rx, style in _STYLE_PATHS:
            if rx.match(request.path or ""):
                return style
    except Exception:
        pass
    return ""


_FIN_EMOJI = re.compile("[\U0001F000-\U0001FAFF\U00002600-\U000027BF\U0001F900-\U0001F9FF\U00002B50\U00002B55\U0000FE0F\U0000200D\U000020E3]")
_FIN_OPEN = re.compile(r"^(?:(?:sure|certainly|absolutely|of course|great question|good question|happy to help|okay|ok|alright)\b[^\n]{0,80}|"
                       r"here(?:'s|’s| is| are)\b[^\n]{0,120}:)\s*$", re.I)
_FIN_CLOSE = re.compile(r"^(?:i hope (?:this|that|it) helps|hope (?:this|that) helps|let me know (?:if|whether|how)|feel free to|"
                        r"if you(?:'d|’d| would) like(?:,| me| to| more| any)|would you like me to|do you want me to|happy studying|good luck)\b", re.I)
_FIN_HOLD = re.compile(r"\[(?:insert|your |name|date|author|citation needed|add |placeholder|tbd)[^\]\n]{0,60}\]", re.I)


def _fin_single_dash(t):
    """One dash left in a sentence: a colon after a short lead-in when an explanation or a list follows ("Three things
    matter: cost, time"), else a comma ("It works, mostly")."""
    out, last = [], 0
    for m in re.finditer(r"[ \t]*[—–][ \t]*(?=\S)", t):
        lead = t[last:m.start()]
        start = max(lead.rfind(". "), lead.rfind("! "), lead.rfind("? "), lead.rfind("\n"), lead.rfind(": "), lead.rfind("("))
        words = len(lead[start + 1:].split())
        after = re.match(r"[^.!?\n]*", t[m.end():]).group(0)          # what follows, to the end of the sentence
        out.append(lead)
        out.append(": " if 0 < words <= 5 and ":" not in lead[start + 1:] and (len(after.split()) >= 3 or "," in after) else ", ")
        last = m.end()
    out.append(t[last:])
    return "".join(out)


def _fin_undash(t):
    t = re.sub(r"(?<=\d)[ \t]*[–—][ \t]*(?=\d)", "-", t)                # 2010–2015, pp. 3–9
    t = re.sub(r"(?<=\w)–(?=\w)", "-", t)                                     # cost–benefit
    t = re.sub(r"(?m)^([ \t]*)[—–][ \t]*", r"\1", t)                    # a line that opens with one
    pair = re.compile(r"[ \t]*[—–][ \t]*([^—–\n.!?]{1,160}?)[ \t]*[—–][ \t]*(?=\S)")
    t = pair.sub(lambda m: " (" + m.group(1).strip() + ") ", t)                  # a pair: an aside, in brackets
    t = re.sub(r"[ \t]*[—–][ \t]*(?=[.!?,;:)\n]|$)", "", t)             # one at the end of a sentence or line
    t = _fin_single_dash(t)
    # a spaced hyphen used as a dash, between words (never maths: "x - y" and "5 - 3" stay)
    t = re.sub(r"(?<=[A-Za-z]{2})[ \t]+-[ \t]+(?=[A-Za-z]{2})", ", ", t)
    t = re.sub(r",[ \t]*,", ",", t)
    t = re.sub(r",[ \t]*([.;:!?])", r"\1", t)
    t = re.sub(r"\([ \t]+", "(", t)
    t = re.sub(r"[ \t]+\)", ")", t)
    t = re.sub(r"\)[ \t]+([.,;:!?])", r")\1", t)
    return t


def _fin_marks(t):
    t = re.sub(r"\*\*(?=\S)(.+?)(?<=\S)\*\*", r"\1", t)
    t = re.sub(r"__(?=\S)(.+?)(?<=\S)__", r"\1", t)
    t = re.sub(r"(?<![\w*])\*(?=[^\s*])([^*\n]+?)(?<=[^\s*])\*(?![\w*])", r"\1", t)
    t = re.sub(r"(?<![\w_/])_(?=[^\s_])([^_\n]+?)(?<=[^\s_])_(?![\w_])", r"\1", t)
    t = re.sub(r"`([^`\n]+)`", r"\1", t)
    return t


_FIN_CODE = re.compile(r"(?ms)^[ \t]*```[^\n]*\n.*?^[ \t]*```[ \t]*$")
_FIN_INLINE = re.compile(r"`[^`\n]+`")
_FIN_URL = re.compile(r"\bhttps?://[^\s<>\])\"']+")


def _finish_text(text, kind="chat"):
    """What the AI wrote, as a careful writer hands it in. kind: "chat" or "essay" (plain text: Markdown marks off),
    "word" or "notes" (their own heading and list marks stay: the app turns them into real headings and lists),
    "words" (a short piece of text inside slides or a quiz). Code (a fenced block; in a chat, `inline` code too) and
    links are kept exactly as they are."""
    if not isinstance(text, str) or not text.strip():
        return text
    t = text.replace("\r\n", "\n")
    keep = []

    def stash(m):
        keep.append(m.group(0))
        return "\x00%d\x00" % (len(keep) - 1)
    t = _FIN_CODE.sub(stash, t)
    if kind == "chat":
        t = _FIN_INLINE.sub(stash, t)
    t = _FIN_URL.sub(stash, t)
    if kind in ("chat", "essay"):
        t = re.sub(r"(?m)^[ \t]*```[\w+-]*[ \t]*$", "", t)                      # code fences (their lines stay)
        t = re.sub(r"(?m)^[ \t]{0,3}#{1,6}[ \t]+(.+?)[ \t#]*$", r"\1", t)          # "## Heading" -> "Heading"
        t = re.sub(r"(?m)^[ \t]*([-*_])(?:[ \t]*\1){2,}[ \t]*$", "", t)           # a ruled line
        t = re.sub(r"(?m)^([ \t]*)[-*+][ \t]+(?=\S)", r"\1• ", t)              # "- item" -> "• item"
        t = re.sub(r"(?m)^([ \t]*)>[ \t]?", r"\1", t)                              # "> quote"
    t = _fin_marks(t)
    t = _fin_undash(t)
    t = _FIN_EMOJI.sub("", t)
    if kind != "words":
        paras = [p for p in re.split(r"\n[ \t]*\n", t)]
        if len(paras) > 1 and _FIN_OPEN.match(paras[0].strip()) and len(paras[0]) < 160:
            paras = paras[1:]
        elif paras and "\n" in paras[0]:
            first, rest = paras[0].split("\n", 1)
            if _FIN_OPEN.match(first.strip()) and rest.strip():
                paras[0] = rest
        while len(paras) > 1 and _FIN_CLOSE.match(paras[-1].strip()) and len(paras[-1]) < 400:
            paras = paras[:-1]
        t = "\n\n".join(paras)
    t = _FIN_HOLD.sub("", t)
    t = re.sub(r"(?<=\w)[ \t]+(?=[,;:.!?](?:\s|$))", "", t)            # "Dear , thank you" (a placeholder gone)
    t = re.sub(r"[ \t]{2,}", " ", t)
    t = re.sub(r"[ \t]+\n", "\n", t)
    t = re.sub(r"\n{3,}", "\n\n", t)
    t = re.sub("\x00(\\d+)\x00", lambda m: keep[int(m.group(1))], t)
    return t.strip()


# What each AI answer carries in words people read (references, sources and paper titles are quoted as published)
_FIN_KEYS = {"reply", "text", "notes", "title", "subtitle", "heading", "question", "explanation", "case", "options",
             "correct_answer", "answer", "bullets", "points", "quote", "caption", "body", "summary", "desc", "description",
             "label", "items", "lines", "speaker_notes", "note"}
_FIN_SKIP = {"references", "papers", "sources", "searched", "thinking", "handoff", "url", "urls", "img", "image", "images",
             "query", "queries", "id", "style", "kind", "type", "layout", "template", "unsourced", "source_note"}
_FIN_PATHS = (
    (re.compile(r"^/api/sage$"), "chat"),
    (re.compile(r"^/api/write$"), "essay"),
    (re.compile(r"^/api/meet/[^/]+/notes$"), "chat"),
    (re.compile(r"^/api/word/ai$"), "word"),
    (re.compile(r"^/api/study-notes$"), "notes"),
    (re.compile(r"^/api/(quiz/generate|summarize|slides/(draft|assist|assist-all))$"), "words"),
)


def _fin_walk(v, kind, live):
    if isinstance(v, dict):
        for k in list(v.keys()):
            if k in _FIN_SKIP:
                continue
            v[k] = _fin_walk(v[k], kind, live or k in _FIN_KEYS)
        return v
    if isinstance(v, list):
        return [_fin_walk(x, kind, live) for x in v]
    if isinstance(v, str) and live:
        return _finish_text(v, kind)
    return v


@app.after_request
def _finish_ai_answers(resp):
    """3.15: the AI's words, tidied on their way out (see _finish_text)."""
    try:
        if request.method != "POST" or resp.status_code != 200 or resp.mimetype != "application/json":
            return resp
        if resp.direct_passthrough or resp.headers.get("Content-Encoding") or resp.headers.get("X-Once"):
            return resp        # 4.9: a kept answer sent again was finished the first time
        kind = next((k for rx, k in _FIN_PATHS if rx.match(request.path or "")), None)
        if not kind:
            return resp
        data = json.loads(resp.get_data(as_text=True))
        data = _fin_walk(data, kind, False)
        resp.set_data(json.dumps(data, ensure_ascii=False))
    except Exception as e:
        logger.warning(f"Finishing an AI answer skipped: {e}")
    return resp


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
    "in Profile -> RamzAI's memory."
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


# 5.0: how long a piece may be, how much one answer holds, and the line that says "there is more to come"
WRITE_MAX_WORDS = 12000
WRITE_ONE_WORDS = 3000          # up to here a piece is one answer, as before 5.0
WRITE_PART_WORDS = 2400
WRITE_MORE_MARK = "\u27e6more\u27e7"
_WRITE_MORE = re.compile(r"\s*\u27e6\s*more\s*\u27e7\s*", re.I)


def _brief_year(v, this_year):
    """A year the person typed (1800 to this year), or None."""
    try:
        y = int(str(v).strip())
    except (TypeError, ValueError):
        return None
    return y if 1800 <= y <= this_year else None


def _write_more(text):
    """(the text without the "more to come" line, whether it was there)."""
    if not text or not _WRITE_MORE.search(text):
        return text, False
    return _WRITE_MORE.sub("\n\n", text).strip(), True


@app.route("/api/write", methods=["POST", "OPTIONS"])
@limiter.limit(AI_LIMIT, key_func=_account_or_ip_key)
@auth_required
@_once
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

    # 4.0: Evidence-Based Writing is Scholar's, and a draft costs POINTS_WRITE — taken before any call, given back on failure
    gate = _scholar_gate(request.current_user, "Automatic research writing", research=True)
    if gate:
        return gate
    try:
        charged_from = _log_sage_message_and_consume(request.current_user, POINTS_WRITE, research=True)
    except ConversionError as e:
        return _points_reply(e)
    outline = parse_outline(topic)

    # Citation style and how recent the sources must be (both optional). 5.0: a style beyond the four is made and checked
    # in author-date citations, then turned into what the person asked for (cite_plan)
    style, cite_plan = _cite_plan(data)
    CS = CITE_STYLES[style]
    # 5.0: the kinds of sources the person chose (when they didn't: everything but editorials, letters and datasets)
    kinds = [k for k in WRITE_KINDS if k in (data.get("kinds") or [])] or list(WRITE_KINDS)
    try:
        recent_years = int(data.get("recent_years") or 0)
    except (TypeError, ValueError):
        recent_years = 0
    recent_years = recent_years if 1 <= recent_years <= 100 else 0       # 5.0: any number of years (4.9: 5 or 10 only)
    note = {"text": ""}
    # 5.0: or the years the person chose themselves ("from 2010 to 2018"); and the length they asked for, which also
    # says how many papers the piece draws on (a long piece needs more than a short one)
    this_year = datetime.now(timezone.utc).year
    year_from, year_to = _brief_year(data.get("year_from"), this_year), _brief_year(data.get("year_to"), this_year)
    if year_from and year_to and year_from > year_to:
        year_from, year_to = year_to, year_from
    since = year_from or ((this_year - recent_years + 1) if recent_years else None)
    until = year_to if (year_to and year_to < this_year) else None
    if year_from or until:
        span_words = (f"from {year_from} to {until}" if (year_from and until) else f"from {year_from} on" if year_from else f"up to {until}")
    else:
        span_words = f"from the last {recent_years} years"
    try:
        length_words = int(data.get("length_words") or 0)
    except (TypeError, ValueError):
        length_words = 0
    length_words = max(150, min(WRITE_MAX_WORDS, length_words)) if length_words > 0 else 0
    n_sources = 8 if length_words <= 1500 else 12 if length_words <= 3000 else 16 if length_words <= 6000 else 20

    def _find(text, n=None, topic=True):
        """Real, downloadable papers for this request, cited in the chosen style — from the years the person asked
        for, others added (with a note) when too few came up."""
        n = n or n_sources
        look = paper_topic_search if topic else search_papers
        papers, diag = look(text, max_results=n, since=since, kinds=kinds, **({"until": until} if until else {}))
        papers = [p for p in papers if (p.get("kind") or "article") in kinds]
        if (since or until) and len(papers) < 3 and diag != "not a topic":
            older, _ = look(text, max_results=n, kinds=kinds)
            older = [p for p in older if (p.get("kind") or "article") in kinds]
            have = {p["url"] for p in papers}
            fresh = len(papers)
            papers = papers + [p for p in older if p["url"] not in have][: n - len(papers)]
            if len(papers) > fresh:
                note["text"] = (f"Only {fresh} paper{'' if fresh == 1 else 's'} {span_words} came up, "
                                f"so {'others' if until else 'older ones'} were added to support the writing properly.")
            if papers:
                diag = None
        return enrich_sources(papers, style), diag

    # The student's own brief (word count, sections, focus…) and a target
    # length, added to EVERY writing call below — including continuations —
    # by this local wrapper, so no path can miss them. They never override
    # the real-sources and exact-citation rules.
    reqs = str(data.get("requirements") or "").strip()[:2500]
    # 5.0: the person's own template or guide (read by /api/brief/template), or the name of one of ours
    template = str(data.get("template") or "").strip()[:4000]
    extra = ""
    if reqs:
        extra += ("\n\nTHE PERSON'S REQUIREMENTS (what they said before starting: who the piece is for, its purpose, its "
                  "readers, and what it must follow) \u2014 follow them exactly: structure, sections, focus, tone, level, "
                  "audience and any word count. They never override the rules above about using only the "
                  "listed real sources with their exact citations.\n<<<\n" + reqs + "\n>>>")
    if template:
        extra += ("\n\nTHE PERSON'S TEMPLATE OR GUIDE \u2014 follow its structure: use its headings, in its order, and its "
                  "instructions on style and layout. Where it asks for something the listed real sources can't support, write "
                  "what they do support under that heading and say plainly what is missing. It never overrides the rules above "
                  "about using only the listed real sources with their exact citations.\n<<<\n" + template + "\n>>>")
    # 5.0: the language the person chose; and kinds they named that aren't in the list
    extra += language_rule(data.get("language"))
    kinds_more = re.sub(r"\s+", " ", str(data.get("kinds_more") or "")).strip()[:200]
    if kinds_more:
        extra += ("\n\nTHE PERSON ALSO ASKED FOR THESE KINDS OF SOURCES: " + kinds_more + ". Use any of the listed sources that are of "
                  "those kinds; never claim a source is of a kind it isn't.")
    extra_fixed = extra
    if length_words and length_words <= WRITE_ONE_WORDS:
        extra += (f"\n\nLENGTH: write about {length_words} words in total (within 10%), not counting the reference list. "
                  "If the piece needs more room than one answer allows, stop at a natural break \u2014 it will be continued.")
    elif length_words:
        # 5.0: a long piece (a thesis chapter, a full report) is written in parts, each one answer long; the person
        # asks for the next part. The mark tells the app there is more to come.
        extra += (f"\n\nLENGTH: the whole piece is about {length_words} words (within 10%), not counting the reference list, "
                  f"written in parts. In THIS answer write at most about {WRITE_PART_WORDS} words: whole sections, ending at the "
                  "end of a section, never in the middle of a sentence. Plan the whole piece so its sections share the length "
                  "sensibly, and write only as much as the listed real sources support \u2014 never pad. If the piece is NOT "
                  f"finished at the end of this answer, write the single line {WRITE_MORE_MARK} after your last paragraph. When "
                  "the piece is complete (its last section written), don't write that line.")
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
            # 4.5: the papers the earlier part cites keep their in-text citations, and a paper found again keeps its key.
            # 4.5.1: a new paper whose first author shares a surname with one the earlier part cites would need the
            # earlier citations changed (2021a/2021b, initials, more names): it is left out, so every citation stays right
            prev = []
            given = {}       # 5.0: in a numbered style, the numbers the earlier parts gave (prev's index -> number)
            n_to_key = {}
            for r in (data.get("previous_refs") or [])[:40]:
                try:
                    if isinstance(r, dict) and str(r.get("akey") or "").startswith("(") and int(r.get("n") or 0) > 0:
                        n_to_key[int(r["n"])] = str(r["akey"])[:200]
                except (TypeError, ValueError):
                    pass
            if cite_plan and cite_plan["marks"]:
                previous_text = marks_back(previous_text, cite_plan["marks"], n_to_key)
            for r in (data.get("previous_refs") or [])[:40]:
                rkey = str(r.get("akey") or r.get("key") or "") if isinstance(r, dict) else ""
                if isinstance(r, dict) and rkey.startswith("(") and r.get("text"):
                    plain_ref = re.sub(r"[*_]", "", str(r["text"]))
                    key = rkey[:200]
                    try:
                        if cite_plan and cite_plan["marks"] and int(r.get("n") or 0) > 0:
                            given[len(prev)] = int(r["n"])
                    except (TypeError, ValueError):
                        pass
                    prev.append({"citation_key": key, "ref_text": str(r["text"])[:1200], "url": str(r.get("url") or "")[:600],
                                 "pdf": r.get("pdf") or None, "doi": _norm_doi(r.get("doi")), "title": plain_ref[:300],
                                 "abstract": None, "group": re.sub(r"[^a-z0-9:]", "", str(r.get("group") or "").lower())[:120] or _key_group(key),
                                 "ref_after": str(r.get("after") or "")[:80], "_sort": _sort_back(r.get("sort"), r["text"])})
            if prev:
                groups = {p["group"] for p in prev if p["group"]}
                fresh = []
                for c in cont_sources or []:
                    twin = next((p for p in prev if (c.get("doi") and c.get("doi") == p.get("doi")) or (c.get("url") and c.get("url") == p.get("url"))), None)
                    if twin is not None:
                        twin["abstract"] = c.get("abstract")         # the same paper: its key stays, its abstract helps the check
                        twin["title"] = c.get("title") or twin["title"]
                        continue
                    if (c.get("group") or _paper_group(c)) in groups:
                        continue
                    fresh.append(c)
                cont_sources = prev + fresh
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
            cont_text, cont_more = _write_more(cont_result["text"].strip())
            if not cont_text:
                raise ConversionError("Couldn't continue — try again")
            cont_text, cont_cited = keep_known_citations(cont_text, cont_sources, style)
            cont_text, cont_cited = _c34_write(cont_text, cont_sources, style, topic, note, started)
            if prev:
                cont_cited = set(cont_cited) | set(range(len(prev)))     # the earlier part still cites its papers
            if cite_plan and cite_plan["marks"]:
                cont_text = _strip_invented(cont_text)      # a number the model wrote itself is never one of ours
            return jsonify(restyle_result({
                "type": "result",
                "text": cont_text,
                "references": build_references_list(cont_sources, cont_cited), "references_all": bool(prev),
                "references_title": CS["heading"], "source_note": note["text"], "style": style,
                "truncated": cont_result["stop_reason"] == "max_tokens" or cont_more,
            }, cont_sources, cont_cited, cite_plan, given, language=data.get("language")))
        except ConversionError as e:
            _refund_sage_message(request.current_user, charged_from, POINTS_WRITE, research=True)
            return jsonify({"error": str(e)}), 502

    try:
        if len(outline) >= 2:
            # Outline mode: e.g. "5.1 Discussion / 5.1.1 General Awareness / 5.1.2 ..."
            # Cap section count so one request can't run unbounded.
            outline = outline[:14]
            combined_headings = "; ".join(s["heading"] for s in outline)
            if length_words > WRITE_ONE_WORDS:
                extra = extra_fixed          # an outline is written section by section: the "in parts" rule isn't for it

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
                sec_text = _write_more(sec_text)[0]
                sec["text"] = sec_text if len(sec_text) >= 25 else "(Couldn't generate this subsection — try again.)"
            cited_all = set()
            for sec in outline:
                sec["text"], c = keep_known_citations(sec["text"], sources, style)
                cited_all |= c
            if sources:
                cited_all = _c34_write_sections(outline, sources, style, topic, note, started)

            if cite_plan and cite_plan["marks"]:
                for sec in outline:
                    sec["text"] = _strip_invented(sec["text"])
            return jsonify(restyle_result({
                "type": "result",
                "sections": [
                    {"number": s["number"], "heading": s["heading"], "level": s["level"], "text": s["text"]}
                    for s in outline
                ],
                "references": build_references_list(sources, cited_all), "references_title": CS["heading"], "source_note": note["text"], "style": style,
                "title": next((s["heading"] for s in outline if s.get("level", 1) <= 1 and s.get("heading")), None),
            }, sources, cited_all, cite_plan, language=data.get("language")))

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
        raw_text, more_to_come = _write_more(result["text"].strip())
        truncated = result["stop_reason"] == "max_tokens" or more_to_come

        thinking, after_thinking = split_thinking(raw_text)
        is_clarify, text = split_clarify(after_thinking)
        memory_note, text = split_memory_note(text)
        noted = None
        if memory_note and memory_on and add_user_memory(user_id, memory_note):
            noted = memory_note[:300]

        if not text:
            raise ConversionError("Couldn't generate a reply — try rephrasing")
        if not is_clarify and len(text) < 15:
            raise ConversionError("Couldn't generate a reply — try rephrasing")

        if is_clarify:
            return jsonify({"type": "clarify", "text": text, "thinking": thinking, "noted": noted})

        text, cited = keep_known_citations(text, sources, style)
        text, cited = _c34_write(text, sources, style, topic, note, started)
        if cite_plan and cite_plan["marks"]:
            text = _strip_invented(text)
        return jsonify(restyle_result({
            "type": "result",
            "text": text,
            "thinking": thinking,
            "references": build_references_list(sources, cited), "references_title": CS["heading"], "source_note": note["text"], "style": style,
            "title": _topic_title(topic) if not attachment else None,
            "truncated": truncated, "noted": noted,
        }, sources, cited, cite_plan, language=data.get("language")))
    except ConversionError as e:
        _refund_sage_message(request.current_user, charged_from, POINTS_WRITE, research=True)
        return jsonify({"error": str(e)}), 502


# ---------- 4.0 Guided citing: statements with real references, one tap at a time ----------
GUIDE_STATEMENTS = 10          # suggestions in one batch
GUIDE_PAPERS = 12              # real papers searched for one batch


def _guide_json(raw):
    cleaned = re.sub(r"^```(?:json)?\s*|\s*```$", "", (raw or "").strip())
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        a, b = cleaned.find("{"), cleaned.rfind("}")
        if a >= 0 and b > a:
            try:
                return json.loads(cleaned[a:b + 1])
            except json.JSONDecodeError:
                return None
    return None


def _guide_clip(v, n):
    return re.sub(r"\s+", " ", str(v or "")).strip()[:n]


@app.route("/api/guide/outline", methods=["POST", "OPTIONS"])
@limiter.limit(AI_LIMIT, key_func=_account_or_ip_key)
@auth_required
@_once
def guide_outline():
    """Subtopics a researcher could build a piece of writing from, for one main topic."""
    data = request.get_json(silent=True) or {}
    topic = _guide_clip(data.get("topic"), 300)
    kind = _guide_clip(data.get("kind"), 40) or "essay"
    if not topic:
        return jsonify({"error": "Type the topic first."}), 400
    gate = _scholar_gate(request.current_user, "Manual research writing", research=True)
    if gate:
        return gate
    try:
        charged_from = _log_sage_message_and_consume(request.current_user, POINTS_REMY, in_research=True)
    except ConversionError as e:
        return _points_reply(e)
    system = ("You help a researcher plan a piece of academic writing. Reply with one JSON object only, no code fences: "
              '{"subtopics": [{"title": "...", "why": "..."}]}. Give 5 to 8 subtopics in the order the writing should run '
              "(background and definitions first, then the body, then gaps or implications). Each title is a heading of 2 to "
              "8 words in sentence case; each why is one sentence (at most 18 words) on what that part should establish. "
              "Write in the topic's language. Never invent studies, numbers or names.")
    try:
        context = _guide_clip(data.get("context"), 700)      # 5.0: who it is for, as the person said before starting
        lang = re.sub(r"[^\w\s\-’'()]", "", str(data.get("language") or "")).strip()[:40]
        raw = call_claude(system + (f" Write the titles and the whys in {lang}." if lang and lang.lower() not in ("english", "en") else ""),
                          f"Topic: {topic}\nKind of writing: {kind}" + (f"\nWritten for: {context}" if context else ""),
                          max_tokens=900, feature="guide")
    except ConversionError as e:
        _refund_sage_message(request.current_user, charged_from, POINTS_REMY, in_research=True)
        return jsonify({"error": str(e)}), 502
    obj = _guide_json(raw)
    subs = []
    for x in (obj or {}).get("subtopics") or []:
        if isinstance(x, dict) and _guide_clip(x.get("title"), 90):
            subs.append({"title": _guide_clip(x.get("title"), 90), "why": _guide_clip(x.get("why"), 200)})
    if not subs:
        _refund_sage_message(request.current_user, charged_from, POINTS_REMY, in_research=True)
        return jsonify({"error": "RamzAI couldn't outline that — try a few more words about the topic."}), 502
    return jsonify({"topic": topic, "subtopics": subs[:8]})


def _c45_pick_papers(papers, known):
    """The papers to offer, in order: one per first-author surname, and none whose surname a different paper already
    cited in the page has (known: [{key, group, doi, url}] from the app)."""
    used = {}
    for k in (known or [])[:80]:
        if isinstance(k, dict) and (k.get("key") or k.get("group")):
            grp = re.sub(r"[^a-z0-9:]", "", str(k.get("group") or "").lower())[:120] or _key_group(k.get("key"))
            if grp:
                used.setdefault(grp, set()).add(_norm_doi(k.get("doi")) or str(k.get("url") or "")[:600] or str(k.get("key") or "")[:200])
    picked, groups = [], set()
    for p in papers:
        grp = _paper_group(p)
        pid = p.get("doi") or p.get("url") or ""
        if grp and (grp in groups or (grp in used and pid not in used[grp])):
            continue
        groups.add(grp)
        picked.append(p)
    return picked


@app.route("/api/guide/statements", methods=["POST", "OPTIONS"])
@limiter.limit(AI_LIMIT, key_func=_account_or_ip_key)
@auth_required
@_once
def guide_statements():
    """About ten statements for one topic or subtopic, each grounded in one real, downloadable paper, with
    the in-text citation and the reference entry in the chosen style, and a paraphrase ready to insert."""
    data = request.get_json(silent=True) or {}
    topic = _guide_clip(data.get("topic"), 300)
    sub = _guide_clip(data.get("subtopic"), 200)
    # 5.0: any style (Vancouver, IEEE, or the person's own): the statements carry author-date citations, which the page
    # turns into numbers where the style numbers them; the references are written in the style asked
    style, cite_plan = _cite_plan(data)
    kinds = [k for k in WRITE_KINDS if k in (data.get("kinds") or [])] or list(WRITE_KINDS)
    lang = re.sub(r"[^\w\s\-’'()]", "", str(data.get("language") or "")).strip()[:40]
    since = data.get("recent_years")
    try:
        since = datetime.now(timezone.utc).year - int(since) + 1 if since else None   # "the last 10 years": this year and the 9 before
    except (TypeError, ValueError):
        since = None
    # 5.0: or the years the person chose themselves ("from 2010 to 2018")
    _this = datetime.now(timezone.utc).year
    _yf, _yt = _brief_year(data.get("year_from"), _this), _brief_year(data.get("year_to"), _this)
    if _yf and _yt and _yf > _yt:
        _yf, _yt = _yt, _yf
    since = _yf or since
    until_kw = {"until": _yt} if (_yt and _yt < _this) else {}
    exclude = set(str(x) for x in (data.get("exclude") or [])[:60] if x)
    if not (topic or sub):
        return jsonify({"error": "Type the topic first."}), 400
    gate = _scholar_gate(request.current_user, "Manual research writing", research=True)
    if gate:
        return gate
    query = f"{topic} {sub}".strip() if sub and topic and sub.lower() not in topic.lower() else (sub or topic)
    papers, diag = search_papers(query, max_results=GUIDE_PAPERS, since=since, kinds=kinds, **until_kw)
    papers = [p for p in papers if (p.get("kind") or "article") in kinds]
    if len(papers) < 3 and sub:                     # the subtopic alone was too narrow: widen to the topic
        more, _ = search_papers(topic or sub, max_results=GUIDE_PAPERS, since=since, kinds=kinds, **until_kw)
        more = [p for p in more if (p.get("kind") or "article") in kinds]
        seen = {p.get("doi") or p["title"] for p in papers}
        papers += [p for p in more if (p.get("doi") or p["title"]) not in seen]
    papers = [p for p in papers if (p.get("doi") or p.get("url") or "") not in exclude and p.get("abstract")][:GUIDE_PAPERS]
    if not papers:
        return jsonify({"error": "No open papers with a PDF and an abstract were found for that. Try other words, or a wider topic.",
                        "no_papers": True}), 404
    # 4.5.1: a paper whose first author shares a surname with one already cited in the page would need the page's own
    # citation changed (2021a/2021b, initials, more names), and two such papers here could leave an "a" with no "b":
    # each surname comes once, and never one the page already has (from a different paper), so every citation in the
    # page stays right as it is
    picked = _c45_pick_papers(papers, data.get("known"))
    if not picked:
        return jsonify({"error": "The papers found for that are by authors already cited in your page. Try other words, or a wider topic.",
                        "no_papers": True}), 404
    sources = enrich_sources(picked, style)
    try:
        charged_from = _log_sage_message_and_consume(request.current_user, POINTS_RESEARCH, research=True)
    except ConversionError as e:
        return _points_reply(e)
    heading = sub or topic
    system = ("You help a researcher write, one evidence-based statement at a time. You are given real papers, each with a "
              "number, its citation key and its abstract. Reply with one JSON object only, no code fences: "
              '{"statements": [{"paper": <number>, "text": "...", "paraphrase": "...", "kind": "..."}]}. '
              f"Write {GUIDE_STATEMENTS} statements a careful academic writer could put under the heading given, in the order "
              "that builds the argument: context and definitions first, then findings and evidence, then limits, gaps or "
              "implications. Each statement says ONE thing the abstract of its paper actually supports; never go beyond the "
              "abstract, never invent numbers, and quote a figure only when the abstract gives it. 18 to 40 words each, plain "
              + (f"academic {lang} (numbers in digits)" if lang and lang.lower() != "english" else "academic English (or the topic's language)")
              + ", no citation inside the text (the app adds it). paraphrase says the "
              "same thing in clearly different words and sentence shape, the version that will go into the document. kind is one "
              "of: context, definition, finding, method, evidence, limitation, gap, implication. Use at least 6 different papers; "
              "each paper at most twice. Use only the numbers given.")
    lines = []
    for i, src in enumerate(sources, 1):
        lines.append(f"[{i}] {src['citation_key']} {src['title']} ({src.get('venue') or 'paper'}, {src.get('year') or 'n.d.'})"
                     + (f"\n    Abstract: {src['abstract'][:900]}" if src.get("abstract") else ""))
    user = f"Heading: {heading}" + (f"\nMain topic: {topic}" if topic and sub else "") + "\n\nPapers:\n" + "\n".join(lines)
    try:
        raw = call_claude(system, user, max_tokens=3200, timeout=120, feature="guide")
    except ConversionError as e:
        _refund_sage_message(request.current_user, charged_from, POINTS_RESEARCH, research=True)
        return jsonify({"error": str(e)}), 502
    obj = _guide_json(raw)
    out = []
    kinds = {"context", "definition", "finding", "method", "evidence", "limitation", "gap", "implication"}
    for x in (obj or {}).get("statements") or []:
        if not isinstance(x, dict):
            continue
        try:
            n = int(x.get("paper"))
        except (TypeError, ValueError):
            continue
        if not 1 <= n <= len(sources):
            continue
        text, para = _guide_clip(x.get("text"), 600), _guide_clip(x.get("paraphrase"), 600)
        if len(text.split()) < 6:
            continue
        if len(para.split()) < 6 or para.lower() == text.lower():
            para = text
        src = sources[n - 1]
        ab = (src.get("title") or "") + " " + (src.get("abstract") or "")
        nums = _c45_numbers(text) + _c45_numbers(para)
        if nums and not _c45_numbers_in(nums, ab):
            continue                                  # a figure the abstract doesn't give
        if _c45_copied(para, src.get("abstract") or ""):
            continue                                  # the words that would go in are the abstract's own: not a paraphrase
        out.append({"id": f"{len(out) + 1}", "paper": n, "text": text, "paraphrase": para,
                    "kind": x.get("kind") if x.get("kind") in kinds else "finding",
                    "cite": src["citation_key"],
                    "ref": {"text": src["ref_text"], "url": src["url"], "pdf": src.get("pdf") or None, "doi": src.get("doi") or None,
                            "key": src["citation_key"], "title": src["title"], "year": src.get("year"), "authors": src.get("authors") or [],
                            "group": src.get("group") or "", "after": src.get("ref_after") or ""}})
        if len(out) >= GUIDE_STATEMENTS:
            break
    # each statement is checked against the papers by a second pass: it stays only when its own paper supports it
    checked = False
    if out:
        ev = [f"[P{i}] {s.get('title', '')} \u2014 {str(s.get('abstract') or '')[:900]}" for i, s in enumerate(sources, 1)]
        dec = _c34_check("guide", heading, ev, [o["paraphrase"] for o in out], "", time.time() + 40)   # the words that go in
        if dec and not dec["skip"]:
            checked = True
            out = [o for i, o in enumerate(out) if f"P{o['paper']}" in (dec["cite"].get(i) or [])]
            for i, o in enumerate(out):
                o["id"] = f"{i + 1}"
    if len(out) < 3:
        _refund_sage_message(request.current_user, charged_from, POINTS_RESEARCH, research=True)
        return jsonify({"error": "RamzAI couldn't build statements those papers fully support. Try again, or say the topic in other words."}), 502
    note = diag or ""
    marks = None
    if cite_plan:
        # the references in the style asked: Vancouver or IEEE exactly; the person's own style like their example (checked)
        marks = cite_plan["marks"]
        by_paper = {}
        for o in out:
            by_paper.setdefault(o["paper"], []).append(o)
        papers_used = sorted(by_paper)
        if cite_plan["marks"] and cite_plan["refs"] in NUM_STYLES:
            texts = {n: numbered_ref(sources[n - 1], cite_plan["refs"]) for n in papers_used}
        else:
            texts = {n: sources[n - 1]["ref_text"] for n in papers_used}
        if cite_plan["style"] == "other" and cite_plan.get("example"):
            laid = refs_like_example([(sources[n - 1], texts[n]) for n in papers_used], cite_plan["example"], cite_plan["name"])
            bad = 0
            for n, (t, ok) in zip(papers_used, laid):
                texts[n] = t
                bad += 0 if ok else 1
            if bad:
                note = (note + " " + f"{bad} reference{'s' if bad > 1 else ''} couldn't be laid out like your example and checked, so "
                        + ("it follows" if bad == 1 else "they follow") + " " + ("Vancouver" if cite_plan["marks"] else CITE_STYLES[cite_plan["refs"]]["name"]) + ".").strip()
        for o in out:
            o["ref"]["text"] = texts[o["paper"]]
            o["ref"]["akey"] = o["cite"]
            if cite_plan["marks"]:
                o["ref"]["after"] = ""
    return jsonify({"topic": topic, "subtopic": sub, "style": cite_plan["style"] if cite_plan else style,
                    "style_name": cite_plan["name"] if cite_plan else CITE_STYLES[style]["name"],
                    "references_title": ref_title_for(lang) or ("References" if marks else CITE_STYLES[style]["heading"]), "statements": out, "cite_marks": marks,
                    "papers": len(sources), "points": POINTS_RESEARCH, "source_note": note, "checked": checked})



# ======================= 4.6: Research (RamzAI) =======================
# The owner, 9 Nov 2026 (from the JustDone and Skolar pictures he sent): research that asks first (the years, the kinds of
# sources), shows its steps as they happen, and shows for every citation the sentence of the paper behind it
# ("Supports" / "Partly supports"), with an evidence table and the person's own library as sources.
# One request, answered one line of JSON at a time (application/x-ndjson) as each step really happens (and a heartbeat
# line every few seconds, so a person who leaves is noticed and given the point back). Nothing is shown as sourced unless
# whole sentences of a paper's own abstract (or of the person's own PDF) state it: the AI names a sentence by its label and
# the server copies it out, so every quote is word for word; every figure, in digits or in words, must be in the quoted
# sentences; a sentence its papers don't fully back is taken out; and so is one copied from any paper.
RS_KINDS = ("article", "review", "book", "preprint", "conference", "thesis", "report")
RS_YEARS = (0, 5, 10, 20)
RS_MAX_PAPERS = 10
RS_MIN_PAPERS = 3
RS_SECTIONS = (("short", "Short answer"), ("found", "What the studies found"), ("unsure", "Where they disagree or are unsure"))
# 5.0: a report written in another language has its three headings in that language too
RS_SECTIONS_LANG = {
    "french": ("Réponse courte", "Ce que montrent les études", "Divergences et incertitudes"),
    "spanish": ("Respuesta breve", "Lo que muestran los estudios", "Discrepancias e incertidumbres"),
    "portuguese": ("Resposta curta", "O que os estudos mostram", "Divergências e incertezas"),
    "german": ("Kurze Antwort", "Ergebnisse der Studien", "Widersprüche und Unsicherheiten"),
    "arabic": ("الإجابة المختصرة", "ما توصلت إليه الدراسات", "مواضع الاختلاف وعدم اليقين"),
    "kiswahili": ("Jibu fupi", "Tafiti zilichogundua", "Tofauti na mashaka"),
}
RS_SECTIONS_LANG["swahili"] = RS_SECTIONS_LANG["kiswahili"]
RS_KIND_WORDS = {"article": ("journal article", "journal articles"), "review": ("review", "reviews"),
                 "conference": ("conference paper", "conference papers"), "thesis": ("thesis", "theses"), "report": ("report", "reports"),
                 "book": ("book or chapter", "books or chapters"), "preprint": ("preprint", "preprints"), "mine": ("of your PDFs", "of your PDFs")}
_RS_ABBR = re.compile(r"(?:\b(?:e\.g|i\.e|et al|etc|vs|cf|approx|ca|fig|figs|no|nos|vol|pp?|eds?|dr|mr|mrs|ms|prof|st|jr|sr|inc|ltd|"
                      r"co|resp|u\.s|u\.k|n|sd|se|ci|or)|(?<![\w.])[A-Za-z])\.$", re.I)
_RS_END = re.compile(r"[.!?][\"”’')\]]*(?=\s+[\"“'(\[]?[A-Z0-9])")
_RS_WHOLE = re.compile(r"[.!?][\"”’')\]]*$")


def _rs_split(text, limit=60):
    """An abstract's whole sentences, word for word (spaces tidied). A label such as "Results:" or a short piece stays with
    the sentence after it. A piece that doesn't end as a sentence ends (cut off) or is too long to be one is left out:
    a quote is always a whole sentence of the paper."""
    text = re.sub(r"\s+", " ", str(text or "")).strip()
    if not text:
        return []
    out, start = [], 0
    for m in _RS_END.finditer(text):
        if _RS_ABBR.search(text[max(start, m.start() - 12):m.start() + 1]):
            continue
        out.append(text[start:m.end()].strip())
        start = m.end()
    if text[start:].strip():
        out.append(text[start:].strip())
    merged = []
    for s in out:
        if merged and len(re.findall(r"[A-Za-z]+", merged[-1])) < 4:          # "Background." or "Methods." joins what follows
            merged[-1] = merged[-1] + " " + s
        else:
            merged.append(s)
    return [s for s in merged if s and _RS_WHOLE.search(s) and len(s) <= 1200][:limit]


# ---- numbers, in digits or in words
_RS_UNITS = {"zero": 0, "one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10,
             "eleven": 11, "twelve": 12, "thirteen": 13, "fourteen": 14, "fifteen": 15, "sixteen": 16, "seventeen": 17, "eighteen": 18,
             "nineteen": 19}
_RS_TENS = {"twenty": 20, "thirty": 30, "forty": 40, "fifty": 50, "sixty": 60, "seventy": 70, "eighty": 80, "ninety": 90}
_RS_W = "|".join(sorted(list(_RS_UNITS) + list(_RS_TENS), key=len, reverse=True))
_RS_NUMWORD = re.compile(r"\b(?:" + _RS_W + r")(?:[\s-]+(?:and[\s-]+)?(?:" + _RS_W + r"|hundred|thousand|million))*\b", re.I)


def _rs_wordnum(phrase):
    """ "three hundred and twelve" -> 312; "forty" -> 40; None when it isn't a number."""
    total, cur, seen = 0, 0, False
    for w in re.findall(r"[a-z]+", phrase.lower()):
        if w in _RS_UNITS:
            cur += _RS_UNITS[w]; seen = True
        elif w in _RS_TENS:
            cur += _RS_TENS[w]; seen = True
        elif w == "hundred":
            cur = (cur or 1) * 100; seen = True
        elif w in ("thousand", "million"):
            total += (cur or 1) * (1000 if w == "thousand" else 1000000); cur = 0; seen = True
        elif w != "and":
            return None
    return total + cur if seen else None


def _rs_digits(text, strict, papers_ok=True, counts=None):
    """The text with numbers written in words as digits ("thirty percent" -> "30 percent", "twice" -> "2 times").
    strict (the report's own sentence): "one" stays a word unless it counts something ("one school", "one year"); when
    papers_ok, a count of the papers themselves stays out of the figures ("two studies", "three of the four trials",
    "2 studies", "Of the four trials, three found…") and goes into counts as (value, of_all) — of_all when it counts
    the papers read ("of the four studies"), not the papers that found something. Both sides turn an ordinal into
    digits only as a label ("sixth-grade" is Grade 6; "the second school" is not 2 schools)."""
    t = re.sub(r"\btwice\b", "2 times", str(text or ""), flags=re.I)
    t = re.sub(r"\bthrice\b", "3 times", t, flags=re.I)
    for rx, mark in _RS_FRAC_MARKS:
        t = rx.sub(mark, t)
    keep = counts if counts is not None else []
    of_all = lambda before: bool(re.search(r"\b(?:of|among|across)\s+(?:(?:the|these|those|all)\s+)?$", before, re.I))
    if strict and papers_ok:
        def lead(m):
            keep.append((_rs_wordnum(m.group(2)) if not m.group(2).isdigit() else int(m.group(2)), False))
            return m.group(1) + "some"
        t = _RS_OF_PAPERS.sub(lead, t)

    def one(m):
        phrase = m.group(0)
        v = _rs_wordnum(phrase)
        if v is None:
            return phrase
        if strict and v <= 10 and not re.search(r"hundred|thousand|million", phrase, re.I):
            after = t[m.end():m.end() + 60]
            if papers_ok and _RS_PAPER_NOUN.match(after):
                keep.append((v, of_all(t[max(0, m.start() - 24):m.start()])))
                return phrase                    # "two studies", "three of the papers"
            if v == 1 and phrase.lower() == "one" and not _RS_COUNTED.match(after) \
                    and not re.match(r"\s*(?:%|percent\b|per cent\b|percentage\b)", after, re.I):
                return phrase                    # "one of the", "one reason" — a word, not a figure
        return str(v)
    t = _RS_NUMWORD.sub(one, t)
    t = _RS_ORD_LABEL.sub(lambda m: str(_RS_ORDINALS[m.group(1).lower()]), t)
    if strict and papers_ok:
        def dig(m):
            keep.append((int(m.group(0)), of_all(t[max(0, m.start() - 24):m.start()])))
            return "some"
        t = re.sub(r"(?<![\w.,\-])(?:10|[1-9])(?=" + _RS_PAPER_NOUN.pattern + ")", dig, t, flags=re.I)
    return t


_RS_DIGITS = re.compile(r"(?<![\w.])(\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d*\.\d+|\d+)(?![\w])")
_RS_ORDINALS = {"first": 1, "second": 2, "third": 3, "fourth": 4, "fifth": 5, "sixth": 6, "seventh": 7, "eighth": 8, "ninth": 9, "tenth": 10,
                "eleventh": 11, "twelfth": 12}
_RS_ORD_LABEL = re.compile(r"\b(" + "|".join(_RS_ORDINALS) + r")(?=[\s-]+(?:grade|graders?|year|form|standard|class|trimester|semester|"
                           r"wave|round|cycle|term|phase)\b)", re.I)
# a count of the papers themselves ("two studies", "one trial") is the writer's, not a figure of a paper — unless the
# quotes count papers too (a review's "Of 24 studies, 18 found gains"); "two study arms" counts something in a paper
_RS_PAPER_WORD = (r"(?:stud(?:y|ies)|papers?|trials?|reviews?|articles?|authors?|sources?|reports?)\b"
                  r"(?!\s+(?:arms?|groups?|sites?|areas?|periods?|participants|populations?|samples?|schools?|districts?|regions?|"
                  r"countries|conditions?|days?|weeks?|months?|years?|sessions?|visits?|centres?|centers?|clinics?|hospitals?|"
                  r"villages?|communities)\b)")
_RS_NOT_PREP = r"(?!(?:in|per|of|from|for|with|at|by|on|to|across|than|among|over|under|about|during|after|before|and|or)\b)"
_RS_SMALL = r"(?:one|two|three|four|five|six|seven|eight|nine|ten|10|[1-9])"
_RS_PAPER_NOUN = re.compile(r"\s+(?:of\s+(?:(?:the|these|those|all)\s+)?)?(?:" + _RS_NOT_PREP + r"[\w-]+\s+){0,2}?" + _RS_PAPER_WORD, re.I)
# "Of the four trials, three found gains": the count right after the comma counts those papers too
_RS_OF_PAPERS = re.compile(r"(\b(?:of|among|across|in)\s+(?:(?:the|these|those|all)\s+)?" + _RS_SMALL + r"\s+(?:" + _RS_NOT_PREP + r"[\w-]+\s+){0,2}?"
                           + _RS_PAPER_WORD + r"[^,;.]{0,60}?,\s*)" + r"(?<![\w.])(" + _RS_SMALL + r")"
                           r"(?=\s+(?:of\s+(?:them|these|those)\b|(?:also\s+|each\s+|all\s+)?(?:found|find|finds|showed|show|shows|reported|"
                           r"report|reports|saw|see|sees|noted|note|notes|observed|observe|observes|suggested|suggest|suggests|concluded|"
                           r"conclude|concludes|measured|measure|measures|looked|look|looks|tested|test|tests|compared|compare|compares|"
                           r"focused|focus|focuses|studied|described|describe|describes|linked|link|links|pointed|point|points|agreed|"
                           r"agree|agrees|differed)\b))", re.I)
# a fraction in words stands as one sign, so its numbers ("two" of "two-thirds") are never read as figures
_RS_FRAC_MARKS = tuple((re.compile(rx, re.I), mark) for rx, mark in (
    (r"\b(?:one|a)[\s-]+half\b", "½"), (r"\btwo[\s-]+thirds\b", "⅔"), (r"\b(?:one|a)[\s-]+third\b", "⅓"),
    (r"\bthree[\s-]+quarters\b", "¾"), (r"\b(?:one|a)[\s-]+quarter\b", "¼"), (r"\b(?:one|a)[\s-]+fifth\b", "⅕"),
    (r"\b(?:two|three|four)[\s-]+fifths\b", "⅘"), (r"\b(?:one|a)[\s-]+tenth\b", "⅒")))
# fractions in words, and the percentages that are the same
_RS_FRACS = ((re.compile(r"\b(?:(?:one|a)[\s-]+half|half\s+(?:of|the)|halves|halved)\b", re.I), "1/2", (50,)),
             (re.compile(r"\b(?:one|a)[\s-]+third\b|\bthird\s+of\b", re.I), "1/3", (33, 33.3, 33.33, 34)),
             (re.compile(r"\btwo[\s-]+thirds\b", re.I), "2/3", (66, 66.7, 66.67, 67)),
             (re.compile(r"\b(?:one|a)[\s-]+quarter\b|\bquarter\s+of\b", re.I), "1/4", (25,)),
             (re.compile(r"\bthree[\s-]+quarters\b", re.I), "3/4", (75,)),
             (re.compile(r"\b(?:one|a)[\s-]+fifth\b", re.I), "1/5", (20,)),
             (re.compile(r"\b(?:one|a)[\s-]+tenth\b", re.I), "1/10", (10,)),
             (re.compile(r"\bdoubled\b", re.I), "x2", ()), (re.compile(r"\btripled\b", re.I), "x3", ()))
_RS_COUNTS_PAPERS = re.compile(r"(?<![\w.,\-])(?!(?:1[5-9]|20)\d\d\b)(?:\d[\d,]*|" + _RS_W + r")\s+(?:" + _RS_NOT_PREP + r"[\w-]+\s+){0,2}?"
                               r"(?:stud(?:y|ies)|papers?|trials?|reviews?|articles?|reports?|sources?)\b", re.I)


def _rs_values(text):
    """The numbers in a text (".45" and "0.45" are one value; 1,200 and 1200 too)."""
    out = set()
    for m in _RS_DIGITS.finditer(text):
        try:
            out.add(round(float(m.group(1).replace(",", "")), 6))
        except ValueError:
            pass
    return out


def _rs_figures_ok(sentence, quoted, n_cited=None, n_read=None):
    """Every figure the report's sentence states (in digits or in words) is in the quoted sentences, as the same value
    and counting the same thing; percentage points are never a percentage; a fraction ("half", "a third") is the same
    fraction or its percentage there. A count of the papers themselves ("two studies found…") is the writer's own,
    unless the quotes count papers too — then it is checked like any figure; when n_cited is given, it can't be more
    than the papers the sentence cites (or, counting all the papers, "of the five studies", than the papers read)."""
    q = _rs_digits(quoted, False)
    own = not _RS_COUNTS_PAPERS.search(str(quoted or ""))
    counts = []
    s = _rs_digits(sentence, True, papers_ok=own, counts=counts)
    if n_cited is not None and any(not v or v > (max(n_cited, n_read or 0) if of_all else n_cited) for v, of_all in counts):
        return False
    if not _rs_values(s) <= _rs_values(q):
        return False
    q_vals = _rs_values(q)
    for rx, name, pcts in _RS_FRACS:
        if rx.search(sentence) and not rx.search(quoted) and not any(v in q_vals for v in pcts):
            return False
    if re.search(r"percentage[- ]points?|\bpp\b", s, re.I) and not re.search(r"percentage[- ]points?|\bpoints?\b|\bpp\b", q, re.I):
        return False
    nums = _c45_numbers(s)
    return not nums or _c45_numbers_in(nums, q)


# ---- design and who it studied, from the abstract's sentences about the study itself
_RS_DESIGNS = (
    (re.compile(r"\bmeta-?analys[ie]s\b", re.I), "Meta-analysis"),
    (re.compile(r"\bsystematic (?:literature )?review\b", re.I), "Systematic review"),
    (re.compile(r"\bscoping review\b", re.I), "Scoping review"),
    (re.compile(r"\brandomi[sz]ed (?:controlled |clinical |cluster |field )?trials?\b|\brandomly (?:assigned|allocated)\b|\bRCT\b", re.I), "Randomised trial"),
    (re.compile(r"\bquasi-?experiment", re.I), "Quasi-experimental"),
    (re.compile(r"\bmixed[- ]methods?\b", re.I), "Mixed methods"),
    (re.compile(r"\blongitudinal\b|\bcohort stud(?:y|ies)\b|\bprospective (?:cohort|study)\b|\bfollowed (?:up )?for\b", re.I), "Longitudinal study"),
    (re.compile(r"\bcross-?sectional\b", re.I), "Cross-sectional study"),
    (re.compile(r"\bcase stud(?:y|ies)\b", re.I), "Case study"),
    (re.compile(r"\bsurvey(?:ed|s)?\b|\bquestionnaires?\b", re.I), "Survey"),
    (re.compile(r"\binterview(?:ed|s)?\b|\bfocus groups?\b|\bqualitative\b|\bthematic analysis\b", re.I), "Qualitative"),
    (re.compile(r"\bexperiment(?:s|al)?\b", re.I), "Experiment"),
)
# a sentence about this study (not about other studies, or what research should do next)
_RS_SELF = re.compile(r"\b(?:we|our|us|this (?:study|paper|article|review|meta-?analysis|trial|survey|research|report|analysis|work|chapter|book|project)|"
                      r"the (?:present|current) (?:study|review|research|paper)|a total of|participants|respondents|data (?:were|was) (?:collected|drawn|gathered)|"
                      r"(?:were|was)\s+(?:randomly\s+)?(?:assigned|allocated|recruited|surveyed|interviewed|enrolled|selected|sampled|followed|tested|"
                      r"compared|observed|included|analy[sz]ed|asked|invited))\b", re.I)
# "unlike previous randomised trials, we…": what other studies did is taken out before the design is read
_RS_OTHERS = re.compile(r"\b(?:unlike|like|as in|following|in contrast to|compared (?:with|to)|building on|extending)\s+(?:the\s+)?(?:previous|prior|earlier|other|most|many|existing|past|recent)\b[^,.;:]*", re.I)
_RS_NOT_SELF = re.compile(r"\b(?:needed|should be|future|further research|more research|recommend\w*|warrant\w*|call(?:s|ed)? for|will|"
                          r"plan(?:s|ned)?\s+to|aims?\s+to\s+(?:recruit|enrol|enroll|assign|randomi[sz]e)|to be (?:recruited|enrolled|randomi[sz]ed|assigned))\b|"
                          r"\b(?:should|would|could|might)\s+(?:[\w-]+\s+){0,2}?(?:stud|research|work|trial)", re.I)
# a clause about other studies ("earlier randomised trials had found gains") says nothing of this study's design
_RS_OTHER_CLAUSE = re.compile(r"\b(?:previous|prior|earlier|existing|other|past|recent|many|most|some|several|few)\s+(?:[\w-]+\s+){0,2}?"
                              r"(?:stud(?:y|ies)|research|work|trials?|reviews?|surveys?|experiments?|literature|evidence)\b", re.I)
_RS_WHO = ("participants|students|pupils|learners|children|adolescents|youths?|adults|teachers|educators|lecturers|patients|"
           "respondents|women|men|girls|boys|mothers|fathers|parents|nurses|doctors|physicians|clinicians|schools|classrooms|"
           "households|families|farmers|employees|workers|residents|undergraduates|graduates|individuals|people|persons|subjects|"
           "volunteers|cases|infants|newborns|babies|smallholders|firms|companies|countries|communities|villages|districts|"
           "hospitals|clinics|universities|colleges|studies|trials|articles|papers")
# "one" counts something when a counted thing follows it ("one school", "one rural school", "one year")
_RS_COUNTED = re.compile(r"\s+(?:[\w-]+\s+)?(?:" + _RS_WHO + r"|participant|student|pupil|learner|child|adolescent|youth|adult|teacher|"
                         r"educator|lecturer|patient|respondent|woman|man|girl|boy|mother|father|parent|nurse|doctor|physician|clinician|"
                         r"school|classroom|household|family|farmer|employee|worker|resident|undergraduate|graduate|individual|person|"
                         r"subject|volunteer|case|infant|newborn|baby|smallholder|firm|company|country|community|village|district|"
                         r"hospital|clinic|university|college|weeks?|months?|years?|days?|hours?|minutes?|sessions?|lessons?|terms?|"
                         r"semesters?|grades?|groups?|arms?|sites?|times?|classes|class|regions?|provinces?|states?)\b", re.I)
_RS_N = re.compile(r"(?<![\w.,\-])(?:n\s*=\s*(\d{1,3}(?:,\d{3})+|\d+)(?![\w.,])|(\d{1,3}(?:,\d{3})+|\d{2,7})\s+((?:[A-Za-z0-9-]+\s+){0,2}?(?:" + _RS_WHO + r"))\b)", re.I)
_RS_ROUGH = re.compile(r"(?:more|less|fewer|greater|higher|lower)\s+than\s+$|(?:over|under|about|around|nearly|almost|roughly|approximately|up to|at least|at most|some|each|per|every)\s+$", re.I)


def _rs_own_part(s):
    """The part of a sentence about this study: clauses about other studies left out."""
    s = _RS_OTHERS.sub(" ", s)
    parts = re.split(r"[;:]|,(?!\d{3})|\s(?:but|whereas|while|although|though|however|yet)\s", s)
    return " ".join(x for x in parts if x.strip() and not _RS_OTHER_CLAUSE.search(x)).strip()


def _rs_self(sents):
    """The sentences about this study itself (only their own part), not about other studies or about plans."""
    out = []
    for s in sents:
        own = _rs_own_part(s)
        if own and _RS_SELF.search(own) and not _RS_NOT_SELF.search(own):
            out.append(own)
    return out


def _rs_design(sents, kind=None):
    for s in _rs_self(sents):
        for rx, name in _RS_DESIGNS:
            if rx.search(s):
                return name
    return "Review" if kind == "review" else None


def _rs_people(sents):
    """Who the study looked at, in its abstract's own words ("312 pupils", "24 studies"), from a sentence about the study
    itself; never a rough number ("more than 50 pupils"), a label ("Grade 10 pupils") or part of a name ("COVID-19")."""
    for sent in _rs_self(sents):
        if len(re.findall(r"\bn\s*=\s*\d", sent, re.I)) > 1:
            continue                        # the sizes of several groups: the whole sample isn't one number here
        for m in _RS_N.finditer(sent):
            if m.group(1):
                return f"{m.group(1)} participants"
            n, what = m.group(2), m.group(3)
            if re.fullmatch(r"(?:1[5-9]|20)\d{2}", n):
                continue
            if set(what.lower().split()[:-1]) & {"percent", "per", "of", "in", "than", "more", "less", "fewer", "times", "out", "from", "to",
                                                 "and", "or", "with", "at", "for", "by", "the", "a", "an"}:
                continue
            before = sent[max(0, m.start() - 24):m.start()]
            if _RS_ROUGH.search(before) or _C45_LABEL.search(before):
                continue
            return f"{n} {what}".strip()
    return None


def _rs_cite_short(key):
    k = str(key or "").strip()
    return k[1:-1] if k.startswith("(") and k.endswith(")") else k


# ---- the person's library
_RS_INITIALS = re.compile(r"^(?:[A-ZÀ-ɏ]\.\s*(?:-\s*)?)+$")


_RS_BODY = re.compile(r"\b(?:ministry|department|organi[sz]ation|agency|council|institute|university|college|commission|"
                      r"office|service|association|society|board|authority|centre|center|committee|foundation|programme|program|"
                      r"government|unesco|unicef|united nations|world bank|world health|bureau|secretariat|directorate)\b", re.I)
_RS_PARTICLE = r"(?:van|von|der|den|de|del|della|da|di|du|la|le|bin|ibn|al|el|dos|das)"
_RS_PERSONLIKE = re.compile(r"^(?:" + _RS_PARTICLE + r"\s+)*[A-ZÀ-ɏ][\w'’\-]+(?:\s+(?:" + _RS_PARTICLE + r"\s+)*[A-ZÀ-ɏ][\w'’\-.]*){1,3}$")
# "Mensah K", "Mensah K.", "van der Berg JA" (surname, then initials); never "John Smith II"
_RS_VANCOUVER = re.compile(r"^((?:[a-z]{1,4}\s+)*[A-ZÀ-ɏ][\w'’\-]+(?:\s+(?:[a-z]{1,4}\s+)*[A-ZÀ-ɏ][\w'’\-]+){0,2})\s+((?:[A-Z]\.?\s?-?){1,3})$")
_RS_SUFFIX = re.compile(r"^(?:jr|sr|ii|iii|iv)\.?$", re.I)
# two-letter surnames written in capitals ("Wei LI") that are seldom initials
_RS_SHORT_SURNAMES = {"LI", "WU", "XU", "HU", "YU", "LU", "NG"}


def _rs_vancouver(a, caps_mode=None):
    """(surname, initials) when a reads as "Surname I" / "Surname I." — None when it's a name with its surname in
    capitals ("Kofi ADU", "Wei LI"). caps_mode: what the other authors of the list show ("initials" or "surname")."""
    m = _RS_VANCOUVER.match(a.strip())
    if not m:
        return None
    raw, letters = m.group(2), re.findall(r"[A-Z]", m.group(2))
    if re.fullmatch(r"(?:II|III|IV)", raw.strip()):
        return None
    if "." not in raw and len(letters) >= 2 and (caps_mode == "surname" or (caps_mode is None and (len(letters) >= 3 or "".join(letters) in _RS_SHORT_SURNAMES))):
        return None
    return m.group(1), letters


def _rs_authors(raw):
    """Authors as a person typed them -> one name per author, in a form the reference builder reads:
    "Mensah, K., Boateng, A., & Owusu, Y." / "Kofi Mensah, Adwoa Boateng" / "Mensah, Kofi; Boateng, Adwoa" /
    "Mensah K, Boateng AB" / "Mensah K., Boateng A." / "Kofi ADU" (-> "Kofi Adu"); "Mensah, K., Jr." keeps its suffix;
    a body's name stays whole ("Food and Agriculture Organization", "Ministry of Education, Ghana"); an affiliation
    typed after the people ("…, University of Ghana") is left out."""
    text = re.sub(r"\s*\([^()]*\)", "", str(raw or ""))                       # "(University of Ghana)"
    text = re.sub(r"(?<=[A-Za-zÀ-ɏ.])[\d*†‡§¶]+(?=\s*(?:,|;|&|\band\b|$))", "", text)   # "Mensah1", "Boateng*"
    text = re.sub(r"\s+", " ", text).strip(" ,;")
    if not text:
        return []
    initials = bool(re.search(r"(?:^|[\s,])[A-Z]\.(?=[\s,;&-]|$)", text))
    pieces = [x.strip() for x in re.split(r"[;,&]|\band\b", text) if x.strip()]
    if _RS_BODY.search(text) and not initials and not any(_rs_vancouver(x) for x in pieces):
        parts = [x.strip() for x in re.split(r"\s*;\s*|,", text) if x.strip()]
        people = [x for x in parts if not _RS_BODY.search(x) and _RS_PERSONLIKE.match(x) and not re.search(r"\b(?:and|of|for|the)\b", x)]
        if not people:
            return [text[:200]]                # one body
        text = ", ".join(x for x in parts if not _RS_BODY.search(x))
    out = []
    for chunk in re.split(r"\s*;\s*|\s*&\s*|\s+and\s+", text):
        chunk = chunk.strip(" ,")
        if not chunk:
            continue
        parts = [p.strip() for p in chunk.split(",") if p.strip()]
        names, suffix = [], {}
        for p_ in parts:                       # "Jr." goes with the name before it
            if _RS_SUFFIX.match(p_) and names:
                suffix[len(names) - 1] = p_.rstrip(".") + "."
            else:
                names.append(p_)
        parts = names
        if len(parts) == 1:
            got = [parts[0]]
        elif all(_RS_INITIALS.match(p_) for p_ in parts[1::2]) and len(parts) % 2 == 0:
            got = [f"{parts[i]}, {parts[i + 1]}" for i in range(0, len(parts), 2)]       # "Surname, I., Surname, I."
            suffix = {i // 2: v for i, v in suffix.items()}
        elif len(parts) % 2 == 0 and all(len(p_.split()) == 1 for p_ in parts):
            got = [f"{parts[i]}, {parts[i + 1]}" for i in range(0, len(parts), 2)]       # "Mensah, Kofi, Boateng, Adwoa"
            suffix = {i // 2: v for i, v in suffix.items()}
        elif len(parts) == 2 and len(parts[0].split()) == 1:
            got = [f"{parts[0]}, {parts[1]}"]                                            # "Surname, Given"
            suffix = {0: v for v in suffix.values()}
        else:
            got = parts                                                                  # "Kofi Mensah, Adwoa Boateng"
        out.extend(f"{g}, {suffix[i]}" if i in suffix else g for i, g in enumerate(got))
    # "Mensah KA" or "Kofi ADU"? The other names of the list decide; on their own, three capitals are a surname
    modes = set()
    for a in out:
        m = _RS_VANCOUVER.match(a.strip())
        if m and "," not in a:
            raw_i = m.group(2)
            caps = "".join(re.findall(r"[A-Z]", raw_i))
            if "." in raw_i or len(caps) == 1:
                modes.add("initials")               # "Mensah K", "Mensah K."
            elif len(caps) >= 3 or caps in _RS_SHORT_SURNAMES:
                modes.add("surname")                # "Kofi ADU", "Wei LI"
    mode = modes.pop() if len(modes) == 1 else None
    fixed = []
    for a in out:
        v = _rs_vancouver(a, mode) if "," not in a else None
        if v:
            fixed.append(f"{v[0]}, " + " ".join(c + "." for c in v[1]))
        else:
            fixed.append(" ".join(w.title() if len(w) >= 2 and w.isalpha() and w.isupper() and not re.fullmatch(r"II|III|IV", w) else w for w in a.split()))
    return [a[:120] for a in fixed][:20]


def _rs_biblio(b):
    if not isinstance(b, dict):
        return {}
    return {k: str(b.get(k))[:20] for k in ("volume", "issue", "first_page", "last_page") if b.get(k) and re.fullmatch(r"[\w.\-–]{1,20}", str(b.get(k)))}


def _rs_library(raw):
    """What the person sent from their library: papers they saved, and their own PDFs' first pages. Checked and capped."""
    out = []
    for it in (raw if isinstance(raw, list) else [])[:25]:
        if not isinstance(it, dict):
            continue
        kind = it.get("kind")
        title = re.sub(r"\s+", " ", str(it.get("title") or "")).strip()[:300]
        if not title:
            continue
        year = it.get("year")
        try:
            year = int(str(year)[:4]) if year else None
        except (TypeError, ValueError):
            year = None
        if year and not 1500 <= year <= datetime.now(timezone.utc).year + 1:
            year = None
        if kind == "saved":
            abstract = re.sub(r"\s+", " ", str(it.get("abstract") or "")).strip()[:9000]
            if len(abstract) < 120:
                continue
            authors = [re.sub(r"\s+", " ", str(a)).strip()[:120] for a in (it.get("authors") or []) if str(a or "").strip()][:40]
            url = str(it.get("url") or "")[:600]
            pdf = str(it.get("pdf") or "")[:600]
            was = it.get("was") if it.get("was") in RS_KINDS else "article"
            out.append({"title": title, "authors": authors or None, "year": year, "venue": (str(it.get("venue") or "")[:200] or None),
                        "doi": _norm_doi(it.get("doi")), "pdf": pdf if pdf.startswith("http") else None, "url": url if url.startswith("http") else None,
                        "abstract": abstract, "abstract_full": abstract, "kind": was, "preprint": was == "preprint", "lib": "saved",
                        "cites": 0, "biblio": _rs_biblio(it.get("biblio"))})
        elif kind == "mine":
            text = re.sub(r"[ \t]+", " ", str(it.get("text") or "")).strip()[:12000]
            if len(text) < 300:
                continue
            out.append({"title": title, "authors": _rs_authors(it.get("authors")) or None, "year": year, "venue": None, "doi": None,
                        "pdf": None, "url": None, "abstract": text, "abstract_full": text, "kind": "mine", "preprint": False, "lib": "mine",
                        "cites": 0, "biblio": {}})
    return out


# a protocol describes a study still to be done: it has no findings to quote
_RS_PROTOCOL = re.compile(r"\b(?:study|trial|review|research|evaluation)\s+protocol\b|(?:^|[:\-–—]\s*)(?:(?:a|the)\s+)?protocol\b(?:\s+(?:of|for)\b|\s*$)|"
                          r"\bprotocol\s+(?:of|for)\s+(?:a|an|the)\b|\brationale\s+and\s+(?:study\s+)?design\b|\bdesign\s+and\s+rationale\b|"
                          r"\bstudy\s+design\s+and\s+methods\b|\bplanned\s+(?:[\w-]+\s+){0,3}?(?:trial|study|evaluation|review)\b|"
                          r"\b(?:design|methods?)\s+(?:of|for)\s+(?:a|an|the)\s+(?:[\w-]+\s+){0,4}?(?:trial|study|evaluation)\b", re.I)
_RS_PROTOCOL_TEXT = re.compile(r"^\W*(?:(?:background|introduction|objectives?|aims?)\W+)?(?:this|the|a|our)\s+(?:study\s+|trial\s+|review\s+)?protocol\b|"
                               r"\b(?:this|the)\s+(?:study\s+|trial\s+|review\s+)?protocol\s+(?:describes|outlines|presents|reports|sets out)\b|"
                               r"\bwe\s+(?:describe|present|outline|report)\s+(?:the|a|our)\s+(?:study\s+|trial\s+)?protocol\b|"
                               r"\b(?:this|the|our)\s+(?:paper|article|report)\s+(?:describes|outlines|presents|sets out)\s+the\s+(?:design|methods?|rationale)\b|"
                               r"^\W*(?:[^.!?]{0,300}[.!?]\s+)?(?:we|this\s+(?:study|trial|evaluation))\s+will\s+(?:randomly\s+)?(?:assign|recruit|enrol|enroll|"
                               r"conduct|evaluate|test|compare|examine|investigate|measure|follow|allocate|collect|interview|survey)\b", re.I)
# a sentence about what a study will do is a plan, not a finding
_RS_PLAN = re.compile(r"\b(?:we|they|it|this|these|that|which|who|pupils|students|participants|data|results|findings|outcomes|study|trial|review)\s+will\b|"
                      r"\bwill\s+(?:be|not|also|then|need|help|have|allow|show|provide|inform|examine|assess|test|compare|explore|investigate|"
                      r"evaluate|measure|determine|use|follow|recruit|assign|enrol|enroll|conduct|collect)\b|\b(?:is|are)\s+planned\b", re.I)


def _rs_pick(found, library, question, kinds, since, until=None):
    """The papers the research reads: from the library first (when they're about the question and fit the years and
    kinds; a saved paper the search also found comes with the search's fuller record), then the search's best, up to
    RS_MAX_PAPERS; each with whole sentences of its own abstract to stand on (never an abstract lent by another version)."""
    terms = _paper_terms(question)

    def text_of(p):
        return p.get("abstract_full") or p.get("abstract") or ""

    def about(p, need):
        hay = ((p.get("title") or "") + " " + text_of(p)[:4000]).lower()
        return sum(1 for t in terms if t in hay) >= need

    def usable(p):
        return (not p.get("abstract_lent") and len(text_of(p)) >= 200 and len(_rs_split(text_of(p))) >= 2
                and not _RS_PROTOCOL.search(p.get("title") or "") and not _RS_PROTOCOL_TEXT.search(text_of(p)[:400]))

    need = 2 if len(terms) >= 3 else 1
    picked = []
    for p in library:
        if p["kind"] != "mine" and p["kind"] not in kinds:
            continue
        if since and p.get("year") and p["year"] < since:
            continue
        if until and p.get("year") and p["year"] > until:
            continue
        if not about(p, need) or len(picked) >= 5:
            continue
        if p["kind"] == "mine" and len([x for x in picked if x["kind"] == "mine"]) >= 3:
            continue
        if p["lib"] == "saved":
            twin = next((q for q in found if _same_work(p, q) and usable(q)), None)
            if twin:
                p = dict(twin, kind=twin.get("kind") or p["kind"], lib="saved")
        picked.append(p)
    for p in found:
        if len(picked) >= RS_MAX_PAPERS:
            break
        kind = p.get("kind") or ("preprint" if p.get("preprint") else "article")
        if kind not in kinds or kind == "other":
            continue
        if not usable(p):
            continue
        if since and p.get("year") and p["year"] < since:
            continue
        if until and p.get("year") and p["year"] > until:
            continue
        if any(_same_work(p, q) for q in picked):
            continue
        picked.append(dict(p, kind=kind))
    out = []
    for p in picked:
        sents = [x for x in _rs_split(text_of(p)) if not _RS_PLAN.search(x)]
        if sents:
            out.append(dict(p, n=len(out) + 1, sents=sents))
    return out


def _rs_json(raw):
    a, b = str(raw or "").find("{"), str(raw or "").rfind("}")
    obj = json.loads(raw[a:b + 1]) if 0 <= a < b else None
    if not isinstance(obj, dict):
        raise ValueError("its reply wasn't the JSON asked for")
    return obj


def _rs_papers_text(papers, labelled=False, cap=2600):
    out = []
    for p in papers:
        head = f"[{p['n']}] {p['title']} ({p.get('year') or 'no date'}; " + ("the person's own PDF" if p["kind"] == "mine" else RS_KIND_WORDS.get(p["kind"], ("paper",))[0]) + ")"
        if labelled:
            body = "\n".join(f"  P{p['n']}-S{i}: {s}" for i, s in enumerate(p["sents"]))
        else:
            body = "  " + " ".join(p["sents"])[:cap]
        out.append(head + "\n" + body)
    return "\n\n".join(out)


def _rs_idx(raw, pn, n_sents):
    """The sentence numbers the AI gave for paper pn, checked: whole numbers or labels like "P3-S2" for this paper, each
    a real sentence. None when anything given is not (a wrong label voids the claim; nothing is guessed)."""
    items = raw if isinstance(raw, list) else [raw]
    out = []
    for x in items[:2]:
        if isinstance(x, bool):
            return None
        if isinstance(x, int):
            k = x
        elif isinstance(x, str):
            m = re.fullmatch(r"\s*(?:P(\d+)\s*-\s*)?S?(\d+)\s*", x, re.I)
            if not m or (m.group(1) and int(m.group(1)) != pn):
                return None
            k = int(m.group(2))
        else:
            return None                    # 1.10 and the like: not a label this server gave
        if not 0 <= k < n_sents:
            return None
        if k not in out:
            out.append(k)
    return sorted(out) or None


def _rs_write(question, papers, deadline, context="", language=""):
    """The report's sentences, by part, each ending in the numbers of the papers that state it. 5.0: in the language the
    person chose (the quotations under each citation stay in the paper's own language)."""
    system = (
        "You are RamzAI, the research helper in Academic-Home, used by students and teachers. Answer the person's QUESTION "
        "from the numbered PAPERS only (their abstracts, or the first pages of the person's own PDF). Rules:\n"
        "- Say only what the papers state. No outside knowledge, no background of your own, no advice, no guesses.\n"
        "- Every sentence ends with the numbers of the papers that state what it says, in square brackets before the full "
        "stop, like [2] or [1, 4]. Cite a paper only for what its own text states. Give each figure exactly as the paper does. "
        "When you count the papers themselves, use words (\"two studies\").\n"
        "- Put it in your own words: never six or more words in a row from a paper. Don't name authors or years in the "
        "sentences (the brackets are the citations), and don't mention 'the abstract'.\n"
        "- Formal, plain English a first-year student understands, unless WRITTEN FOR says who the reader is: then write "
        "for that reader (a doctoral researcher or an organisation reads fuller, more exact language; a school pupil simpler). "
        "One idea per sentence. WRITTEN FOR never changes the rules above.\n"
        "- If the papers don't answer the question, say so plainly in the short answer, and say what they do show.\n"
        "Reply with JSON only: {\"short\": [2 or 3 sentences that answer the question], \"found\": [3 to 6 sentences on what "
        "the studies found], \"unsure\": [0 to 3 sentences on where they disagree, are limited or are unsure]}."
        + STYLE_PROSE
        + (f"\n- Write every sentence in {language}, in the formal register of academic writing in {language}, numbers in digits; "
           "keep the paper numbers in square brackets exactly as above." if language and language.lower() not in ("english", "en") else ""))
    user = "QUESTION: " + question + ("\nWRITTEN FOR: " + context if context else "") + "\n\nPAPERS:\n" + _rs_papers_text(papers)
    obj = None
    for attempt in (1, 2):
        # 4.9.1: one more try when the AI is busy, slow or answers with something unreadable (the evidence table reads
        # the same papers at the same moment, and a busy answer to one of the two stopped the whole research)
        try:
            raw = call_claude(system_prompt=system, user_message=user, max_tokens=1400,
                              timeout=max(10, min(90, deadline - time.time() - 60)), feature="research-write")
            obj = _rs_json(raw)
            break
        except Exception as e:
            logger.warning(f"Research: the writing didn't run (try {attempt}): {e.__class__.__name__}: {str(e)[:120]}")
            if attempt == 2 or deadline - time.time() < 95:
                raise
            time.sleep(5)
    parts = []
    for key, head in RS_SECTIONS:
        sents = [re.sub(r"\s+", " ", _finish_text(str(s), "words")).strip() for s in (obj.get(key) or []) if isinstance(s, str) and s.strip()]
        parts.append((head, [s for s in sents if s][:8]))
    return parts


_RS_MARK = re.compile(r"\s*\[(\d{1,2}(?:\s*[,;–\-]\s*\d{1,2})*)\]")


def _rs_unmark(sentence):
    """A written sentence without its [n] marks, and the papers it named."""
    named = set()
    for m in _RS_MARK.finditer(sentence):
        for x in re.split(r"\s*[,;]\s*", m.group(1)):
            if re.fullmatch(r"\d{1,2}\s*[–\-]\s*\d{1,2}", x):
                a, b = [int(y) for y in re.split(r"\s*[–\-]\s*", x)]
                named.update(range(a, min(b, a + 9) + 1))
            elif x.strip().isdigit():
                named.add(int(x))
    text = _RS_MARK.sub("", sentence).strip()
    text = re.sub(r"\s+([.!?,;:])", r"\1", text)
    if text and text[-1] not in ".!?":
        text += "."
    return text, named


def _rs_check(question, papers, sentences, deadline):
    """For each sentence: the papers, and their sentences by label, that state it ("full") or part of it ("part").
    Returns ({i: [(paper n, [sentence numbers], "full"|"part")]}, {i: True when something the AI gave for it was
    invalid}) and the set dropped, or None (no check)."""
    system = (
        "You check a research summary so that every sentence is backed by the papers' own words. Below: the PAPERS, each "
        "sentence of each paper labelled like P3-S2 (paper 3, sentence 2), then the person's QUESTION, then the numbered "
        "SENTENCES of the summary. For each summary sentence, give the paper sentences that state what it says:\n"
        "- \"full\" when that paper's sentences you give state everything the summary sentence says, every figure included;\n"
        "- \"part\" when they state only part of it.\n"
        "Keep a summary sentence only when the papers you give together state all of it (at least one \"full\", or \"part\" "
        "ones that together cover every part). Otherwise mark it drop. Any paper may back any sentence. At most 3 papers per "
        "sentence and at most 2 sentences per paper. Never pair a summary sentence with a paper sentence that doesn't state "
        "it. Reply with JSON only, like {\"sentences\": [{\"n\": 1, \"cite\": [{\"p\": 3, \"s\": [\"P3-S2\"], \"v\": \"full\"}]}, "
        "{\"n\": 2, \"drop\": true}]}.\n\nPAPERS:\n" + _rs_papers_text(papers, labelled=True))
    user = ("QUESTION: " + question + "\n\nSENTENCES:\n" + "\n".join(f"{i + 1}. {s}" for i, s in enumerate(sentences)))
    obj = None
    for attempt in (1, 2):
        try:
            raw = call_claude(system_prompt=system, user_message=user, max_tokens=min(2500, 90 * len(sentences) + 200),
                              timeout=max(8, min(60, deadline - time.time())), feature="research-check")
            obj = _rs_json(raw)
            break
        except Exception as e:
            logger.warning(f"Research: the sentence check didn't run ({str(e)[:100]})")
            if attempt == 2 or deadline - time.time() < 40:
                return None
    by_n = {p["n"]: p for p in papers}
    out, doubt, drop = {}, {}, set()
    for it in obj.get("sentences") or []:
        try:
            if not isinstance(it, dict):
                continue
            i = it.get("n")
            if isinstance(i, bool) or not isinstance(i, int):
                i = int(str(i)) if re.fullmatch(r"\d{1,3}", str(i or "")) else None
            if i is None or not 1 <= i <= len(sentences):
                continue
            i -= 1
            if it.get("drop"):
                drop.add(i)
                continue
            cites, bad = [], False
            given = it.get("cite") if isinstance(it.get("cite"), list) else []
            for c in given[:3]:
                pn = c.get("p") if isinstance(c, dict) else None
                if isinstance(pn, bool) or not isinstance(pn, int):
                    pn = int(pn) if isinstance(pn, str) and re.fullmatch(r"\s*\d{1,2}\s*", pn) else None
                p = by_n.get(pn)
                v = c.get("v") if isinstance(c, dict) else None
                nums = _rs_idx(c.get("s"), pn, len(p["sents"])) if p else None
                if not p or not nums or v not in ("full", "part") or any(pn == x[0] for x in cites):
                    bad = True
                    continue
                cites.append((pn, nums, v))
            if len(given) > 3:
                bad = True
            if cites:
                out[i] = cites
                doubt[i] = bad
            else:
                drop.add(i)
        except Exception:
            continue
    return out, doubt, drop


def _rs_keep(sentence, cites, doubtful, by_n, all_text):
    """Whether a checked sentence may stay: backed fully (one "full"; or "part" from at least two papers when nothing the
    AI gave was wrong), every figure in its quoted sentences, and in its own words (no run of ten words from any paper)."""
    if not cites:
        return False
    full = any(v == "full" for _, _, v in cites)
    if not full and (doubtful or len(cites) < 2):
        return False
    quoted = " ".join(by_n[pn]["sents"][k] for pn, ks, _ in cites for k in ks)
    if not _rs_figures_ok(sentence, quoted, n_cited=len({pn for pn, _, _ in cites}), n_read=len(by_n)):
        return False
    if _c45_copied(sentence, all_text, run=10):
        return False
    return True


def _rs_evidence(question, papers, deadline):
    """For each paper: the sentence(s) of its own that best state what it found about the question, and whether that
    answers the question yes, partly, no, or is unclear (only when the question can be answered yes or no)."""
    system = (
        "For each numbered PAPER below (each sentence labelled like P3-S2: paper 3, sentence 2), pick the one or two sentences "
        "that best state what that paper found about the person's QUESTION. If the QUESTION can be answered yes or no, also "
        "say how that paper's finding answers it: \"yes\", \"partly\", \"no\", or \"unclear\" (when it doesn't really answer "
        "it). Judge only by the paper's own sentences. Reply with JSON only, like {\"yesno\": true, \"papers\": [{\"p\": 1, "
        "\"s\": [\"P1-S2\"], \"a\": \"yes\"}]}; when the question isn't a yes-or-no question, give \"yesno\": false and leave "
        "\"a\" out.")
    user = "QUESTION: " + question + "\n\nPAPERS:\n" + _rs_papers_text(papers, labelled=True)
    obj = None
    for attempt in (1, 2):
        try:
            raw = call_claude(system_prompt=system, user_message=user, max_tokens=min(1500, 60 * len(papers) + 200),
                              timeout=max(8, min(70, deadline - time.time() - 40)), feature="research-evidence")
            obj = _rs_json(raw)
            break
        except Exception as e:
            logger.warning(f"Research: the evidence table's reading didn't run ({str(e)[:100]})")
            if attempt == 2 or deadline - time.time() < 70:
                return {}
    by_n = {p["n"]: p for p in papers}
    yesno = obj.get("yesno") is True
    out = {}
    for it in obj.get("papers") or []:
        try:
            pn = it.get("p") if isinstance(it, dict) else None
            if isinstance(pn, bool) or not isinstance(pn, int):
                pn = int(pn) if isinstance(pn, str) and re.fullmatch(r"\s*\d{1,2}\s*", pn) else None
            p = by_n.get(pn)
            if not p or pn in out:
                continue
            nums = _rs_idx(it.get("s"), pn, len(p["sents"])) or []
            a = it.get("a") if yesno and nums and it.get("a") in ("yes", "partly", "no", "unclear") else None
            out[pn] = (nums, a)
        except Exception:
            continue
    return out


def _rs_result(question, years, kinds, papers, parts, evidence, stats, style="apa"):
    """What the app shows: the report (sentences with their papers, quotes by number and verdicts), the papers (with their
    references, citations, quotes, design and who they studied) and what it rests on. The citations and the reference
    list are worked out from the cited papers alone (APA), so a paper that isn't cited never changes them."""
    cited = set()
    sections = []
    for head, sents in parts:
        keep = [s for s in sents if s.get("c")]
        for s in keep:
            cited.update(c["p"] for c in s["c"])
        if keep:
            sections.append({"h": head, "s": keep})
    strip = lambda p: {k: v for k, v in p.items() if k not in ("sents", "abstract_full")}
    on_list = [p for p in papers if p["n"] in cited]
    keys = {p["n"]: e for p, e in zip(on_list, enrich_sources([strip(p) for p in on_list], style))}
    # a paper read but not cited: its reference and label are worked out alone (no a/b from the other papers), with the
    # start of its title added when a cited paper, or another paper read, has the same label
    alone = {p["n"]: enrich_sources([strip(p)], style)[0] for p in papers if p["n"] not in keys}
    taken = {_rs_cite_short(e.get("citation_key")) for e in keys.values()}
    clash = Counter(_rs_cite_short(e.get("citation_key")) for e in alone.values())
    out_papers = []
    for p in papers:
        e = keys.get(p["n"]) or alone[p["n"]]
        cite = _rs_cite_short(e.get("citation_key"))
        if p["n"] in alone and (cite in taken or clash[cite] > 1):
            words = (p.get("title") or "").split()
            cite = f"{cite}, “{' '.join(words[:5])}{'…' if len(words) > 5 else ''}”"
        url = p.get("url") or (f"https://doi.org/{p['doi']}" if p.get("doi") else None)
        ref = e.get("ref_text") or p["title"]
        if url:
            ref = ref.rstrip() + " " + url
        finding, answer = evidence.get(p["n"], ([], None))
        out_papers.append({
            "n": p["n"], "title": p["title"], "authors": p.get("authors") or [], "year": p.get("year"), "venue": p.get("venue"),
            "kind": p["kind"], "doi": p.get("doi"), "url": url, "pdf": p.get("pdf"), "biblio": p.get("biblio") or {},
            "cite": cite, "ref": ref, "sents": p["sents"],
            "design": _rs_design(p["sents"], p["kind"]), "people": _rs_people(p["sents"]),
            "finding": finding, "answer": answer if finding else None, "cited": p["n"] in cited, "lib": p.get("lib")})
    yrs = [p["year"] for p in papers if p.get("year")]
    counts = Counter(p["kind"] for p in papers)
    kinds_words = ", ".join(f"{counts[k]} {RS_KIND_WORDS[k][0 if counts[k] == 1 else 1]}" for k in ("article", "review", "book", "conference", "thesis", "report", "preprint", "mine") if counts.get(k))
    limits = (f"{len(papers)} paper{'' if len(papers) == 1 else 's'}" + (f" from {min(yrs)} to {max(yrs)}" if yrs and min(yrs) != max(yrs) else f" from {yrs[0]}" if yrs else "")
              + (f": {kinds_words}" if kinds_words else "") + ". RamzAI read their abstracts"
              + (", and the first pages of your PDFs," if counts.get("mine") else "") + " not the whole papers, so read a paper before you rely on it.")
    if counts.get("preprint"):
        limits += " Preprints are not yet peer reviewed."
    return {"v": 1, "question": question, "asked": {"years": years, "kinds": kinds}, "at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            "style": style, "papers": out_papers, "sections": sections, "limits": limits,
            "stats": dict(stats, years=[min(yrs), max(yrs)] if yrs else None, kinds=dict(counts))}


@app.route("/api/once/<key>", methods=["GET", "OPTIONS"])
@limiter.limit("900 per hour", key_func=_account_or_ip_key)
@auth_required
def once_state(key):
    """4.9: the answer to a request whose answer never arrived (a dropped connection): the kept answer, or 202 while
    it's still being worked on, or 404 if it never arrived or didn't go through (the app then says so)."""
    if not _ONCE_KEY.match(key or ""):
        return jsonify({"error": "Not found", "unknown": True}), 404
    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT route, state, status, body, mimetype, created_at FROM request_once WHERE user_id = %s AND key = %s",
                        (request.current_user["id"], key))
            row = cur.fetchone()
    finally:
        conn.close()
    if row and row["state"] == "done":
        return _once_replay(row)
    if row and row["state"] == "working" and _once_age(row) < _ONCE_STUCK_S:
        return jsonify({"working": True, "retry_in": 3}), 202
    return jsonify({"error": "Not found", "unknown": True}), 404


# ---------- 5.0 the brief: what RamzAI finds out before research begins (the owner, 10 Nov 2026) ----------
# "The app is set to support users, not to dictate for them": someone unsure of a topic is helped to choose one, an
# unclear answer is asked about again, and the person's own template (any file with the guide in readable words) is
# read into its headings and instructions. "Make everything broad … everyone stands a chance to make research": the
# questions never assume a student: a pupil, a student at any level up to a doctorate, a teacher, a researcher, a
# professional, an organisation, an institution or someone looking into something for themselves. Every question
# comes with a wide range of answers to tap, and the person can always answer in their own words instead.
# None of these takes points; none of them keeps the file or the words.
BRIEF_LIMIT = "90 per hour;400 per day"
BRIEF_KINDS = {"question": "a research question to look up in published papers",
               "auto": "a piece of writing (an essay, a report, a thesis chapter, a policy brief, an article…) drafted from published papers",
               "manual": "a piece of writing the person will write themselves, citing published papers"}
BRIEF_ANYONE = ("The person may be anyone: a school pupil, a student at any level up to a doctorate, a teacher, a lecturer, a "
                "researcher, a professional, an organisation, an institution, or someone looking into a subject for themselves. "
                "Never assume they are a student unless ABOUT THE PERSON says so.")


def _brief_json(raw):
    a, b = str(raw or "").find("{"), str(raw or "").rfind("}")
    obj = json.loads(raw[a:b + 1]) if 0 <= a < b else None
    if not isinstance(obj, dict):
        raise ValueError("not the JSON asked for")
    return obj


def _brief_line(v, cap):
    return re.sub(r"\s+", " ", str(v or "")).strip()[:cap]


def _brief_list(v, n, cap):
    """Up to n short lines from a list the app sent."""
    return [x for x in (_brief_line(i, cap) for i in (v if isinstance(v, list) else [])[:n * 2]) if x][:n]


def _brief_options(v, avoid=()):
    """The answers to tap: short, different from each other, none the person has already been shown."""
    seen, out = {re.sub(r"\W+", "", a.lower()) for a in avoid}, []
    for o in (v if isinstance(v, list) else [])[:10]:
        if not isinstance(o, str):
            continue
        t = _finish_text(_brief_line(o, 80), "words").strip(" .")
        k = re.sub(r"\W+", "", t.lower())
        if not t or not k or k in seen or k in ("other", "somethingelse", "noneoftheabove", "noneofthese"):
            continue           # "something else" is always there in the app: the person's own words
        seen.add(k)
        out.append(t)
    return out[:6]


@app.route("/api/brief/topics", methods=["POST", "OPTIONS"])
@limiter.limit(BRIEF_LIMIT, key_func=_account_or_ip_key)
@auth_required
def brief_topics():
    """Someone who isn't sure what to research: five topics to choose from, from what they study, do or care about."""
    data = request.get_json(silent=True) or {}
    interest = _brief_line(data.get("interest"), 400)
    kind = data.get("kind") if data.get("kind") in BRIEF_KINDS else "auto"
    about = _brief_line(data.get("about") or data.get("level"), 400)
    avoid = _brief_list(data.get("avoid"), 15, 200)
    if len(interest) < 3:
        return jsonify({"error": "Say what you study, work on or are curious about, in a few words."}), 400
    system = (
        "You help a person who isn't sure what to research choose a topic. " + BRIEF_ANYONE + " From what they say they "
        "study, work on or care about, suggest five different, specific, researchable topics for " + BRIEF_KINDS[kind] + ". "
        "Each must be narrow enough to finish, likely to have free published papers, fitted to who they are and what the "
        "work is for (ABOUT THE PERSON, when given: a doctoral thesis needs a sharper, more original angle than a school "
        "assignment; an organisation needs a question it can act on), and worded plainly (at most 18 words)"
        + (", as a question that papers can answer" if kind == "question" else "") + ". Range widely: different angles, "
        "groups, places, methods or sides of the subject, not five wordings of one idea. Never invent facts. Give each one "
        "short reason (at most 14 words) why it is a good choice. Never repeat a topic in ALREADY SHOWN. Reply with JSON "
        "only: {\"topics\": [{\"t\": \"the topic\", \"why\": \"the reason\"}]}.")
    user = ("WHAT THEY STUDY, WORK ON OR CARE ABOUT: " + interest + ("\nABOUT THE PERSON: " + about if about else "")
            + ("\nALREADY SHOWN:\n- " + "\n- ".join(avoid) if avoid else ""))
    try:
        obj = _brief_json(call_claude(system_prompt=system, user_message=user, max_tokens=700, timeout=45, feature="brief-topics"))
    except Exception as e:
        logger.warning(f"Brief: topic ideas didn't run: {str(e)[:120]}")
        return jsonify({"error": "RamzAI couldn’t suggest topics just now. Please try again, or type a topic of your own."}), 502
    out, seen = [], {re.sub(r"\W+", "", a.lower()) for a in avoid}
    for x in (obj.get("topics") or [])[:8]:
        t = _brief_line(x.get("t"), 200) if isinstance(x, dict) else ""
        k = re.sub(r"\W+", "", t.lower())
        if t and k not in seen:
            seen.add(k)
            out.append({"t": _finish_text(t, "words"), "why": _finish_text(_brief_line(x.get("why"), 160), "words")})
    if len(out) < 2:
        return jsonify({"error": "RamzAI couldn’t suggest topics just now. Please try again, or type a topic of your own."}), 502
    return jsonify({"topics": out[:5]})


_BRIEF_VAGUE = re.compile(r"^(?:idk|i don'?t know|not sure|anything|something|any|whatever|help|n/?a|none|nothing|\?+|\.+|ok|okay|yes|no)$", re.I)


@app.route("/api/brief/check", methods=["POST", "OPTIONS"])
@limiter.limit(BRIEF_LIMIT, key_func=_account_or_ip_key)
@auth_required
def brief_check():
    """Is what the person typed clear enough to start from? If not: one short question to ask them, with a wide range
    of answers to tap (they can always answer in their own words instead). When the check itself can't run, the answer
    is taken as it is (nobody is held up). other=true: the person asked for other answers to the same question."""
    data = request.get_json(silent=True) or {}
    text = _brief_line(data.get("text"), 600)
    kind = data.get("kind") if data.get("kind") in BRIEF_KINDS else "auto"
    told = _brief_list(data.get("told"), 4, 300)             # what they answered before
    about = _brief_line(data.get("about"), 400)
    avoid = _brief_list(data.get("avoid"), 18, 80)           # answers they have already been shown
    asked = _brief_line(data.get("ask"), 220)                # other=true: the question those answers were for
    other = bool(data.get("other")) and bool(asked)
    if not text:
        return jsonify({"clear": False, "ask": "What would you like this to be about?", "options": []})
    if not other and (_BRIEF_VAGUE.match(text) or len(text) < 3):
        return jsonify({"clear": False, "unsure": True,
                        "ask": "Could you say a little more? A few words on the subject, the people or the place are enough.", "options": []})
    if not other and len(told) >= 2:                          # asked twice already: their words are taken as they are
        return jsonify({"clear": True})
    range_rule = ("Give 4 to 6 short answers they could tap (each at most 9 words), as different from each other as you "
                  "can, so that together they cover the wide range of what a person could mean: different groups of people, "
                  "places, fields, levels or meanings, never several versions of one. Don't include 'other' or 'something "
                  "else': the app always lets them answer in their own words.")
    if other:
        system = ("A person typed the TOPIC below to start " + BRIEF_KINDS[kind] + ", and was asked the QUESTION to make it "
                  "clear. None of the answers shown fitted, and they asked to see others. " + BRIEF_ANYONE + " " + range_rule +
                  " Never repeat an answer in ALREADY SHOWN. Reply with JSON only: {\"options\": [\"...\"]}.")
        user = ("TOPIC: " + text + "\nQUESTION: " + asked + ("\nABOUT THE PERSON: " + about if about else "")
                + ("\nALREADY SHOWN:\n- " + "\n- ".join(avoid) if avoid else ""))
    else:
        system = (
            "A person typed the TOPIC below to start " + BRIEF_KINDS[kind] + ". " + BRIEF_ANYONE + " Decide whether it is "
            "clear enough to begin: a reader can tell what subject it is about. Be generous: a broad topic, a title in "
            "capitals, informal words, spelling mistakes or a topic that could be narrower are all CLEAR. It is NOT clear only "
            "when the subject itself can't be told (a word with several unrelated meanings, an unfinished thought, a reference "
            "to something you can't see like 'my assignment' or 'chapter 3', or words that don't say a subject at all), or when "
            "it is one or two words that name only a very wide subject (stress, education, climate change, marketing): that is "
            "too wide to research well, so ask which part of it, or which people or place. When it is not clear, write ONE "
            "short, kind question (at most 22 words, plain English) that would settle it. "
            + range_rule + " Never refuse, never lecture. Reply with JSON only: {\"clear\": true} or "
            "{\"clear\": false, \"ask\": \"...\", \"options\": [\"...\"]}.")
        user = ("TOPIC: " + text + ("\nABOUT THE PERSON: " + about if about else "")
                + ("\nWHAT THEY SAID BEFORE: " + " | ".join(told) if told else ""))
    try:
        obj = _brief_json(call_claude(system_prompt=system, user_message=user, max_tokens=320, timeout=12, feature="brief-check"))
    except Exception as e:
        logger.warning(f"Brief: the clarity check didn't run: {str(e)[:120]}")
        if other:
            return jsonify({"error": "RamzAI couldn’t think of other answers just now. You can say it in your own words."}), 502
        return jsonify({"clear": True, "unchecked": True})
    if other:
        opts = _brief_options(obj.get("options"), avoid)
        if not opts:
            return jsonify({"error": "RamzAI has no other answers to suggest. You can say it in your own words."}), 404
        return jsonify({"clear": False, "ask": asked, "options": opts})
    if obj.get("clear") is not False:
        return jsonify({"clear": True})
    ask = _finish_text(_brief_line(obj.get("ask"), 220), "words")
    if not ask:
        return jsonify({"clear": True})
    return jsonify({"clear": False, "ask": ask, "options": _brief_options(obj.get("options"))})


BRIEF_REVIEW_MAX = 3            # questions at one time
BRIEF_REVIEW_ROUNDS = 6         # answers already given in this brief: after that, RamzAI starts with what it has


@app.route("/api/brief/review", methods=["POST", "OPTIONS"])
@limiter.limit(BRIEF_LIMIT, key_func=_account_or_ip_key)
@auth_required
def brief_review():
    """Before the person confirms: RamzAI reads the whole brief the way a careful supervisor or editor would, and asks
    only for what is still vital (missing, unclear or at odds), each question with a wide range of answers to tap. Most
    short pieces need nothing more. When this can't run, the brief goes ahead as it is."""
    data = request.get_json(silent=True) or {}
    kind = data.get("kind") if data.get("kind") in BRIEF_KINDS else "auto"
    b = data.get("brief") if isinstance(data.get("brief"), dict) else {}
    topic = _brief_line(b.get("topic"), 400)
    if not topic:
        return jsonify({"questions": []})
    lines = [("TOPIC", topic), ("WHO IT IS FOR", _brief_line(b.get("for"), 200)), ("WHAT IT IS FOR", _brief_line(b.get("purpose"), 160)),
             ("FOCUS (place or setting)", _brief_line(b.get("place"), 120)), ("WHAT IT IS", _brief_line(b.get("piece"), 120)),
             ("LENGTH", _brief_line(b.get("length"), 60)), ("CITATIONS", _brief_line(b.get("style"), 60)),
             ("PAPERS FROM", _brief_line(b.get("years"), 80)), ("KINDS OF SOURCES", _brief_line(b.get("kinds"), 160)),
             ("READERS", _brief_line(b.get("readers"), 160)), ("IT MUST FOLLOW", _brief_line(b.get("more"), 800)),
             ("TEMPLATE", _brief_line(b.get("template"), 160)), ("LANGUAGE", _brief_line(b.get("language"), 60)),
             ("THE WORD FILE", _brief_line(b.get("look"), 160))]
    heads = _brief_list(b.get("headings"), 30, 90)
    answered = []
    for x in (b.get("asked") if isinstance(b.get("asked"), list) else [])[:12]:
        if isinstance(x, dict) and _brief_line(x.get("q"), 220):
            answered.append((_brief_line(x.get("q"), 220), _brief_line(x.get("a"), 300) or "(skipped: don't ask again)"))
    if len(answered) >= BRIEF_REVIEW_ROUNDS:
        return jsonify({"questions": []})
    most = 2 if kind == "question" else BRIEF_REVIEW_MAX
    system = (
        "You are RamzAI, the research helper in Academic-Home. A person is about to start " + BRIEF_KINDS[kind] + ". "
        + BRIEF_ANYONE + " Read their BRIEF the way a careful supervisor, editor or client lead would before work begins. "
        "Decide whether anything VITAL is still missing, unclear or at odds: something that would change what is written, "
        "and that a careful human would ask before starting. Ask at most " + str(most) + " questions, the most important "
        "first. Ask NONE when the brief is good enough to start well: most short pieces need none, and it is better to "
        "start than to quiz people. The higher the stakes (a thesis or dissertation, a proposal, a journal article, a policy "
        "brief, a report for an organisation or institution), the more it is worth asking.\n"
        "What is worth asking, when the brief doesn't already say and it matters for this kind of work: which chapter or "
        "part of a thesis or report this is; the research question, aim or objectives (a proposal, a thesis, an article); "
        "the planned approach (reviewing what is published, numbers, interviews, both) when the piece must describe a "
        "method; the theory, model or framework to use (postgraduate work); the group of people, the place or the period "
        "when the topic is wide and the level calls for a focus; the argument or position (an essay that must argue); the "
        "decision, problem or audience behind a report or briefing; anything in the brief that contradicts itself (a "
        "500-word thesis chapter; a policy brief for children).\n"
        "Never ask what the BRIEF or ALREADY ANSWERED covers. Never ask the same thing in other words. Never ask for "
        "personal details, for their own data or results (the work is written from published papers), or for anything you "
        "couldn't use. Never ask about length, citation style, years, kinds of sources, language, the Word file or "
        "template when they are given.\n"
        "Each question: plain, kind English, at most 24 words, one thing only. With each give 4 to 6 short answers to tap "
        "(each at most 9 words), as different from each other as you can so that together they cover the wide range of "
        "what a person could want; include 'I’m not sure, you choose' as the last answer when choosing for them would "
        "be sensible. Don't include 'other' or 'something else': the app always lets them answer in their own words. Give "
        "each a short reason (at most 14 words) why it helps, and a short id (one or two words, lower case). Reply with "
        "JSON only: {\"questions\": [{\"id\": \"...\", \"ask\": \"...\", \"why\": \"...\", \"options\": [\"...\"]}]} "
        "(an empty list when nothing is needed).")
    user = ("BRIEF\n" + "\n".join(f"{k}: {v}" for k, v in lines if v)
            + ("\nTEMPLATE HEADINGS:\n- " + "\n- ".join(heads) if heads else "")
            + ("\n\nALREADY ANSWERED\n" + "\n".join(f"Q: {q}\nA: {a}" for q, a in answered) if answered else ""))
    try:
        obj = _brief_json(call_claude(system_prompt=system, user_message=user, max_tokens=700, timeout=18, feature="brief-review"))
    except Exception as e:
        logger.warning(f"Brief: the review didn't run: {str(e)[:120]}")
        return jsonify({"questions": [], "unchecked": True})
    out, seen = [], {re.sub(r"\W+", "", q.lower()) for q, _ in answered}
    for x in (obj.get("questions") if isinstance(obj.get("questions"), list) else [])[:most + 2]:
        if not isinstance(x, dict):
            continue
        ask = _finish_text(_brief_line(x.get("ask"), 240), "words")
        k = re.sub(r"\W+", "", ask.lower())
        if not ask or k in seen:
            continue
        seen.add(k)
        qid = re.sub(r"[^a-z0-9]+", "-", _brief_line(x.get("id"), 30).lower()).strip("-") or f"q{len(out) + 1}"
        opts = []
        for o in (x.get("options") if isinstance(x.get("options"), list) else [])[:8]:
            t = _finish_text(_brief_line(o, 80), "words").strip(" .") if isinstance(o, str) else ""
            if t and re.sub(r"\W+", "", t.lower()) not in ("other", "somethingelse") and t not in opts:
                opts.append(t)
        out.append({"id": qid, "ask": ask, "why": _finish_text(_brief_line(x.get("why"), 140), "words"), "options": opts[:6]})
        if len(out) >= most or len(answered) + len(out) >= BRIEF_REVIEW_ROUNDS:
            break
    return jsonify({"questions": out})


_BRIEF_HEAD = re.compile(r"^(?:(?:chapter|section|part)\s+\w+[:.\s-]*|\d+(?:\.\d+){0,3}[.)]?\s+|[IVXLC]{1,6}[.)]\s+|[A-Z][.)]\s+)?(.{2,90})$", re.I)


def _brief_headings(text):
    """The lines of a template that read as headings: numbered ("2.1 Methods", "Chapter One"), in capitals, or a short
    line without a full stop standing on its own."""
    out, seen = [], set()
    lines = [re.sub(r"\s+", " ", l).strip() for l in str(text or "").splitlines()]
    for i, l in enumerate(lines):
        if not l or len(l) > 90 or len(l.split()) > 12:
            continue
        numbered = bool(re.match(r"^(?:(?:chapter|section|part)\s+\w+|\d+(?:\.\d+){0,3}[.)]?\s+\S|[IVXLC]{1,6}[.)]\s+\S)", l, re.I))
        caps = l.upper() == l and sum(c.isalpha() for c in l) >= 4
        alone = (l[-1] not in ".,;:?!" and len(l.split()) <= 7 and l[0].isupper()
                 and (i == 0 or not lines[i - 1]) and (i + 1 >= len(lines) or not lines[i + 1] or len(lines[i + 1]) > 60))
        if not (numbered or caps or alone):
            continue
        key = l.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(l)
        if len(out) >= 40:
            break
    return out


@app.route("/api/brief/template", methods=["POST", "OPTIONS"])
@limiter.limit(HEAVY_LIMIT, key_func=_account_or_ip_key)
@auth_required
def brief_template():
    """The person's own template or guide (a Word file, a PDF, slides, plain text, a photo of a page…): its headings and
    its instructions, in words. The file is read and deleted at once; nothing of it is kept here."""
    if "file" not in request.files or not request.files["file"].filename:
        return jsonify({"error": "Choose the file that holds your template or guide."}), 400
    file = request.files["file"]
    ext = file.filename.rsplit(".", 1)[-1].lower() if "." in file.filename else ""
    work_dir = tempfile.mkdtemp(prefix=f"mc_tpl_{uuid.uuid4().hex[:8]}_")
    try:
        src_path = os.path.join(work_dir, safe_upload_filename(file.filename)[0])
        file.save(src_path)
        text = extract_text(src_path, ext)
    except ConversionError as e:
        return jsonify({"error": str(e)}), 422
    except Exception as e:
        logger.error(f"Brief: a template couldn't be read: {e}", exc_info=True)
        return jsonify({"error": "Academic-Home couldn’t read that file. Try a Word file, a PDF or plain text."}), 422
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)
    text = re.sub(r"[ \t]+", " ", str(text or "")).strip()
    text = re.sub(r"\n{3,}", "\n\n", text)
    if len(re.sub(r"\W", "", text)) < 30:
        return jsonify({"error": "That file doesn’t hold a guide in words Academic-Home can read. Try another file, or choose "
                                 "one of the templates here."}), 422
    heads = _brief_headings(text)
    return jsonify({"name": safe_download_name(file.filename), "headings": heads, "guide": text[:3500],
                    "cut": len(text) > 3500, "words": len(text.split())})


_RS_DOWN = re.compile(r"HTTP (?:4(?:08|29)|5\d\d)|failed|timed? ?out|busy|Timeout|ConnectionError|rate", re.I)


@app.route("/api/research", methods=["POST", "OPTIONS"])
@limiter.limit(AI_LIMIT, key_func=_account_or_ip_key)
@auth_required
def research_endpoint():
    started = time.time()
    data = request.get_json(silent=True) or {}
    question = re.sub(r"\s+", " ", str(data.get("question") or "")).strip()
    # 5.0: who the report is for (a doctoral student in nursing, an organisation…), as the person said before starting
    context = re.sub(r"\s+", " ", str(data.get("context") or "")).strip()[:700]
    language = re.sub(r"[^\w\s\-’'()]", "", str(data.get("language") or "")).strip()[:40]      # 5.0: the language of the report
    if len(question) < 12 or len(question.split()) < 3:
        return jsonify({"error": "Say a little more: who, where and what you want to know."}), 400
    if len(question) > 400:
        return jsonify({"error": "Make the question shorter (400 characters at most)."}), 400
    try:
        years = int(data.get("years") or 0)
    except (TypeError, ValueError):
        years = 0
    years = years if 1 <= years <= 100 else 0            # 5.0: any number of years (4.9: 5, 10 or 20 only)
    # 5.0: or the years the person chose themselves ("from 2010 to 2018")
    this_year = datetime.now(timezone.utc).year
    year_from, year_to = _brief_year(data.get("year_from"), this_year), _brief_year(data.get("year_to"), this_year)
    if year_from and year_to and year_from > year_to:
        year_from, year_to = year_to, year_from
    until = year_to if (year_to and year_to < this_year) else None
    kinds = [k for k in RS_KINDS if k in (data.get("kinds") or [])] or ["article", "review"]
    library = _rs_library(data.get("library"))
    user = request.current_user
    gate = _scholar_gate(user, "Research", research=True)
    if gate:
        return gate
    try:
        charged = _log_sage_message_and_consume(user, POINTS_ASK_RESEARCH, research=True)
    except ConversionError as e:
        return _points_reply(e)
    state = {"settled": False}

    def settle(ok):
        if state["settled"]:
            return
        state["settled"] = True
        if not ok:
            _refund_sage_message(user, charged, POINTS_ASK_RESEARCH, research=True)

    def line(obj):
        return json.dumps(obj, ensure_ascii=False) + "\n"

    def run():
        deadline = started + 200
        pool = _futures.ThreadPoolExecutor(max_workers=3)

        def waiting(futs):
            """Yields a heartbeat every 4 seconds until one of futs is done (so a person who left is noticed soon)."""
            while True:
                done, _ = _futures.wait(futs, timeout=4, return_when=_futures.FIRST_COMPLETED)
                if done:
                    return
                if time.time() > deadline + 15:
                    raise ConversionError("timed out")
                yield line({"beat": 1})

        try:
            yield line({"step": "search", "state": "now"})
            stats = {}
            since = year_from or ((this_year - years + 1) if years else None)
            terms = _paper_terms(question)
            query = " ".join(terms[:8]) if len(terms) >= 2 else _search_topic(question)
            fs = pool.submit(search_papers, query, 24, since, True, kinds, stats, **({"until": until} if until else {}))
            yield from waiting({fs})
            found, diag = fs.result()
            yield line({"step": "search", "state": "done", "found": stats.get("found", len(found)), "from": stats.get("from", [])})
            yield line({"step": "keep", "state": "now"})
            papers = _rs_pick(found, library, question, kinds, since, until)
            if len(papers) < RS_MIN_PAPERS:
                if not found and diag and _RS_DOWN.search(diag):
                    msg = "The paper search services aren’t answering right now. Please try again in a few minutes."
                else:
                    msg = ((f"RamzAI found only {len(papers)} free paper{'' if len(papers) == 1 else 's'} that fit" if papers
                            else "RamzAI found no free papers that fit") + ". Try other words, more years or more kinds of sources.")
                yield line({"step": "keep", "error": msg + " This didn’t use any points."})
                return
            yrs = [p["year"] for p in papers if p.get("year")]
            yield line({"step": "keep", "state": "done", "kept": len(papers), "years": [min(yrs), max(yrs)] if yrs else None,
                        "kinds": dict(Counter(p["kind"] for p in papers))})
            yield line({"step": "write", "state": "now"})
            yield line({"step": "evidence", "state": "now"})
            by_n = {p["n"]: p for p in papers}
            all_text = " ".join(" ".join(p["sents"]) for p in papers)
            fw = pool.submit(_rs_write, question, papers, deadline, context, language)
            fe = pool.submit(_rs_evidence, question, papers, deadline)
            pending, evidence, flat, fc = {fw, fe}, None, [], None
            while pending:
                yield from waiting(pending)
                for f in [x for x in pending if x.done()]:
                    pending.discard(f)
                    if f is fe:
                        evidence = f.result()
                        yield line({"step": "evidence", "state": "done", "rows": len(papers)})
                    elif f is fw:
                        try:
                            parts = f.result()      # may raise: the research stops, and the point is given back
                        except Exception as e:
                            logger.warning(f"Research failed at the writing: {e.__class__.__name__}: {str(e)[:160]}")
                            settle(False)
                            # 4.9.1: say what happened when the AI itself said it (busy, slow, not available)
                            why = str(e) if isinstance(e, ConversionError) and str(e) in (AI_MSG_BUSY, AI_MSG_SLOW, AI_MSG_UNAVAILABLE) \
                                else "RamzAI couldn’t finish this research. Please try again."
                            yield line({"step": "write", "error": why + " This didn’t use any points."})
                            return
                        yield line({"step": "write", "state": "done"})
                        yield line({"step": "check", "state": "now"})
                        for head, sents in parts:
                            for s in sents:
                                text, _named = _rs_unmark(s)
                                flat.append({"h": head, "t": text})
                        fc = pool.submit(_rs_check, question, papers, [x["t"] for x in flat], deadline)
                        pending.add(fc)
                    elif f is fc:
                        chk = f.result()
                        if chk is None:
                            raise ConversionError("the check didn't run")
                        cites, doubt, drop = chk
                        for i, x in enumerate(flat):
                            c = cites.get(i)
                            if i in drop or not c or not _rs_keep(x["t"], c, doubt.get(i, False), by_n, all_text):
                                x["c"] = None
                                continue
                            x["c"] = [{"p": pn, "q": ks, "v": v} for pn, ks, v in c]
                        kept = sum(1 for x in flat if x.get("c"))
                        yield line({"step": "check", "state": "done", "checked": len(flat), "removed": len(flat) - kept})
            kept = [x for x in flat if x.get("c")]
            if len(kept) < 2:
                yield line({"step": "check", "error": "The papers RamzAI found don’t back a report on this question. Try other words, "
                                                      "more years or more kinds of sources. This didn’t use any points."})
                return
            parts_out = [(head, [{"t": x["t"], "c": x["c"]} for x in flat if x["h"] == head and x.get("c")]) for _key, head in RS_SECTIONS]
            stats.update(checked=len(flat), removed=len(flat) - len(kept))
            result = _rs_result(question, years, kinds, papers, parts_out, evidence or {}, stats)
            if year_from or until:           # 5.0: the years the person chose themselves
                result["asked"].update({k: v for k, v in (("year_from", year_from), ("year_to", until)) if v})
            if language and language.lower() not in ("english", "en"):      # 5.0: the language it is written in
                result["asked"]["language"] = language
                heads = RS_SECTIONS_LANG.get(language.lower())
                if heads:
                    names = {en: tr for (_k, en), tr in zip(RS_SECTIONS, heads)}
                    for sec in result.get("sections") or []:
                        sec["h"] = names.get(sec["h"], sec["h"])
            yield line({"done": True, "result": result})
            settle(True)
            logger.info(f"Research: {len(papers)} papers, {len(flat)} sentences, {len(kept)} kept, {time.time() - started:.0f}s")
        except GeneratorExit:
            settle(False)
            raise
        except Exception as e:
            logger.warning(f"Research failed: {e.__class__.__name__}: {str(e)[:160]}")
            settle(False)
            yield line({"error": "RamzAI couldn’t finish this research. Please try again. This didn’t use any points."})
        finally:
            settle(False)
            pool.shutdown(wait=False, cancel_futures=True)      # a person who left doesn't hold a thread for the AI's answer

    resp = Response(stream_with_context(run()), mimetype="application/x-ndjson",
                    headers={"Cache-Control": "no-store, no-transform", "X-Accel-Buffering": "no"})
    resp.call_on_close(lambda: settle(False))
    return resp


# ---------- 4.1 RamzAI plans before it works ----------
REMY_PLAN_TOOLS = """
chat: answer a question, explain something, talk, or write a short text (an email, a message, a study plan). No params.
quiz: a practice quiz from a topic, pasted notes or an attached file. params: topic (what it is on, or "the attached file"),
  count (5, 8, 12, 15, 25 or 50; default 8), style (mixed, mcq, short, tf or case; default mixed), level (mixed, easy,
  medium, hard; default mixed), audience (student or tutor; default student).
deck: a whole slide deck drafted from a topic, from real research papers; no file needed. params: topic, group (student or tutor).
summarize: notes or an attached file turned into slides, a one-page revision sheet, a mind map, likely exam questions or
  key terms. Needs text or an attached file. params: kind (slides, sheet, map, questions or terms), audience (student or tutor).
write: automatic research writing: RamzAI drafts an essay, report or section from real papers, with citations and
  references. params: topic, style (apa, harvard, mla or chicago; default apa), length_words (300 to 3000; default 800).
cite: manual research writing in the Word editor: the person writes; real papers wait beside the page, one statement at a
  time, each with its citation and reference. params: topic.
convert: convert a file to another format. params: from_kind (docs, slides, pics, sheets, sound or video), to (the format
  wanted, like pdf, docx, pptx, jpg, mp3, or empty).
meeting: start an online meeting or class, with an agenda RamzAI can write. params: title, minutes (30 to 300).
focus: start the focus timer with calm sounds. params: minutes (25, 45 or 60).
pdf: edit, sign, fill in, protect or shrink a PDF. No params.
word: open or start a Word document to edit by hand. No params.
scan: scan paper pages with the camera into a clean PDF. No params.
studio: open Slide Studio to make or edit slides by hand. No params.
"""
REMY_PLAN_POINTS = {"quiz": POINTS_REMY, "deck": POINTS_DECK, "summarize": POINTS_REMY, "write": POINTS_WRITE, "cite": POINTS_RESEARCH, "chat": POINTS_REMY}
REMY_PLAN_SCHOLAR = {"write": "Automatic research writing", "cite": "Manual research writing"}
REMY_PLAN_SYSTEM = (
    "You are RamzAI's planner inside Academic-Home, a study app. You read what a person asks for and decide how RamzAI should "
    "respond, in one JSON object and nothing else. You never do the task yourself.\n\nWhat the app can do:" + REMY_PLAN_TOOLS +
    "\nAnswer with exactly one of these shapes:\n"
    '1. {"do":"chat"} when the person asks a question, wants something explained, is talking, or wants a short text written '
    "(an email, a message, a plan). Most messages are this.\n"
    '2. {"do":"ask","say":"<one short question>","options":["<choice>","<choice>"],"free":true} when the task is clear but one '
    "thing RamzAI must know is missing (for a quiz, deck, essay or citing: the topic; for summarize: whether there are notes or a "
    "file). Ask ONE question at a time, in plain words; give 2 to 4 short tappable choices when choices make sense, else an "
    "empty list. free is true when typing an answer makes sense.\n"
    '3. {"do":"plan","tool":"<tool>","title":"<one line: what RamzAI will make>","detail":["<a choice made, in plain words>", ...],'
    '"params":{...},"next":["<something useful to do afterwards>", ...]} when enough is known. Fill every param, using the '
    "defaults when the person did not say; put the choices that matter in detail (2 to 4 lines, so the person can change them). "
    "next has 2 or 3 short things. For tools with no params use an empty object.\n"
    '4. {"do":"suggest","say":"<one friendly line>","options":["<a concrete thing RamzAI can do>", ...]} when the person is vague '
    'about what they want ("help me study", "I have an exam on Friday", "what can you do"). Give 3 or 4 concrete, specific '
    "options written as things the person could say.\n\n"
    "Rules: keep every word short and plain, the way you would speak to a first-year student; never invent facts, names or "
    "numbers; use the person's own words for the topic; write in the person's language; a file the person attached is "
    "referred to as 'your file'. When the message continues an earlier exchange (given as history), use what was already said "
    "and do not ask it again."
)


@app.route("/api/remy/plan", methods=["POST", "OPTIONS"])
@limiter.limit("90 per hour;400 per day", key_func=_account_or_ip_key)
@auth_required
def remy_plan():
    data = request.get_json(silent=True) or {}
    text = re.sub(r"\s+", " ", str(data.get("text") or "")).strip()[:1500]
    if not text:
        return jsonify({"error": "Say what you want RamzAI to do."}), 400
    turns = []
    for t in (data.get("history") or [])[-8:]:
        if isinstance(t, dict) and t.get("role") in ("user", "assistant") and str(t.get("content") or "").strip():
            turns.append({"role": t["role"], "content": str(t["content"]).strip()[:1200]})
    has_file = bool(data.get("has_file"))
    file_name = str(data.get("file_name") or "")[:120]
    # 4.5: the names of other recent files are not used (an older app may still send them)
    context = ("The person attached a file" + (f" called {file_name}" if file_name else "") + ".\n" if has_file else "No file is attached.\n")
    # the exchange so far, then the new message, flattened so the planner sees it in one go
    convo = "".join(("Person: " if t["role"] == "user" else "RamzAI asked: ") + t["content"] + "\n" for t in turns)
    user_msg = context + (("Earlier in this exchange:\n" + convo) if convo else "") + "Person now says: " + text
    try:
        raw = call_claude(REMY_PLAN_SYSTEM, user_msg, max_tokens=500, timeout=40, feature="plan")
    except ConversionError as e:
        return jsonify({"do": "chat", "note": str(e)[:200]})
    obj = _guide_json(raw) if isinstance(raw, str) else None
    if not isinstance(obj, dict) or obj.get("do") not in ("chat", "ask", "plan", "suggest"):
        return jsonify({"do": "chat"})
    out = {"do": obj["do"]}
    clip = lambda v, n: re.sub(r"\s+", " ", str(v or "")).strip()[:n]
    if obj["do"] in ("ask", "suggest"):
        out["say"] = clip(obj.get("say"), 300) or ("What should it be about?" if obj["do"] == "ask" else "Here are a few things I can do.")
        out["options"] = [clip(o, 90) for o in (obj.get("options") or []) if clip(o, 90)][:4]
        out["free"] = bool(obj.get("free", True))
    if obj["do"] == "plan":
        tool = clip(obj.get("tool"), 20)
        if tool not in REMY_PLAN_POINTS and tool not in ("convert", "meeting", "focus", "pdf", "word", "scan", "studio"):
            return jsonify({"do": "chat"})
        params = obj.get("params") if isinstance(obj.get("params"), dict) else {}
        out.update({"tool": tool, "title": clip(obj.get("title"), 140) or "RamzAI will do this", "detail": [clip(d, 160) for d in (obj.get("detail") or []) if clip(d, 160)][:4],
                    "params": {str(k)[:30]: (clip(v, 400) if isinstance(v, str) else v) for k, v in params.items() if k},
                    "next": [clip(n, 90) for n in (obj.get("next") or []) if clip(n, 90)][:3],
                    "points": REMY_PLAN_POINTS.get(tool, 0)})
        if tool in REMY_PLAN_SCHOLAR and not _is_pro(request.current_user) and not _research_open(request.current_user):
            out["scholar"] = REMY_PLAN_SCHOLAR[tool]
    try:
        pts = _points_state(request.current_user)
        out["have"] = pts[0]
    except Exception:
        pass
    return jsonify(out)


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
             "numbered": data.get("numbered") is True,
             # 5.0: the numbers' look (Vancouver, IEEE, the person's own numbered style), the look the person chose for the
             # file, and the citation style it follows (which says where the text is aligned when they didn't choose)
             "marks": data.get("cite_marks") if data.get("cite_marks") in NUM_MARKS + ("ieee",) else None,
             "look": data.get("look") if isinstance(data.get("look"), dict) else None, "style": str(data.get("style") or "apa")[:20],
             "front": data.get("front") if isinstance(data.get("front"), dict) else None},     # 5.0: the title page
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
    if ext not in PPT_EXTS:
        return jsonify({"error": "Choose a PowerPoint file (.pptx, .ppt, .pps, .ppsx, .pptm, .pot, .potx or .odp)."}), 400
    work_dir = tempfile.mkdtemp(prefix=f"mc_slides_{uuid.uuid4().hex[:8]}_")
    try:
        raw = os.path.join(work_dir, "deck." + ext)
        f.save(raw)
        try:
            path = as_pptx(raw, work_dir)
        except ConversionError as e:
            return jsonify({"error": str(e)}), 422
        made = path != raw
        deck = parse_pptx_to_deck(path)
        deck["filename"] = safe_download_name(f.filename)
        if made:
            deck["filename"] = re.sub(r"\.[A-Za-z0-9]+$", "", deck["filename"]) + ".pptx"
        # The file's own design, kept (2.9). Rendering can fail on an unusual file: then it
        # opens in one of our designs, as before.
        original = None
        try:
            original = parse_pptx_original(path, work_dir, to_pdf=lambda src_path, out: _soffice_convert(src_path, "pdf", out))
        except Exception as e:
            logger.warning(f"Keeping the deck's own design failed: {e}")
        out = {"deck": deck, "original": original}
        if made and original:
            with open(path, "rb") as fh:
                out["pptx"] = base64.b64encode(fh.read()).decode("ascii")
        return jsonify(out)
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


# RamzAI inside Slide Studio: one tap to tighten a slide's points, put them in
# simpler words, draft points from its title, or write its speaker notes.
# Like Smart Summarize it's a slide tool, free on every plan, so it doesn't
# use the RamzAI message allowance — only the per-account AI rate limit.
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
    "You are RamzAI, the study assistant in Academic-Home, improving one slide of a student's or lecturer's presentation. "
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
        raise ConversionError("RamzAI couldn't improve this slide — please try again.")
    if not isinstance(data, dict):
        raise ConversionError("RamzAI couldn't improve this slide — please try again.")
    if action == "notes":
        notes = data.get("notes")
        if not isinstance(notes, str) or not notes.strip():
            raise ConversionError("RamzAI couldn't write notes for this slide — please try again.")
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
        raise ConversionError("RamzAI couldn't improve this slide — please try again.")
    return {"bullets": points}


# ======================= 3.14: RamzAI for a whole deck (Slide Studio → All slides) =======================
_ALL_ASSIST = {
    "notes": ("For every slide, write speaker notes: what the presenter says while it is shown — 2 to 4 short paragraphs of "
              "natural spoken English, about 80 to 150 words, covering each point. Reply for each slide with "
              '{"i": <its number>, "notes": "..."}.', 6000),
    "tighten": ("For every slide, rewrite its points so each is short and punchy — ideally under 12 words — keeping every "
                "idea and its meaning; keep sub-points as sub-points (level 1); at most 7 points. Reply for each slide with "
                '{"i": <its number>, "bullets": [{"text": "...", "level": 0}]}.', 4000),
    "spelling": ("For every slide, correct the spelling, grammar and punctuation of its title and points. Change nothing "
                 "else: keep the words, meaning, order and levels. Reply for each slide with "
                 '{"i": <its number>, "title": "...", "bullets": [{"text": "...", "level": 0}]}.', 4000),
    "translate": ("Translate every slide's title, points and speaker notes into LANGUAGE, keeping the order and levels. "
                  'Reply for each slide with {"i": <its number>, "title": "...", "bullets": [{"text": "...", "level": 0}], '
                  '"notes": "..."} (notes only where the slide has them).', 6000),
}
_ALL_LANGS = WORD_AI_LANGS


def _all_points(raw, limit=12):
    out = []
    for b in (raw or [])[:40]:
        if isinstance(b, str):
            b = {"text": b, "level": 0}
        if not isinstance(b, dict):
            continue
        text = str(b.get("text") or "").strip()[:600]
        if not text:
            continue
        try:
            level = max(0, min(int(b.get("level") or 0), 3))
        except (TypeError, ValueError):
            level = 0
        out.append({"text": text, "level": level})
    return out[:limit]


@app.route("/api/slides/assist-all", methods=["POST", "OPTIONS"])
@limiter.limit(AI_LIMIT, key_func=_account_or_ip_key)
@auth_required
@_once
def slides_assist_all_endpoint():
    if request.method == "OPTIONS":
        return "", 204
    data = request.get_json(silent=True) or {}
    action = str(data.get("action") or "")
    if action not in _ALL_ASSIST:
        return jsonify({"error": "Unknown request for RamzAI."}), 400
    lang = str(data.get("lang") or "")
    if action == "translate" and lang not in _ALL_LANGS:
        return jsonify({"error": "Choose a language."}), 400
    slides = []
    for s in (data.get("slides") or [])[:12]:
        if not isinstance(s, dict):
            continue
        try:
            i = int(s.get("i"))
        except (TypeError, ValueError):
            continue
        title = str(s.get("title") or "").strip()[:200]
        pts = _all_points(s.get("bullets"), 40)
        notes = str(s.get("notes") or "").strip()[:2000]
        if title or pts or (action == "translate" and notes):
            slides.append((i, title, pts, notes))
    if not slides:
        return jsonify({"error": "There's nothing on the slides for RamzAI yet."}), 400
    deck = str(data.get("deck_title") or "").strip()[:160]
    lines = [f"Presentation: {deck}"] if deck else []
    for i, title, pts, notes in slides:
        lines.append(f"\nSLIDE {i}\nTitle: {title or '(none)'}")
        if pts:
            lines.append("Points:\n" + "\n".join(("  " * p["level"]) + "- " + p["text"] for p in pts))
        if notes and action == "translate":
            lines.append("Notes: " + notes)
    instruction, max_tokens = _ALL_ASSIST[action]
    instruction = instruction.replace("LANGUAGE", lang)
    try:
        charged = _log_sage_message_and_consume(request.current_user)
    except ConversionError as e:
        return _points_reply(e)
    system = ("You are RamzAI, the study assistant in Academic-Home, improving a whole presentation slide by slide. Reply with "
              'one JSON object only — no preamble, no code fences: {"slides": [ ... one object per slide given, with its number '
              'as "i" ... ]}. ' + instruction + " Only rework what the slides already say or clearly imply: never invent "
              "statistics, dates, citations or facts. Write in the slides' own language unless asked to translate.")
    try:
        raw = call_claude(system, "\n".join(lines) + "\n", max_tokens=max_tokens, feature="slides-all-" + action, timeout=170)
        cleaned = (raw or "").strip()
        if cleaned.startswith("```"):
            cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
            cleaned = re.sub(r"\s*```$", "", cleaned)
        try:
            got = json.loads(cleaned)
        except json.JSONDecodeError:
            m = re.search(r"\{.*\}", cleaned, re.S)
            got = json.loads(m.group(0)) if m else None
        items = got.get("slides") if isinstance(got, dict) else None
        if not isinstance(items, list):
            raise ConversionError("RamzAI couldn't finish that — please try again.")
    except (ConversionError, ValueError) as e:
        _refund_sage_message(request.current_user, charged)
        msg = str(e) if isinstance(e, ConversionError) else "RamzAI couldn't finish that — please try again."
        return jsonify({"error": msg + AI_MSG_NOT_COUNTED}), 502
    asked = {i for i, *_ in slides}
    out = []
    for it in items:
        if not isinstance(it, dict):
            continue
        try:
            i = int(it.get("i"))
        except (TypeError, ValueError):
            continue
        if i not in asked:
            continue
        o = {"i": i}
        if action in ("spelling", "translate") and isinstance(it.get("title"), str) and it["title"].strip():
            o["title"] = it["title"].strip()[:200]
        if action in ("tighten", "spelling", "translate"):
            pts = _all_points(it.get("bullets"), 12)
            if pts:
                o["bullets"] = pts
        if action in ("notes", "translate") and isinstance(it.get("notes"), str) and it["notes"].strip():
            o["notes"] = it["notes"].strip()[:2000]
        if len(o) > 1:
            out.append(o)
    if not out:
        _refund_sage_message(request.current_user, charged)
        return jsonify({"error": "RamzAI couldn't finish that — please try again." + AI_MSG_NOT_COUNTED}), 502
    return jsonify({"slides": out})


@app.route("/api/slides/assist", methods=["POST", "OPTIONS"])
@limiter.limit(AI_LIMIT, key_func=_account_or_ip_key)
@auth_required
def slides_assist_endpoint():
    data = request.get_json(silent=True) or {}
    action = str(data.get("action") or "")
    if action not in _ASSIST_ACTIONS:
        return jsonify({"error": "Unknown request for RamzAI."}), 400
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
        return jsonify({"error": "RamzAI couldn't reach the AI just now — please try again."}), 502


# --------------------------------------------------------- drafting a whole deck
# Starting a new deck in Slide Studio no longer has to mean a blank slide:
# given a topic, RamzAI drafts a full outline (title slide plus content slides)
# that opens straight into the editor to fix up. It uses one RamzAI message per
# draft (the same allowance and Pay-per-Use messages as RamzAI chat), since a
# whole deck is a substantial AI request; editing the deck afterwards, and
# RamzAI's per-slide help, stay free.
_DRAFT_MAX_SLIDES = 10
_DRAFT_SYSTEM = (
    "You are RamzAI, the study assistant in Academic-Home, drafting a full slide-deck outline from a short topic someone "
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
        raise ConversionError("RamzAI couldn't draft this topic — please try again, or add a bit more detail.")

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
        raise ConversionError("RamzAI couldn't draft this topic — please try again, or add a bit more detail.")
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
    "You are RamzAI, the assistant in Academic-Home, writing a COMPLETE slide deck from a short request, for a {group} to present. "
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


def _chart_nearest(x, m, cats, common, words):
    """Which category a number in an abstract sentence is about: the nearest words that set a category apart, looking
    back ("tutored pupils gained 12% in fluency"), or ahead after "of", "among" or "for" ("12% of tutored pupils")."""
    low = x.lower()
    ahead = bool(re.match(r"\s?(?:%|percent|per cent)?\s*(?:of|among|for)\b", low[m.end():]))
    best = None
    for cj, c in enumerate(cats):
        for w in words(c) - common:
            for mm in re.finditer(r"\b" + re.escape(w) + r"\b", low):
                d = (mm.start() - m.end()) if ahead else (m.start() - mm.end())
                if d < 0:
                    continue
                if best is None or d < best[0]:
                    best = (d, cj)
    if best is None:                       # nothing on that side: the nearest on the other
        for cj, c in enumerate(cats):
            for w in words(c) - common:
                for mm in re.finditer(r"\b" + re.escape(w) + r"\b", low):
                    d = abs(mm.start() - m.start())
                    if best is None or d < best[0]:
                        best = (d, cj)
    return best[1] if best else None


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
    sentences = [x for x in re.split(r"(?<=[.!?])\s+", abstract.replace(",", "")) if x.strip()]
    words = lambda t: {w for w in re.findall(r"[a-z]{4,}", str(t).lower())}
    # 4.5.1: each number must be stated about its own bar: the words that set its category apart from the others in the
    # same sentence (or, when nothing does, the series' name), any year the category names, and the number must not
    # count something else there ("40 schools" is not a fluency gain of 40)
    common = set.intersection(*(words(c) for c in cats)) if len(cats) > 1 else set()
    for s in series:
        sw = words(s["name"]) - {"value", "values"}
        for ci, v in enumerate(s["values"]):
            forms = {f"{v:g}"} | ({str(int(v))} if v == int(v) else set())
            dist = words(cats[ci]) - common
            years = set(re.findall(r"\b(?:1[5-9]|20)\d{2}\b", cats[ci]))
            label = words(cats[ci]) | sw
            ok = False
            for x in sentences:
                xw = words(x)
                if (dist and not dist & xw) or (not dist and sw and not sw & xw) or (years and not all(y in x for y in years)):
                    continue
                for m in re.finditer(r"(?<![\d.])(\d+(?:\.\d+)?)(?![\d])", x):
                    if m.group(1) not in forms or m.group(1) in years:
                        continue
                    if len(cats) > 1 and dist and _chart_nearest(x, m, cats, common, words) != ci:
                        continue                 # "tutored 12% and untutored 5%": 5 is the untutored bar's
                    after = x[m.end():]
                    if re.match(r"\s?(?:%|percent|per cent)", after, re.I):
                        ok = True
                        break
                    nxt = re.match(r"\s*([A-Za-z][A-Za-z\-]*)", after)
                    u = nxt.group(1).lower() if nxt else ""
                    if not u or u in _C45_STOP or any(u[:4] == w[:4] for w in label):
                        ok = True
                        break
                if ok:
                    break
            if not ok:
                return None                  # a number the paper doesn't state, or states about something else: no chart
    kind = spec.get("chart_type") if spec.get("chart_type") in ("column", "bar", "line", "pie") else "column"
    return {"type": "chart", "chart_type": kind, "categories": cats, "series": series, "title": str(spec.get("title") or "")[:80]}


def _c45_deck_verify(slides, sources, style, topic, deadline):
    """4.5.1: a deck's points are checked like a piece of writing. A point that states something keeps its citation only
    when that paper supports it (or gets the paper that does); a point no paper supports goes, and so does one that
    copies ten or more words of an abstract. A point that states no fact (a topic, an activity, a question) stays."""
    items = [b for sl in slides for name in ("bullets", "left", "right") for b in (sl.get(name) or [])]
    if not items or not sources:
        return True
    if time.time() > deadline - 12:
        return False
    lines, mark = _c34_write_parts(sources)

    def own(t):
        return {f"S{i + 1}" for i in keep_known_citations(t, sources, style)[1]}

    def bare(t):
        out = re.sub(r"[ \t]*\(([^()]{2,400})\)", lambda m: "" if keep_known_citations("(" + m.group(1) + ")", sources, style)[1] else m.group(0), t)
        return re.sub(r"[ \t]+([.,;:!?])", r"\1", out).strip()

    todo = [b for b in items if own(b["text"]) or _c34_needs(b["text"])]
    if not todo:
        return True
    shown = [bare(b["text"]) for b in todo]
    context = "\n".join(shown)[:6000]
    chunks = [list(range(i, min(i + C34_CHUNK, len(todo)))) for i in range(0, len(todo), C34_CHUNK)][:C34_MAX_CHUNKS]
    results = [None] * len(chunks)
    with _futures.ThreadPoolExecutor(max_workers=3) as pool:
        jobs = {pool.submit(_c34_check, "deck", topic, lines, [shown[i] for i in ch], context, deadline): ci for ci, ch in enumerate(chunks)}
        for job, ci in jobs.items():
            try:
                results[ci] = job.result(timeout=max(2, deadline - time.time() + 3))
            except Exception:
                results[ci] = None
    dropped = moved = given = 0
    for ci, ch in enumerate(chunks):
        r = results[ci]
        if not r or r["skip"]:
            continue
        for j, k in enumerate(ch):
            b = todo[k]
            if j in r["keep"] and not own(b["text"]):
                continue
            new = set(r["cite"].get(j, []))
            old = own(b["text"])
            if not new:
                b["text"] = ""; dropped += 1
                continue
            keep = (old & new) or new
            ev = " ".join(sources[int(x[1:]) - 1].get("abstract") or "" for x in keep if x[1:].isdigit() and 1 <= int(x[1:]) <= len(sources))
            if _c45_copied(bare(b["text"]), ev, run=10):
                b["text"] = ""; dropped += 1
                continue
            if keep == old:
                continue
            m = mark(sorted(keep))
            if m:
                b["text"] = _c34_put(bare(b["text"]), m)
                moved += 1 if old else 0
                given += 0 if old else 1
            else:
                b["text"] = ""; dropped += 1
    for sl in slides:
        for name in ("bullets", "left", "right"):
            if sl.get(name):
                sl[name] = [b for b in sl[name] if str(b.get("text") or "").strip()]
    logger.info(f"Citations (deck): {len(todo)} points checked → {given} given a source, {moved} given the right one, {dropped} taken out")
    return all(r and not r["skip"] for r in results) and len(todo) <= C34_CHUNK * C34_MAX_CHUNKS


def _parse_full_deck(raw, group, sources, style="apa", deck_max=None, topic="", deadline=None):
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
        raise ConversionError("RamzAI couldn't write this deck — please try again, or add a bit more detail.")
    ts_raw = data.get("title_slide") if isinstance(data.get("title_slide"), dict) else {}
    clip = lambda v, n: re.sub(r"\s+", " ", str(v or "")).strip()[:n]
    if group == "tutor":
        title_slide = {"course": clip(ts_raw.get("course"), 80), "code": clip(ts_raw.get("code"), 30), "subtitle": clip(ts_raw.get("subtitle"), 120),
                       "lecturer": clip(ts_raw.get("lecturer"), 80), "date": clip(ts_raw.get("date"), 40)}
    else:
        lines = [clip(l, 120) for l in (ts_raw.get("lines") or []) if clip(l, 120)][:3]
        title_slide = {"lines": lines, "main": clip(ts_raw.get("main"), 200), "sub": clip(ts_raw.get("sub"), 200)}
    def check(points):
        """Only real citations stay; a point that gives a figure needs a source that gives that figure."""
        kept = []
        for b in points:
            b["text"], c = keep_known_citations(b["text"], sources, style)
            nums = _c45_numbers(b["text"])
            if nums and sources:
                ev = " ".join((sources[i].get("title") or "") + " " + (sources[i].get("abstract") or "") for i in c)
                if not c or not _c45_numbers_in(nums, ev):
                    continue
            if b["text"].strip():
                kept.append(b)
        return kept

    slides = []
    deck_max = deck_max or _DECK_MAX
    for sd in (data.get("slides") or [])[:deck_max + 4]:
        if not isinstance(sd, dict):
            continue
        kind = sd.get("kind") if sd.get("kind") in ("content", "columns", "section") else "content"
        title = clip(sd.get("title"), 140)
        notes = keep_known_citations(clip(sd.get("notes"), 900), sources or [], style)[0]   # no papers: no citation can stand
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
    unchecked = False
    if sources and deadline:
        unchecked = not _c45_deck_verify(slides, sources, style, topic, deadline)
        slides = [sl for sl in slides if sl["kind"] == "section" or sl.get("media") or sl.get("bullets") or sl.get("left") or sl.get("right")]
    body = slides[:deck_max]
    if not any(s["kind"] != "section" for s in body):
        raise ConversionError("RamzAI couldn't write this deck — please try again, or add a bit more detail.")
    # the References slide lists the papers the slides that are left cite (not those of slides that were cut)
    cited = set()
    for sl in body:
        for b in (sl.get("bullets") or []) + (sl.get("left") or []) + (sl.get("right") or []):
            cited.update(keep_known_citations(b.get("text") or "", sources, style)[1])
        if sl.get("notes"):
            cited.update(keep_known_citations(sl["notes"], sources, style)[1])
    refs = build_references_list(sources, cited) if sources else []
    lines = lambda rs: [re.sub(r"[*_]", "", r["text"]) + (" " + r["url"] if r.get("url") else "") + (" " + r["after"] if r.get("after") else "") for r in rs]

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
        out = plain("Presented by", items("presenters", 10, 120) or ["Your name and index number"], 6)
        out += plain("Presentation Outline", outline, 6) + body + plain("References", lines(refs), 4)
    return {"template": ("tutor" if group == "tutor" else "student") + "-plain", "title_slide": title_slide, "slide_numbers": True,
            "slides": out, "unchecked": unchecked}, refs


@app.route("/api/slides/draft", methods=["POST", "OPTIONS"])   # a whole deck: POINTS_DECK
@limiter.limit(AI_LIMIT, key_func=_account_or_ip_key)
@auth_required
@_once
def slides_draft_endpoint():
    data = request.get_json(silent=True) or {}
    group = "tutor" if data.get("group") == "tutor" else "student"
    topic = re.sub(r"\s+", " ", str(data.get("topic") or "")).strip()[:300]
    if not topic:
        return jsonify({"error": "Tell RamzAI what the deck is about first."}), 400
    started = time.time()
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
    # A whole drafted deck is real AI work, so it draws on the same RamzAI
    # messages as chatting with RamzAI: the month's allowance first, then any
    # Pay-per-Use messages. Checked before the AI call (never spend on a
    # request already over its limit), and given back below if no usable
    # deck arrives — a message only counts when the answer does.
    try:
        charged_from = _log_sage_message_and_consume(request.current_user, POINTS_DECK)
    except ConversionError as e:
        return _points_reply(e)
    try:
        raw = call_claude(system, f"Request: {topic}", max_tokens=12000 if pro else 9000, timeout=150 if pro else None, feature="deck")
        deck, refs = _parse_full_deck(raw, group, sources, deck_max=content_max + 2, topic=topic, deadline=started + 160)
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
# account goes with it — conversion history and RamzAI's memory notes (removed
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
            # 4.5: the day counts of invitation emails are tied to the account number and have no link to the account row
            cur.execute("SELECT to_regclass('public.meeting_invites') AS t")
            if (cur.fetchone() or {}).get("t"):
                cur.execute("DELETE FROM meeting_invites WHERE user_id = %s", (user["id"],))
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
    cur.execute("ALTER TABLE meetings ADD COLUMN IF NOT EXISTS captions BOOLEAN NOT NULL DEFAULT FALSE")   # 3.9: live captions
    for col in ("starts_at TIMESTAMPTZ", "duration INTEGER", "agenda TEXT NOT NULL DEFAULT ''", "repeat TEXT NOT NULL DEFAULT ''",
                "tz TEXT NOT NULL DEFAULT ''", "notes_sent_at TIMESTAMPTZ"):
        cur.execute("ALTER TABLE meetings ADD COLUMN IF NOT EXISTS " + col)       # 3.14: a meeting's time, agenda
    cur.execute("CREATE TABLE IF NOT EXISTS meeting_invites (user_id INTEGER NOT NULL, day DATE NOT NULL, n INTEGER NOT NULL DEFAULT 0, "
                "PRIMARY KEY (user_id, day))")
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
                       "captions": bool(m.get("captions")),
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
            "role": p["role"], "name": p["name"], "e2ee": m["e2ee"], "title": m["title"], "code": m["code"],
            "agenda": m.get("agenda") or ""}      # 3.14: everyone let in sees the agenda


def _meet_running(m):
    return bool(m.get("started_at") and not m.get("ended_at"))


def _meet_public(m, host_name=None, full=False):
    state = "live" if _meet_running(m) else ("ended" if m["ended_at"] else "waiting")
    out = {"code": m["code"], "title": m["title"], "host": host_name, "state": state, "passcode": bool(m["passcode_hash"]),
           "lobby": m["lobby"], "e2ee": m["e2ee"], "locked": m["locked"], "maxPeople": m["max_people"],
           "minutes": m["limit_minutes"], "endsAt": m["ends_at"].isoformat() if m["ends_at"] else None}
    # 3.14: when it's on, and what it's about
    s = m.get("starts_at")
    nxt = _meet_next(m)
    out.update(startsAt=s.isoformat() if s else None, next=nxt.isoformat() if nxt else None, duration=m.get("duration"),
               repeat=m.get("repeat") or "", tz=m.get("tz") or "", agenda=_meet_agenda_for(m, host=full))
    return out


# ======================= 3.14: scheduled meetings =======================
MEET_INVITES_A_DAY = 60


def _meet_next(m):
    """The next time a scheduled meeting is on: its start, or — every week — the next week's (until 3 hours after)."""
    s = m.get("starts_at") if isinstance(m, dict) else None
    if not s:
        return None
    if (m.get("repeat") or "") == "weekly":
        now = datetime.now(timezone.utc)
        if s < now - timedelta(hours=3):
            weeks = int((now - timedelta(hours=3) - s) / timedelta(weeks=1)) + 1
            s = s + timedelta(weeks=weeks)
    return s


def _meet_when(data, m=None):
    """The schedule fields from a request: (starts_at, duration, repeat, tz, agenda) — each None when not given."""
    starts = None
    raw = _text(data.get("starts_at"))
    if raw:
        try:
            starts = datetime.fromisoformat(raw.replace("Z", "+00:00"))
        except ValueError:
            raise MeetError("That date and time couldn’t be read.", 400)
        if starts.tzinfo is None:
            raise MeetError("That date and time couldn’t be read.", 400)
        now = datetime.now(timezone.utc)
        if starts < now - timedelta(days=1) or starts > now + timedelta(days=400):
            raise MeetError("Choose a time from today to a year ahead.", 400)
    dur = None
    if data.get("duration") is not None:
        try:
            dur = max(10, min(480, int(data.get("duration"))))
        except (TypeError, ValueError):
            dur = 60
    rep = "weekly" if data.get("repeat") == "weekly" else ("" if "repeat" in data else None)
    tz = re.sub(r"[^A-Za-z0-9_+\-/]", "", _text(data.get("tz")))[:64] if data.get("tz") is not None else None
    agenda = _text(data.get("agenda")).replace("\r", "").strip()[:4000] if data.get("agenda") is not None else None
    return starts, dur, rep, tz, agenda


def _meet_save_when(cur, m, data):
    starts, dur, rep, tz, agenda = _meet_when(data, m)
    sets, args = [], []
    for col, val in (("starts_at", starts), ("duration", dur), ("repeat", rep), ("tz", tz), ("agenda", agenda)):
        if val is not None:
            sets.append(col + " = %s")
            args.append(val)
    if data.get("unschedule"):
        sets += ["starts_at = NULL", "repeat = ''"]
    if not sets:
        return m
    cur.execute("UPDATE meetings SET " + ", ".join(sets) + " WHERE id = %s RETURNING *", args + [m["id"]])
    return cur.fetchone()


def _meet_tz(m):
    """The host's time zone and its name — UTC when it isn't known here."""
    try:
        from zoneinfo import ZoneInfo
        return ZoneInfo(m.get("tz") or "UTC"), (m.get("tz") or "UTC")
    except Exception:
        return timezone.utc, "UTC"


def _meet_agenda_for(m, host=False):
    """A meeting with a passcode keeps its agenda for the people who got in (and its host)."""
    return (m.get("agenda") or "") if (host or not m.get("passcode_hash")) else ""


def _meet_when_words(m):
    """"Friday 30 October 2026, 3:00–4:00 pm (Africa/Accra)" — in the host's time zone."""
    s = _meet_next(m)
    if not s:
        return ""
    z, zname = _meet_tz(m)
    a = s.astimezone(z)
    b = a + timedelta(minutes=m.get("duration") or 60)
    t = lambda d: d.strftime("%I:%M").lstrip("0") + ("am" if d.hour < 12 else "pm")
    words = a.strftime("%A ") + str(a.day) + a.strftime(" %B %Y") + ", " + t(a) + "–" + t(b)
    return words + " (" + zname.replace("_", " ") + ")" + (", every week" if m.get("repeat") == "weekly" else "")


def _meet_gcal(m, link):
    from urllib.parse import quote
    s = _meet_next(m)
    if not s:
        return ""
    e = s + timedelta(minutes=m.get("duration") or 60)
    f = lambda d: d.astimezone(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    details = "Join: " + link + ("\n\nAgenda:\n" + m["agenda"] if m.get("agenda") else "")
    return ("https://calendar.google.com/calendar/render?action=TEMPLATE&text=" + quote(m["title"]) + "&dates=" + f(s) + "/" + f(e) +
            "&details=" + quote(details) + ("&recur=RRULE:FREQ%3DWEEKLY" if m.get("repeat") == "weekly" else ""))


def _meet_ics(m, link, agenda=""):
    s = _meet_next(m) or datetime.now(timezone.utc)
    e = s + timedelta(minutes=m.get("duration") or 60)
    f = lambda d: d.astimezone(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    esc = lambda t: str(t or "").replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\n", "\\n")
    lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Academic-Home//Meetings//EN", "CALSCALE:GREGORIAN", "METHOD:PUBLISH", "BEGIN:VEVENT",
             "UID:" + m["code"] + "@academic-home", "DTSTAMP:" + f(datetime.now(timezone.utc)), "DTSTART:" + f(s), "DTEND:" + f(e),
             "SUMMARY:" + esc(m["title"]), "URL:" + link, "LOCATION:" + esc(link),
             "DESCRIPTION:" + esc("Join: " + link + ("\n\nAgenda:\n" + agenda if agenda else ""))]
    if m.get("repeat") == "weekly":
        lines.append("RRULE:FREQ=WEEKLY")
    lines += ["BEGIN:VALARM", "TRIGGER:-PT10M", "ACTION:DISPLAY", "DESCRIPTION:" + esc(m["title"] + " starts in 10 minutes"), "END:VALARM",
              "END:VEVENT", "END:VCALENDAR"]
    out = []
    for ln in lines:                       # lines of at most 75 octets, folded as the standard says
        b = ln.encode("utf-8")
        while len(b) > 75:
            cut = 75
            while cut > 0 and (b[cut] & 0xC0) == 0x80:
                cut -= 1
            out.append(b[:cut].decode("utf-8"))
            b = b" " + b[cut:]
        out.append(b.decode("utf-8"))
    return "\r\n".join(out) + "\r\n"


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
@app.route("/api/meet/<code>/ics", methods=["GET", "OPTIONS"])
@limiter.limit("120 per hour", key_func=_meet_code_key)
@_meet_reply
def meet_ics(cur, code):
    m = _meet_get(cur, code)
    link = f"{FRONTEND_URL}/?meet={m['code']}"
    resp = app.response_class(_meet_ics(m, link, _meet_agenda_for(m)), mimetype="text/calendar")
    resp.headers["Content-Disposition"] = "attachment; filename=\"" + safe_download_name((m["title"] or "Meeting")[:60] + ".ics") + "\""
    return resp


@app.route("/api/meet/<code>/invite", methods=["POST", "OPTIONS"])
@limiter.limit("12 per hour", key_func=_account_or_ip_key)
@auth_required
@_once
@_meet_reply
def meet_invite(cur, code):
    user = request.current_user
    m = _meet_get(cur, code)
    if user["id"] != m["host_id"]:
        raise MeetError("Only the host can send the invitation.", 403)
    if m["e2ee"]:
        raise MeetError("An end-to-end encrypted meeting’s link carries its key, so Academic-Home can’t email it — share the link yourself.", 400)
    if user.get("email_verified", True) is False:
        raise MeetError("Confirm your own email first — then you can email invitations.", 403)
    if not _email_ready():
        raise MeetError("Email isn’t set up on the server yet — share the link instead.", 503)
    data = _json_body()
    raw = data.get("emails") if isinstance(data.get("emails"), list) else re.split(r"[\s,;]+", _text(data.get("emails")))
    emails = []
    for e in raw[:40]:
        e = _text(e).strip().lower()
        if e and re.fullmatch(r"[^@\s<>\"']+@[^@\s<>\"']+\.[a-z]{2,}", e) and e not in emails:
            emails.append(e)
    if not emails:
        raise MeetError("Type at least one email address.", 400)
    if len(emails) > 20:
        raise MeetError("Up to 20 people at a time.", 400)
    cur.execute("INSERT INTO meeting_invites (user_id, day, n) VALUES (%s, CURRENT_DATE, %s) ON CONFLICT (user_id, day) "
                "DO UPDATE SET n = meeting_invites.n + EXCLUDED.n RETURNING n", (user["id"], len(emails)))
    if cur.fetchone()["n"] > MEET_INVITES_A_DAY:
        raise MeetError(f"That’s more than {MEET_INVITES_A_DAY} invitations today — share the link instead, or try tomorrow.", 429)
    link = f"{FRONTEND_URL}/?meet={m['code']}"
    host = (user.get("name") or "Someone").strip()[:60]
    when = _meet_when_words(m)
    gcal = _meet_gcal(m, link)
    base = (os.environ.get("RENDER_EXTERNAL_URL") or request.url_root).rstrip("/")     # this server's own address
    if base.startswith("http://") and not re.match(r"http://(localhost|127\.)", base):
        base = "https://" + base[7:]
    ics = f"{base}/api/meet/{m['code']}/ics"
    subject = ("Invitation: " + m["title"] + (" — " + when.split(" (")[0] if when else ""))[:180]
    body = (f"{host} invited you to a meeting on Academic-Home.\n\n{m['title']}\n" + (f"When: {when}\n" if when else "") +
            f"Join: {link}\nCode: {m['code']}\n" + (("Passcode: ask " + host + "\n") if m["passcode_hash"] else "") +
            (f"\nAgenda:\n{m['agenda']}\n" if m.get("agenda") else "") +
            ((f"\nAdd it to your calendar:\nGoogle Calendar: {gcal}\nOther calendars (a .ics file): {ics}\n") if when else "") +
            "\nYou can join from a phone or a computer — no app needed.\n")
    from html import escape as h
    html = ("<div style=\"font-family:Arial,sans-serif;font-size:15px;line-height:1.5;color:#111\">" +
            f"<p>{h(host)} invited you to a meeting on Academic-Home.</p><h2 style=\"margin:0 0 6px\">{h(m['title'])}</h2>" +
            (f"<p style=\"margin:0 0 12px\"><b>When:</b> {h(when)}</p>" if when else "") +
            f"<p><a href=\"{h(link)}\" style=\"display:inline-block;background:#1B4A9C;color:#fff;padding:10px 18px;border-radius:8px;text-decoration:none;font-weight:bold\">Join the meeting</a></p>" +
            f"<p>Code: <b>{h(m['code'])}</b></p>" +
            (f"<p><b>Agenda</b><br>{h(m['agenda']).replace(chr(10), '<br>')}</p>" if m.get("agenda") else "") +
            (f"<p>Add it to your calendar: <a href=\"{h(gcal)}\">Google Calendar</a> · <a href=\"{h(ics)}\">Other calendars</a></p>" if when else "") +
            "</div>")
    sent = 0
    for e in emails:
        try:
            send_email(e, subject, body, html)
            sent += 1
        except Exception as ex:
            logger.warning(f"Meet: an invitation wasn't sent ({ex.__class__.__name__})")
    if not sent:
        raise MeetError("The invitations couldn’t be sent just now — please try again.", 502)
    return jsonify({"sent": sent})


@app.route("/api/meet/agenda", methods=["POST", "OPTIONS"])
@limiter.limit(AI_LIMIT, key_func=_account_or_ip_key)
@auth_required
@_once
def meet_agenda():
    if request.method == "OPTIONS":
        return "", 204
    data = request.get_json(silent=True) or {}
    title = str(data.get("title") or "").strip()[:120]
    about = str(data.get("about") or "").strip()[:1500]
    try:
        minutes = max(10, min(480, int(data.get("minutes") or 60)))
    except (TypeError, ValueError):
        minutes = 60
    if not title and not about:
        return jsonify({"error": "Give the meeting a name first — RamzAI writes the agenda from it."}), 400
    try:
        charged = _log_sage_message_and_consume(request.current_user)
    except ConversionError as e:
        return _points_reply(e)
    system = ("You are RamzAI, in Academic-Home's meetings, writing the agenda for a meeting, class or study group. Reply with the "
              f"agenda only: 4 to 7 lines, each \"- <minutes> min — <item>\", the minutes adding up to about {minutes}; start "
              "with a short welcome and end with questions or next steps. Plain words, no heading, no other text. Never invent "
              "names, dates or facts that weren't given.")
    try:
        out = call_claude(system, "Meeting: " + (title or "(no name)") + "\nLength: " + str(minutes) + " minutes" + ("\nAbout: " + about if about else ""),
                          max_tokens=500, feature="meet-agenda")
    except ConversionError as e:
        _refund_sage_message(request.current_user, charged)
        return jsonify({"error": str(e) + AI_MSG_NOT_COUNTED}), 502
    lines = [l.strip() for l in (out or "").splitlines() if l.strip()]
    lines = [l if l.startswith("- ") else "- " + l.lstrip("-•* ") for l in lines if not l.lower().startswith(("here", "agenda:"))][:10]
    if not lines:
        _refund_sage_message(request.current_user, charged)
        return jsonify({"error": "RamzAI couldn’t write the agenda — please try again." + AI_MSG_NOT_COUNTED}), 502
    return jsonify({"agenda": "\n".join(lines)})


@app.route("/api/meet/<code>/notes/send", methods=["POST", "OPTIONS"])
@limiter.limit("10 per hour", key_func=_account_or_ip_key)
@auth_required
@_once
@_meet_reply
def meet_notes_send(cur, code):
    user = request.current_user
    m = _meet_get(cur, code, lock=True)
    if user["id"] != m["host_id"]:
        raise MeetError("Only the host can send the notes.", 403)
    if m["e2ee"]:
        raise MeetError("Notes can’t be emailed from an end-to-end encrypted meeting.", 400)
    notes = _text(_json_body().get("notes")).replace("\r", "").strip()[:12000]
    if not notes:
        raise MeetError("There are no notes to send yet.", 400)
    if m.get("notes_sent_at") and datetime.now(timezone.utc) - m["notes_sent_at"] < timedelta(minutes=10):
        raise MeetError("The notes were sent a moment ago — try again in a few minutes.", 429)
    cur.execute("SELECT DISTINCT u.email FROM meeting_people p JOIN users u ON u.id = p.user_id WHERE p.meeting_id = %s "
                "AND p.joined_at IS NOT NULL AND u.email IS NOT NULL LIMIT 150", (m["id"],))
    to = [r["email"] for r in cur.fetchall()]
    if user.get("email") and user["email"] not in to:
        to.append(user["email"])
    sent = 0
    host = (user.get("name") or "The host").strip()[:60]
    for e in to:
        try:
            send_email(e, ("Notes from " + m["title"])[:180], f"{host} shared the notes from “{m['title']}”, written by RamzAI.\n\n{notes}\n")
            sent += 1
        except Exception as ex:
            logger.warning(f"Meet: notes weren't sent to one person ({ex.__class__.__name__})")
    cur.execute("UPDATE meetings SET notes_sent_at = now() WHERE id = %s", (m["id"],))
    return jsonify({"sent": sent})


@app.route("/api/meet/config", methods=["GET", "OPTIONS"])
def meet_config():
    if request.method == "OPTIONS":
        return "", 204
    return jsonify({"enabled": _meet_on(), "maxPeople": MEET_MAX_PEOPLE, "freeMinutes": MEET_FREE_MINUTES,
                    "proMinutes": MEET_PRO_MINUTES, "audienceOver": MEET_AUDIENCE_OVER})


@app.route("/api/meet", methods=["POST", "OPTIONS"])
@limiter.limit("20 per hour", key_func=_account_or_ip_key)
@auth_required
@_once
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
    m = _meet_save_when(cur, m, data)                 # 3.14: its time and agenda, when given
    logger.info(f"Meet: meeting {m['id']} made (lobby={m['lobby']} e2ee={m['e2ee']} speak={speak} minutes={minutes} scheduled={bool(m.get('starts_at'))})")
    return jsonify({"meeting": _meet_public(m, user.get("name"), full=True), "link": f"{FRONTEND_URL}/?meet={code}"})


@app.route("/api/meet/mine", methods=["GET", "OPTIONS"])
@auth_required
@_meet_reply
def meet_mine(cur):
    user = request.current_user
    cur.execute("SELECT m.*, (SELECT count(*) FROM meeting_people p WHERE p.meeting_id = m.id AND p.joined_at IS NOT NULL) AS came "
                "FROM meetings m WHERE host_id = %s ORDER BY (starts_at IS NOT NULL AND (repeat = 'weekly' OR "
                "starts_at > now() - interval '3 hours')) DESC, created_at DESC LIMIT 30", (user["id"],))   # 3.14: what's coming first
    out = []
    for m in cur.fetchall():
        d = _meet_public(m, user.get("name"), full=True)
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
    if d["youHost"]:
        d["agenda"] = m.get("agenda") or ""
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
    if action == "schedule":                          # 3.14: its name, time and agenda
        if actor["role"] != "host":
            raise MeetError("Only the host can do that.", 403)
        title = data.get("title")
        if title is not None:
            cur.execute("UPDATE meetings SET title = %s WHERE id = %s RETURNING *", (_meet_clean_name(title, fallback=m["title"])[:80], m["id"]))
            m = cur.fetchone()
        m = _meet_save_when(cur, m, data)
        return jsonify({"meeting": _meet_public(m, full=True)})
    if not _meet_running(m) and action not in ("end", "delete", "lock", "unlock", "lobby_on", "lobby_off", "chat_on", "chat_off",
                                               "speak_all", "speak_chosen", "admit", "deny", "admit_all", "cc_on", "cc_off"):
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
             "chat_on": ("chat", True), "chat_off": ("chat", False), "cc_on": ("captions", True), "cc_off": ("captions", False)}
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
    return jsonify({"ok": True, "meeting": _meet_public(cur.fetchone(), full=True)})


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
@_once
@_meet_reply
def meet_notes(cur, code):
    """RamzAI's notes of a meeting, from its chat, questions and slide titles — each point says where it came from."""
    data = _json_body()
    m = _meet_get(cur, code)
    if request.current_user["id"] != m["host_id"]:
        raise MeetError("Only the host can ask RamzAI for the notes.", 403)
    if m["e2ee"]:
        # 4.5: an end-to-end encrypted meeting's chat and words stay between the people in it — never sent to us or the AI
        raise MeetError("RamzAI’s notes aren’t available in an end-to-end encrypted meeting: its chat and words stay "
                        "between the people in it.", 400)
    lines = []
    for it in (data.get("chat") or [])[:400]:
        if isinstance(it, dict) and _text(it.get("text")).strip():
            lines.append(f"[{_text(it.get('time'))[:5]} {_meet_clean_name(it.get('name'))}] {_text(it.get('text'))[:600]}")
    slides = [f"Slide {i + 1}: {_text(s)[:200]}" for i, s in enumerate((data.get("slides") or [])[:200]) if _text(s).strip()]
    said = []
    for it in (data.get("said") or [])[:1200]:      # 3.9: the live captions the host's device saw
        if isinstance(it, dict) and _text(it.get("text")).strip():
            said.append(f"[{_text(it.get('time'))[:5]} {_meet_clean_name(it.get('name'))}] {_text(it.get('text'))[:400]}")
    if not lines and not slides and not said:
        raise MeetError("There’s nothing to make notes from yet — no chat, captions or slides.", 400)
    material = (("WHAT WAS SAID (live captions):\n" + "\n".join(said)[:40000] + "\n\n") if said else "") + \
        ("CHAT:\n" + "\n".join(lines)[:30000] if lines else "") + ("\n\nSLIDES SHOWN:\n" + "\n".join(slides) if slides else "")
    try:
        charged = _log_sage_message_and_consume(request.current_user)
    except ConversionError as e:
        raise MeetError(str(e), 429)
    system = ("You write clear meeting notes for an Academic-Home meeting called \"" + m["title"] + "\" from what was said (live "
              "captions, when they were on — speech turned into text, so expect small mistakes and fix obvious ones only), its chat "
              "and the titles of the slides that were shown. Use only that material — never add facts it doesn't contain. Every point "
              "ends with where it came from in square brackets: the person and time for what was said or written — like [Ama 10:42] — "
              "or the slide, like [Slide 3]. Sections: "
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
                        "AND COALESCE(started_at, created_at) < now() - make_interval(days => %s) "
                        "AND (starts_at IS NULL OR starts_at < now() - make_interval(days => %s))", (MEET_KEEP_DAYS, MEET_KEEP_DAYS))
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


def _once_tidy():
    """4.9: answers kept for a request sent again are deleted when they're 50 minutes old. Returns how many went."""
    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute(f"DELETE FROM request_once WHERE created_at < now() - interval '{ONCE_KEEP}'")
            return cur.rowcount
    finally:
        conn.close()


def _once_tidy_loop():
    """Every 5 minutes, so nothing is kept an hour."""
    time.sleep(45)
    while True:
        try:
            _once_tidy()
        except Exception as e:
            logger.warning(f"once: tidy-up failed: {e}")
        time.sleep(300)


if "fake" not in os.environ.get("DATABASE_URL", "fake"):
    _threading.Thread(target=_once_tidy_loop, daemon=True, name="once-tidy").start()


init_db()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    app.run(host="0.0.0.0", port=port)
