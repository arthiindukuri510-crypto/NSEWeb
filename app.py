"""
NSE Company Terminal - combined Flask app
-------------------------------------------
One search box (company name OR NSE symbol) that brings back, for that
company: CMP / 5-day low / 20-day low, and a single merged news feed
where each item can carry BOTH a Category tag (Penalty / Dividend /
AGM_or_EGM / Financial_Results / Others) and a Sentiment tag
(positive / negative / neutral) side by side.

This merges what used to be three separate apps:
  - the company-news search app (Company_News.xlsx + Sentiment + Category)
  - the NSE lookup app (Low_Price.xlsx)
  - the NSE symbol master (EQUITY_L.csv) used to resolve names -> symbols

FILES THIS EXPECTS  (edit the CONFIG block below)
--------------------------------------------------
  NEWS_FILE (required) - your Company_News.xlsx
      columns: Company Name, Date, Summary, Sentiment, Category
      Company Name is only filled on the first row of each company's
      block (merged-cell style export) - forward-filled automatically.
      Category is read straight from this file now (Penalty / Dividend
      / AGM or EGM / Financial Results / Others) - no separate
      symbol-matched category file needed any more.

  SYMBOL_MASTER_FILE (recommended) - NSE's EQUITY_L.csv
      columns: SYMBOL, NAME OF COMPANY, SERIES, DATE OF LISTING, ...
      Used to resolve each news company name to its NSE symbol and merge
      spelling/casing variants ("Aarti Drugs" / "AARTIDRUGS" / "Aarti
      Drugs Limited") into one entity. Without this file, search still
      works by folded company name, just without a symbol.

  LOW_PRICE_FILE (optional) - Low_Price.xlsx
      columns: SYMBOL, LAST_PRICE, 5Day_Low, 20Day_Low
      Powers the CMP / 5-day low / 20-day low stat row. Set to None to skip.
      CMP shown to the user is simply LAST_PRICE from this file - no
      live/streaming price lookup. Refresh this file (and restart the
      app, or re-run whatever regenerates it) whenever you want newer
      prices to show up.

  TREND_3DAY_FILE / TREND_5DAY_FILE (optional) - uptrend_output.xlsx /
      uptrend_5day.xlsx (output of your trend_uptrend_only.py /
      trend_uptrend_5day.py scripts)
      columns: SYMBOL, DATE, TREND, SUPPORT_HIGH, SUPPORT_LOW, CMP,
      [LAST_PRICE]. Power the "3-Day Rising" / "5-Day Rising" tables on
      the front page. Set either to None to skip.

BEFORE YOU RUN THIS LOCALLY
  pip install -r requirements.txt
  Drop your Excel/CSV files into the "data" folder next to this file
  (create it if it isn't there), using the filenames below - or edit
  the CONFIG block to point at wherever your files actually live.
  Then:
  python app.py
  -> open http://127.0.0.1:5000 in a browser

DEPLOYMENT NOTE (Render / gunicorn)
  This is now a plain synchronous Flask app - no WebSocket, no
  background thread. Any standard gunicorn setup works, e.g.:
      gunicorn -w 2 app:app
  requirements.txt only needs: flask, pandas, openpyxl, gunicorn, requests.

LOGIN
  Admin (password only) sees everything; registered users see only the
  3-Day / 5-Day Rising tables. User accounts live in a Google Sheet behind
  an Apps Script (see UserAuth_AppsScript.gs). Set SECRET_KEY,
  ADMIN_PASSWORD, APPS_SCRIPT_URL and APPS_SCRIPT_SECRET in Render.
"""

import hmac
import os
import re
import time
from datetime import datetime, timedelta
from functools import wraps
from pathlib import Path

import pandas as pd
import requests
from flask import Flask, jsonify, redirect, render_template, request, session
from werkzeug.security import check_password_hash, generate_password_hash

# ======================= CONFIG - EDIT THESE =======================
# By default everything is read from a "data" folder next to this file,
# so anyone you share this project with just drops their own Excel/CSV
# files in there and runs it - no path editing needed. If your files
# live somewhere else instead, replace any of the lines below with a
# full path, e.g. NEWS_FILE = r"C:\NSEDATA\EQdata\Company_News.xlsx"
DATA_DIR            = Path(__file__).resolve().parent / "data"

NEWS_FILE           = DATA_DIR / "Company_News.xlsx"
SYMBOL_MASTER_FILE  = DATA_DIR / "EQUITY_L.csv"
LOW_PRICE_FILE      = DATA_DIR / "Low_Price.xlsx"          # set to None to skip
TREND_3DAY_FILE     = DATA_DIR / "uptrend_output.xlsx"     # set to None to skip
TREND_5DAY_FILE     = DATA_DIR / "uptrend_5day.xlsx"       # set to None to skip
TREND_ALL_FILE      = DATA_DIR / "trend_output.xlsx"       # ALL companies (Up/Down/Sideways) - set to None to skip

# "Positive News" tab: how many calendar days (counted back from the
# newest date in NEWS_FILE) the default view covers
POSITIVE_DAYS       = 5
# ====================================================================

app = Flask(__name__)

# ============================ LOGIN / ROLES ============================
# Set these in Render -> your service -> Environment (never hard-code them):
#   SECRET_KEY          long random text - signs the login cookie
#   ADMIN_PASSWORD      the admin's password
#   APPS_SCRIPT_URL     the Google Apps Script web-app URL (ends in /exec)
#   APPS_SCRIPT_SECRET  same text as API_SECRET in the Apps Script
#
# ADMIN -> everything.   USER -> only the 3-Day / 5-Day Rising tables.
def _env(name):
    """Environment value with stray spaces/newlines removed ('' if not set)."""
    return (os.environ.get(name) or "").strip()


SECRET_KEY         = _env("SECRET_KEY")
ADMIN_PASSWORD     = _env("ADMIN_PASSWORD")
APPS_SCRIPT_URL    = _env("APPS_SCRIPT_URL")
APPS_SCRIPT_SECRET = _env("APPS_SCRIPT_SECRET")

# one line in the Render Logs showing which settings are present (values are never printed)
print("[login config] " + ", ".join(
    f"{n}={'set' if v else 'MISSING'}" for n, v in [
        ("SECRET_KEY", SECRET_KEY), ("ADMIN_PASSWORD", ADMIN_PASSWORD),
        ("APPS_SCRIPT_URL", APPS_SCRIPT_URL), ("APPS_SCRIPT_SECRET", APPS_SCRIPT_SECRET),
    ]))

IS_HOSTED = "PORT" in os.environ
if not SECRET_KEY:
    if IS_HOSTED:
        raise RuntimeError("SECRET_KEY is not set - add it in Render -> Environment.")
    SECRET_KEY = "dev-only-secret-change-me"      # local testing only

app.config.update(
    SECRET_KEY=SECRET_KEY,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=IS_HOSTED,
    PERMANENT_SESSION_LIFETIME=timedelta(hours=12),
)

# used so a wrong name and a wrong password take the same time to reject
_DUMMY_HASH = generate_password_hash("not-a-real-password")

# simple brute-force guard: too many wrong attempts -> wait 10 minutes
_FAILS = {}
MAX_FAILS, FAIL_WINDOW = 6, 600


def _too_many(key):
    now = time.time()
    hits = [t for t in _FAILS.get(key, []) if now - t < FAIL_WINDOW]
    _FAILS[key] = hits
    return len(hits) >= MAX_FAILS


def _record_fail(key):
    _FAILS.setdefault(key, []).append(time.time())


def _clear_fails(key):
    _FAILS.pop(key, None)


def login_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if session.get("role") not in ("admin", "user"):
            if request.path.startswith("/api/"):
                return jsonify({"error": "login required"}), 401
            return redirect("/login")
        return f(*args, **kwargs)
    return wrapper


def admin_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        role = session.get("role")
        if role == "admin":
            return f(*args, **kwargs)
        if role == "user":
            return jsonify({"error": "admin only"}), 403
        return jsonify({"error": "login required"}), 401
    return wrapper


@app.after_request
def _no_store(resp):
    # never let the browser cache pages/data behind the login
    if request.path in ("/", "/login") or request.path.startswith("/api/"):
        resp.headers["Cache-Control"] = "no-store"
    return resp


def call_sheet(payload):
    """Talk to the Google Apps Script that keeps the Users sheet."""
    if not APPS_SCRIPT_URL or not APPS_SCRIPT_SECRET:
        missing = [n for n, v in [("APPS_SCRIPT_URL", APPS_SCRIPT_URL), ("APPS_SCRIPT_SECRET", APPS_SCRIPT_SECRET)] if not v]
        print(f"[warn] user login not set up - missing in Render Environment: {', '.join(missing)}")
        return {"ok": False, "code": "server", "error": "User login is not set up on the server yet."}
    try:
        r = requests.post(APPS_SCRIPT_URL, json={**payload, "secret": APPS_SCRIPT_SECRET}, timeout=25)
        return r.json()
    except Exception as e:
        print(f"[warn] Apps Script call failed: {e}")
        return {"ok": False, "code": "server", "error": "Could not reach the user database. Please try again."}


def clean_phone(raw):
    """Any of 9876543210 / +91 98765 43210 / 09876543210 -> '9876543210', else None."""
    d = re.sub(r"\D", "", str(raw))
    if len(d) == 12 and d.startswith("91"):
        d = d[2:]
    elif len(d) == 11 and d.startswith("0"):
        d = d[1:]
    return d if re.fullmatch(r"\d{10}", d) else None


_GENERIC_WORDS = re.compile(r"\b(limited|ltd|company|co|corporation|corp|the|of|india)\b")


def normalize(name):
    """Loose match key: lowercase, drop punctuation and generic
    corporate words, so 'Aarti Drugs Limited' and 'AARTI DRUGS' land
    on the same key."""
    s = str(name).lower().replace("&", "and")
    s = re.sub(r"[.,()]", "", s)
    s = _GENERIC_WORDS.sub("", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def safe_symbol(raw: str) -> str:
    """Uppercase + strip anything unsafe in a symbol string."""
    return re.sub(r"[^A-Z0-9&\-]", "", str(raw).upper())


# ------------------------- data loaded once at startup -------------------------


def load_symbol_master():
    if not SYMBOL_MASTER_FILE or not os.path.exists(SYMBOL_MASTER_FILE):
        return None
    try:
        df = pd.read_csv(SYMBOL_MASTER_FILE)
        df.columns = [c.strip() for c in df.columns]
        df["SYMBOL"] = df["SYMBOL"].astype(str).str.strip().str.upper()
        df["NAME OF COMPANY"] = df["NAME OF COMPANY"].astype(str).str.strip()
        df["norm"] = df["NAME OF COMPANY"].map(normalize)
        return df
    except Exception as e:
        print(f"[warn] could not load SYMBOL_MASTER_FILE: {e}")
        return None


def load_news_sentiment():
    """Company_News.xlsx -> forward-filled company name, normalized
    sentiment/summary/category. This is the required, base news source
    - Category now comes straight from this file (Penalty / Dividend /
    AGM or EGM / Financial Results / Others), normalized to
    underscore-separated so it's a safe, consistent CSS class name
    (e.g. "AGM or EGM" -> "AGM_or_EGM")."""
    df = pd.read_excel(NEWS_FILE)
    df.columns = [c.strip() for c in df.columns]

    df["Company Name"] = df["Company Name"].ffill()
    df["Company Name"] = df["Company Name"].astype(str).str.strip()
    df["Date_parsed"] = pd.to_datetime(df["Date"], errors="coerce", dayfirst=True)
    df["Sentiment"] = df["Sentiment"].fillna("unknown").astype(str).str.strip().str.lower()
    df["Summary"] = df["Summary"].astype(str).str.strip()

    if "Category" in df.columns:
        df["Category"] = (
            df["Category"].fillna("Others").astype(str).str.strip()
            .str.replace(r"\s+", "_", regex=True)
        )
    else:
        df["Category"] = None

    df = df.dropna(subset=["Date_parsed"])
    df = df[df["Summary"].str.len() > 0]

    # guard against stray typo'd dates in the source file (e.g. a year
    # keyed as 2531 instead of 2026) so they can't sort to the top of
    # "latest news" as if they were the newest thing that happened
    max_valid_date = datetime.now() + timedelta(days=2)
    bad_dates = df[df["Date_parsed"] > max_valid_date]
    if len(bad_dates):
        print(f"[warn] dropping {len(bad_dates)} row(s) in NEWS_FILE with an implausible future date:")
        for _, r in bad_dates.iterrows():
            print(f"        {r['Company Name']!r} -> {r['Date_parsed']}")
        df = df[df["Date_parsed"] <= max_valid_date]

    return df


def load_low_price():
    if not LOW_PRICE_FILE or not os.path.exists(LOW_PRICE_FILE):
        return None
    try:
        df = pd.read_excel(LOW_PRICE_FILE)
        df.columns = [c.strip() for c in df.columns]
        df["SYMBOL"] = df["SYMBOL"].astype(str).str.strip().str.upper()
        return df.set_index("SYMBOL")
    except Exception as e:
        print(f"[warn] could not load LOW_PRICE_FILE: {e}")
        return None


def load_trend_file(path):
    """Read a 3-day or 5-day uptrend output file (SYMBOL, DATE, TREND,
    SUPPORT_HIGH, SUPPORT_LOW, CMP, [LAST_PRICE]). Returns a plain list
    of row dicts, or None if the file isn't configured/found."""
    if not path or not os.path.exists(path):
        return None
    try:
        df = pd.read_excel(path)
        df.columns = [c.strip() for c in df.columns]
        df["SYMBOL"] = df["SYMBOL"].astype(str).str.strip().str.upper()
        df["TREND"] = df["TREND"].astype(str).str.strip()
        rows = []
        for _, row in df.iterrows():
            symbol = row["SYMBOL"]
            rows.append({
                "symbol": symbol,
                "name": SYMBOL_TO_NAME.get(symbol, ""),
                "trend": row.get("TREND"),
                "support_high": row.get("SUPPORT_HIGH"),
                "support_low": row.get("SUPPORT_LOW"),
                "cmp": row.get("CMP"),
                "last_price": row.get("LAST_PRICE"),
            })
        return rows
    except Exception as e:
        print(f"[warn] could not load trend file {path}: {e}")
        return None


# ------------------------- build the combined index -------------------------


def build_index():
    news = load_news_sentiment()
    master = load_symbol_master()

    if master is not None:
        symbol_set = set(master["SYMBOL"])
        norm_to_symbol = dict(zip(master["norm"], master["SYMBOL"]))
        symbol_to_name = dict(zip(master["SYMBOL"], master["NAME OF COMPANY"]))
    else:
        symbol_set, norm_to_symbol, symbol_to_name = set(), {}, {}

    # fold pure-casing duplicates first (e.g. "OM INFRA LIMITED" / "Om Infra Limited")
    counts = news["Company Name"].value_counts()
    fold_map = {}
    for name in counts.index:
        key = name.lower()
        if key not in fold_map:
            fold_map[key] = name  # counts.index sorted by frequency desc
    news["Company_norm"] = news["Company Name"].str.lower().map(fold_map)

    # resolve each folded name to an NSE symbol, if possible
    group_to_symbol = {}
    for g in news["Company_norm"].unique():
        n = normalize(g)
        direct = g.strip().upper()
        if n in norm_to_symbol:
            group_to_symbol[g] = norm_to_symbol[n]
        elif direct in symbol_set:
            group_to_symbol[g] = direct
        else:
            group_to_symbol[g] = None

    news["symbol"] = news["Company_norm"].map(group_to_symbol)
    news["symbol"] = news["symbol"].where(news["symbol"].notna(), None)
    # entity key: the symbol when we found one (merges every spelling
    # variant together), otherwise fall back to the folded name itself
    news["entity_key"] = news["symbol"].fillna(news["Company_norm"])

    def display_name(row):
        if row["symbol"]:
            return symbol_to_name.get(row["symbol"], row["Company_norm"])
        return row["Company_norm"]

    news["display_name"] = news.apply(display_name, axis=1)
    news["date_str"] = news["Date_parsed"].dt.strftime("%Y-%m-%d")

    combined = news[["entity_key", "display_name", "symbol", "date_str", "Summary", "Sentiment", "Category"]].copy()
    combined = combined.rename(columns={"Category": "category"})
    combined = combined.sort_values("date_str", ascending=False)
    return combined, symbol_to_name


NEWS_DF, SYMBOL_TO_NAME = build_index()
LOW_PRICE_DF = load_low_price()
TREND_3DAY = load_trend_file(TREND_3DAY_FILE)
TREND_5DAY = load_trend_file(TREND_5DAY_FILE)
TREND_ALL = load_trend_file(TREND_ALL_FILE)

# symbol -> {trend, support_high, support_low, cmp} for the small
# Up/Down/Sideways badge shown next to a searched company's price -
# this is a lookup, not a browsable list
TREND_BY_SYMBOL = {row["symbol"]: row for row in (TREND_ALL or [])}

# one row per entity: entity_key, display_name, symbol, article count
ENTITIES = (
    NEWS_DF[["entity_key", "display_name", "symbol"]]
    .drop_duplicates("entity_key")
    .assign(count=NEWS_DF.groupby("entity_key")["entity_key"].transform("count"))
    .drop_duplicates("entity_key")
    .sort_values("display_name", key=lambda s: s.str.lower())
    .to_dict("records")
)
for _e in ENTITIES:
    if pd.isna(_e["symbol"]):
        _e["symbol"] = None

ENTITY_BY_KEY = {e["entity_key"]: e for e in ENTITIES}


# ------------------------------- helpers -------------------------------


def get_news_for_entity(key: str):
    return NEWS_DF[NEWS_DF["entity_key"] == key].sort_values("date_str", ascending=False)


def resolve_query_to_entity(query: str):
    """Exact match on a symbol or a display name -> entity_key, else None."""
    q = str(query).strip()
    q_upper = q.upper()
    for e in ENTITIES:
        if e["symbol"] == q_upper or e["display_name"].strip().upper() == q_upper:
            return e["entity_key"]
    return None


def get_price_quote(symbol):
    """Static CMP/5-day-low/20-day-low lookup straight from
    Low_Price.xlsx - no live/streaming price fetch."""
    if not symbol or LOW_PRICE_DF is None or symbol not in LOW_PRICE_DF.index:
        return None
    row = LOW_PRICE_DF.loc[symbol]
    return {
        "cmp": row.get("LAST_PRICE"),
        "low_5day": row.get("5Day_Low"),
        "low_20day": row.get("20Day_Low"),
    }


def get_trend_badge(symbol):
    """Small Up/Down/Sideways badge info for a searched company, from
    trend_output.xlsx - just a lookup, not the full company list."""
    if not symbol:
        return None
    row = TREND_BY_SYMBOL.get(symbol)
    if not row:
        return None
    raw = str(row.get("trend") or "").strip()
    letter = raw[:1].upper()
    direction = {"U": "up", "D": "down"}.get(letter, "side")
    label = re.sub(r"^[A-Z]\s*", "", raw).strip()  # e.g. "▲ UpTrend"
    return {"direction": direction, "label": label}


# --------------------------------- routes ---------------------------------


@app.route("/login")
def login_page():
    if session.get("role") in ("admin", "user"):
        return redirect("/")
    return render_template("login.html")


@app.route("/api/register", methods=["POST"])
def api_register():
    d = request.get_json(silent=True) or {}
    name = re.sub(r"\s+", " ", str(d.get("name", "")).strip())
    phone = clean_phone(d.get("phone", ""))
    pw = str(d.get("password", ""))
    confirm = str(d.get("confirm", ""))

    if not (2 <= len(name) <= 40) or not re.search(r"[^\W\d_]", name):
        return jsonify({"ok": False, "error": "Enter your name (2-40 characters, with letters)."}), 400
    if not phone:
        return jsonify({"ok": False, "error": "Enter a valid 10-digit phone number."}), 400
    if not (6 <= len(pw) <= 64):
        return jsonify({"ok": False, "error": "Password must be 6 to 64 characters."}), 400
    if pw != confirm:
        return jsonify({"ok": False, "error": "Passwords do not match."}), 400

    # only a salted hash ever leaves this server - never the password itself
    res = call_sheet({"action": "add", "name": name, "phone": phone, "hash": generate_password_hash(pw)})
    if res.get("ok"):
        return jsonify({"ok": True, "phone": phone})
    code = res.get("code")
    if code == "phone_exists":
        return jsonify({"ok": False, "error": "This phone number is already registered. Please log in."}), 409
    if code == "name_exists":
        return jsonify({"ok": False, "error": "This name is already taken. Please use a different name."}), 409
    return jsonify({"ok": False, "error": res.get("error", "Registration failed.")}), 503


@app.route("/api/login", methods=["POST"])
def api_login():
    d = request.get_json(silent=True) or {}
    mode = d.get("mode")
    pw = str(d.get("password", ""))

    # ---- admin: password only ----
    if mode == "admin":
        key = "admin"
        if _too_many(key):
            return jsonify({"ok": False, "error": "Too many wrong attempts. Try again in 10 minutes."}), 429
        if ADMIN_PASSWORD and hmac.compare_digest(pw.encode(), ADMIN_PASSWORD.encode()):
            _clear_fails(key)
            session.clear()
            session["role"], session["name"] = "admin", "Admin"
            session.permanent = True
            return jsonify({"ok": True})
        _record_fail(key)
        return jsonify({"ok": False, "error": "Wrong admin password."}), 401

    # ---- user: name OR phone + password ----
    ident = re.sub(r"\s+", " ", str(d.get("identifier", "")).strip())
    if not ident or not pw:
        return jsonify({"ok": False, "error": "Enter your name or phone number and password."}), 400

    looks_like_phone = bool(re.fullmatch(r"[\d\s+\-]+", ident))
    phone = clean_phone(ident) if looks_like_phone else None
    if looks_like_phone and not phone:
        return jsonify({"ok": False, "error": "Enter a valid 10-digit phone number, or your name."}), 400

    key = "user|" + (phone or ident.lower())
    if _too_many(key):
        return jsonify({"ok": False, "error": "Too many wrong attempts. Try again in 10 minutes."}), 429

    res = call_sheet({"action": "find", "kind": "phone" if phone else "name", "identifier": phone or ident})
    if not res.get("ok"):
        return jsonify({"ok": False, "error": res.get("error", "Login is unavailable right now.")}), 503

    user = res.get("user")
    good = check_password_hash(user["hash"] if user else _DUMMY_HASH, pw)
    if not user or not good:
        _record_fail(key)
        return jsonify({"ok": False, "error": "Wrong name/phone or password."}), 401
    if str(user.get("status", "")).strip().lower() != "active":
        return jsonify({"ok": False, "error": "Your account is not active. Please contact the admin."}), 403

    _clear_fails(key)
    session.clear()
    session["role"], session["name"] = "user", user.get("name", "User")
    session.permanent = True
    return jsonify({"ok": True})


@app.route("/logout", methods=["POST"])
def logout():
    session.clear()
    return jsonify({"ok": True})


@app.route("/")
@login_required
def index():
    return render_template("index.html", role=session["role"], name=session.get("name", ""))


@app.route("/api/companies")
@admin_required
def api_companies():
    """Company/symbol matches for the search box. Matches against both
    the display name and the NSE symbol."""
    q = request.args.get("q", "").strip().lower()
    matches = ENTITIES
    if q:
        matches = [
            e for e in ENTITIES
            if q in e["display_name"].lower() or (e["symbol"] and q in e["symbol"].lower())
        ]
        matches.sort(key=lambda e: (
            0 if (e["symbol"] and e["symbol"].lower().startswith(q)) else 1,
            e["display_name"].lower(),
        ))
    matches = matches[:30]
    return jsonify([
        {"name": e["display_name"], "symbol": e["symbol"], "count": int(e["count"])}
        for e in matches
    ])


@app.route("/api/news/<path:query>")
@admin_required
def api_news(query):
    key = resolve_query_to_entity(query)
    if key is None:
        return jsonify({"company": query, "symbol": None, "news": [], "price": None, "price_error": None})

    rows = get_news_for_entity(key)
    entity = ENTITY_BY_KEY[key]
    symbol = entity["symbol"]

    news_records = [
        {
            "date": row["date_str"],
            "summary": row["Summary"],
            "sentiment": row["Sentiment"],
            "category": None if pd.isna(row["category"]) else row["category"],
        }
        for _, row in rows.iterrows()
    ]

    quote = get_price_quote(symbol)
    trend = get_trend_badge(symbol)

    return jsonify({
        "company": entity["display_name"],
        "symbol": symbol,
        "news": news_records,
        "quote": quote,
        "trend": trend,
    })


@app.route("/api/stats")
@admin_required
def api_stats():
    return jsonify({
        "companies": len(ENTITIES),
        "articles": len(NEWS_DF),
    })


@app.route("/api/trend/3day")
@login_required
def api_trend_3day():
    if TREND_3DAY is None:
        return jsonify({"rows": [], "available": False})
    return jsonify({"rows": TREND_3DAY, "available": True})


@app.route("/api/trend/5day")
@login_required
def api_trend_5day():
    if TREND_5DAY is None:
        return jsonify({"rows": [], "available": False})
    return jsonify({"rows": TREND_5DAY, "available": True})


@app.route("/api/latest-news")
@admin_required
def api_latest_news():
    """Every article published on the most recent date present in
    NEWS_FILE, plus the date immediately before it - not a fixed
    top-N count. If you add a new day's news and restart the app,
    this automatically follows the new latest two dates; it never
    stays stuck showing older days' items."""
    if NEWS_DF.empty:
        return jsonify([])
    latest_two_dates = sorted(NEWS_DF["date_str"].unique(), reverse=True)[:2]
    rows = NEWS_DF[NEWS_DF["date_str"].isin(latest_two_dates)].sort_values("date_str", ascending=False)
    records = [
        {
            "company": row["display_name"],
            "symbol": None if pd.isna(row["symbol"]) else row["symbol"],
            "date": row["date_str"],
            "summary": row["Summary"],
            "sentiment": row["Sentiment"],
            "category": None if pd.isna(row["category"]) else row["category"],
        }
        for _, row in rows.iterrows()
    ]
    return jsonify(records)


@app.route("/api/positive-news")
@admin_required
def api_positive_news():
    """Positive-sentiment news from Company_News.xlsx only.
    Default: the latest POSITIVE_DAYS calendar days (counted back from
    the newest date in the file). With ?all=1: every positive item."""
    if NEWS_DF.empty:
        return jsonify([])

    pos = NEWS_DF[NEWS_DF["Sentiment"] == "positive"]

    if request.args.get("all") != "1":
        latest = pd.to_datetime(NEWS_DF["date_str"]).max()
        cutoff = (latest - timedelta(days=POSITIVE_DAYS - 1)).strftime("%Y-%m-%d")
        pos = pos[pos["date_str"] >= cutoff]

    pos = pos.sort_values("date_str", ascending=False)
    return jsonify([
        {
            "company": row["display_name"],
            "symbol": None if pd.isna(row["symbol"]) else row["symbol"],
            "date": row["date_str"],
            "summary": row["Summary"],
            "sentiment": row["Sentiment"],
            "category": None if pd.isna(row["category"]) else row["category"],
        }
        for _, row in pos.iterrows()
    ])


if __name__ == "__main__":
    # Render (and most hosts) set PORT for you and expect the app to
    # bind 0.0.0.0. debug=True is a security risk on a public server
    # (it exposes a live Python console on error pages), so it's off
    # whenever PORT is set - i.e. whenever this is actually deployed.
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=not IS_HOSTED)
