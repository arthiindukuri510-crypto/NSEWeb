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

  TREND_3DAY_FILE / TREND_5DAY_FILE (optional) - uptrend_output.xlsx /
      uptrend_5day.xlsx (output of your trend_uptrend_only.py /
      trend_uptrend_5day.py scripts)
      columns: SYMBOL, DATE, TREND, SUPPORT_HIGH, SUPPORT_LOW, CMP,
      [LAST_PRICE]. Power the "3-Day Rising" / "5-Day Rising" tables on
      the front page. Set either to None to skip.

LIVE CMP
  A background thread polls Yahoo Finance's free quote endpoint every
  POLL_INTERVAL_SECONDS - but ONLY for symbols actually visible right
  now: the 3-day/5-day rising tables, plus whichever companies someone
  has looked up recently (see WATCHED_SYMBOL_TTL_SECONDS below). It
  pushes price updates to every open browser tab over a WebSocket - no
  page refresh needed.

  A company that's just been searched but isn't in either rising table
  gets its price fetched on demand, right in that request, instead of
  waiting for the next background sweep - see get_price_quote().

  This is Yahoo's free public quote data for NSE tickers, which runs
  about 15 minutes behind the live exchange feed - not true tick-by-
  tick, but good enough to "just keep updating" a client demo without
  needing a paid broker API. If you ever want true live tick data,
  swap fetch_live_quotes() for a broker API call (Zerodha Kite
  Connect, Upstox, etc.) - everything downstream (the socket push,
  the frontend flash-on-change) stays the same.

DEPLOYMENT NOTE (Render / gunicorn)
  This app uses Flask-SocketIO for the live-price WebSocket push, which
  needs an async-friendly worker. A few things matter for that to work
  once you're behind gunicorn instead of `python app.py`:

    1. `gevent.monkey.patch_all()` must run before anything else imports
       (threading, requests, etc.) - it's the very first thing this
       file does, below - so those libraries become green-thread aware
       under gevent. (This app originally targeted eventlet, but
       gunicorn 26+ dropped the eventlet worker entirely - eventlet
       itself is now deprecated upstream too - so this uses gevent +
       gevent-websocket instead, which gunicorn still ships proper
       WebSocket support for.)
    2. The background poller is started at *import time* (module level,
       guarded by `_background_started`), not inside
       `if __name__ == "__main__":` - gunicorn imports this module and
       calls the `app`/`socketio` objects directly, it never executes
       that `__main__` block, so anything that used to live only there
       (like starting the poll thread) would silently never run.

  Start command on Render:
      gunicorn -k geventwebsocket.gunicorn.workers.GeventWebSocketWorker -w 1 app:app
  (Use exactly 1 worker - Flask-SocketIO + gevent needs a message queue
  like Redis to coordinate WebSocket state across more than one worker
  process; -w 1 sidesteps that entirely for a demo-scale app.)

  requirements.txt needs: flask, flask-socketio, gevent, gevent-websocket,
  pandas, openpyxl, requests, gunicorn - see requirements.txt alongside
  this file.

BEFORE YOU RUN THIS LOCALLY
  pip install -r requirements.txt
  Drop your Excel/CSV files into the "data" folder next to this file
  (create it if it isn't there), using the filenames below - or edit
  the CONFIG block to point at wherever your files actually live.
  Then:
  python app.py
  -> open http://127.0.0.1:5000 in a browser
"""

# gevent's monkey-patch MUST happen before anything else imports
# threading/socket/requests etc, or Flask-SocketIO's gevent worker
# ends up mixing real OS threads with green threads and you get
# hangs/deadlocks that are miserable to debug. This has to be the
# first executable line in the whole module.
import gevent.monkey
gevent.monkey.patch_all()

import os
import re
import threading
import time
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import requests
from flask import Flask, jsonify, render_template, request
from flask_socketio import SocketIO

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

# --- live CMP polling ---
LIVE_CMP_ENABLED     = True   # set False to fall back to the old static Low_Price.xlsx CMP only
POLL_INTERVAL_SECONDS = 300   # how often to refresh prices (Yahoo's NSE data itself only moves ~every 15 min, so
                               # there's little point going below ~60s - lower this only if you want the "live"
                               # feel more than the underlying number actually changing that often)
YAHOO_CHUNK_SIZE      = 150   # symbols per Yahoo request - keep to two-ish hundred to avoid Yahoo rejecting the call
WATCHED_SYMBOL_TTL_SECONDS = 900   # a searched company stays in the background poll for 15 min after being viewed
ON_DEMAND_FETCH_TIMEOUT = 5    # seconds - keep short so an unresponsive Yahoo call never stalls a page load
# ====================================================================

app = Flask(__name__)
# async_mode left unset so flask-socketio auto-picks the best available:
# eventlet if it's installed (needed for gunicorn + WebSockets on Render -
# see the __main__ block below), otherwise falls back to plain threading
# for local `python app.py` runs.
socketio = SocketIO(app, cors_allowed_origins="*", async_mode="gevent")


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
    if not symbol or LOW_PRICE_DF is None or symbol not in LOW_PRICE_DF.index:
        return None
    row = LOW_PRICE_DF.loc[symbol]

    now = time.time()
    with _LIVE_CMP_LOCK:
        RECENTLY_VIEWED[symbol] = now  # keeps this symbol in the background poll for a while
        live = LIVE_CMP.get(symbol)

    if live is None:
        # nobody's polled this one yet in this process - fetch it right
        # now instead of making the page wait for the next background
        # sweep. Short timeout, so a slow Yahoo response never stalls
        # the page for more than ON_DEMAND_FETCH_TIMEOUT seconds.
        live = fetch_live_quote_single(symbol)
        if live is not None:
            with _LIVE_CMP_LOCK:
                LIVE_CMP[symbol] = live

    return {
        # prefer the live price; fall back to the static Low_Price.xlsx
        # snapshot if Yahoo has nothing for this symbol (just started,
        # delisted, illiquid, or the on-demand fetch above failed/timed out)
        "cmp": live if live is not None else row.get("LAST_PRICE"),
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


# --------------------------------- live CMP polling ---------------------------------

LIVE_CMP = {}          # symbol -> latest price we've fetched from Yahoo
RECENTLY_VIEWED = {}   # symbol -> epoch time it was last looked up (keeps it in the background poll for a while)
_LIVE_CMP_LOCK = threading.Lock()

YAHOO_QUOTE_URL = "https://query1.finance.yahoo.com/v7/finance/quote"
_YAHOO_HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; NSECompanyTerminal/1.0)"}


def build_symbol_universe():
    """Only the symbols that actually need a *background-refreshed*
    live price right now: the 3-day/5-day rising tables (shown to
    every visitor, several rows at once, worth a periodic bulk
    refresh), plus whichever companies someone has looked up in the
    last WATCHED_SYMBOL_TTL_SECONDS (so a company someone has open
    keeps updating live for a while, without needing every one of the
    ~2700 NSE companies polled just in case someone looks at it).

    This used to pull in the ENTIRE company universe (EQUITY_L.csv, or
    later Low_Price.xlsx + trend_output.xlsx - both of which also cover
    ~2700-2765 companies) every single cycle. That meant ~19 chunked
    Yahoo requests every 5 minutes for companies nobody was even
    looking at, which was competing for CPU/network with real visitor
    requests on a free-tier host and is what made the site feel slow/
    unresponsive. A specific company you search for now gets its price
    fetched on demand instead (see get_price_quote()), so nothing has
    to wait for a background sweep to reach it."""
    symbols = set()
    if TREND_3DAY:
        symbols.update(row["symbol"] for row in TREND_3DAY if row.get("symbol"))
    if TREND_5DAY:
        symbols.update(row["symbol"] for row in TREND_5DAY if row.get("symbol"))
    now = time.time()
    with _LIVE_CMP_LOCK:
        symbols.update(
            sym for sym, last_seen in RECENTLY_VIEWED.items()
            if now - last_seen < WATCHED_SYMBOL_TTL_SECONDS
        )
    return sorted(s for s in symbols if s)


def fetch_live_quotes(symbols):
    """Batch-fetch current prices for a list of NSE symbols from
    Yahoo Finance's free quote endpoint (symbol.NS). Returns
    {symbol: price} for whichever symbols Yahoo actually returned a
    price for - missing/unknown symbols are just left out."""
    quotes = {}
    for i in range(0, len(symbols), YAHOO_CHUNK_SIZE):
        chunk = symbols[i:i + YAHOO_CHUNK_SIZE]
        yahoo_symbols = [s + ".NS" for s in chunk]
        try:
            resp = requests.get(
                YAHOO_QUOTE_URL,
                params={"symbols": ",".join(yahoo_symbols)},
                headers=_YAHOO_HEADERS,
                timeout=15,
            )
            resp.raise_for_status()
            data = resp.json()
            for item in data.get("quoteResponse", {}).get("result", []):
                sym = str(item.get("symbol", "")).replace(".NS", "")
                price = item.get("regularMarketPrice")
                if sym and price is not None:
                    quotes[sym] = price
        except Exception as e:
            print(f"[warn] live CMP fetch failed for a batch of {len(chunk)} symbol(s): {e}")
        time.sleep(0.3)  # be polite between batches
    return quotes


def fetch_live_quote_single(symbol, timeout=ON_DEMAND_FETCH_TIMEOUT):
    """One-symbol version of fetch_live_quotes(), used for an on-demand
    lookup when someone searches a company that isn't already in
    LIVE_CMP from the background sweep. Kept to a short timeout so a
    slow/unreachable Yahoo call can never stall a page load for long -
    on any failure we just return None and the caller falls back to
    the static Low_Price.xlsx value."""
    try:
        resp = requests.get(
            YAHOO_QUOTE_URL,
            params={"symbols": symbol + ".NS"},
            headers=_YAHOO_HEADERS,
            timeout=timeout,
        )
        resp.raise_for_status()
        data = resp.json()
        results = data.get("quoteResponse", {}).get("result", [])
        if results:
            price = results[0].get("regularMarketPrice")
            if price is not None:
                return price
    except Exception as e:
        print(f"[warn] on-demand live CMP fetch failed for {symbol}: {e}")
    return None


def poll_live_cmp_forever():
    while True:
        universe = build_symbol_universe()
        if universe:
            new_quotes = fetch_live_quotes(universe)
            if new_quotes:
                with _LIVE_CMP_LOCK:
                    LIVE_CMP.update(new_quotes)
                socketio.emit("price_update", {"prices": new_quotes})
                print(f"[live-cmp] pushed {len(new_quotes)} updated price(s) for {len(universe)} watched symbol(s)")
        time.sleep(POLL_INTERVAL_SECONDS)


# --------------------------------- routes ---------------------------------


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/companies")
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
def api_stats():
    return jsonify({
        "companies": len(ENTITIES),
        "articles": len(NEWS_DF),
    })


@app.route("/api/trend/3day")
def api_trend_3day():
    if TREND_3DAY is None:
        return jsonify({"rows": [], "available": False})
    return jsonify({"rows": TREND_3DAY, "available": True})


@app.route("/api/trend/5day")
def api_trend_5day():
    if TREND_5DAY is None:
        return jsonify({"rows": [], "available": False})
    return jsonify({"rows": TREND_5DAY, "available": True})


@app.route("/api/latest-news")
def api_latest_news():
    """Every article published on the most recent date present in
    NEWS_FILE - not a fixed top-N count. If you add a new day's news
    and restart the app, this automatically follows the new latest
    date; it never stays stuck showing an older day's items."""
    if NEWS_DF.empty:
        return jsonify([])
    latest_date = NEWS_DF["date_str"].max()
    rows = NEWS_DF[NEWS_DF["date_str"] == latest_date].sort_values("date_str", ascending=False)
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


# --------------------------- start the background poller ---------------------------
# This runs at *import time*, not inside `if __name__ == "__main__":`, so it
# fires whether the app is launched with `python app.py` (dev) or with
# `gunicorn -k geventwebsocket.gunicorn.workers.GeventWebSocketWorker -w 1 app:app`
# (Render/production) - gunicorn imports this module and never executes the
# __main__ block below, so anything the live-price feature needs has to be
# started here instead.
_background_started = False
if LIVE_CMP_ENABLED and not _background_started:
    socketio.start_background_task(poll_live_cmp_forever)
    _background_started = True


if __name__ == "__main__":
    # Render (and most hosts) set PORT for you and expect the app to
    # bind 0.0.0.0. debug=True is a security risk on a public server
    # (it exposes a live Python console on error pages), so it's off
    # whenever PORT is set - i.e. whenever this is actually deployed.
    # socketio.run replaces app.run so the WebSocket server starts too.
    port = int(os.environ.get("PORT", 5000))
    is_hosted = "PORT" in os.environ
    socketio.run(app, host="0.0.0.0", port=port, debug=not is_hosted, allow_unsafe_werkzeug=True)
