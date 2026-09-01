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

BEFORE YOU RUN THIS
  pip install -r requirements.txt
  Drop your Excel/CSV files into the "data" folder next to this file
  (create it if it isn't there), using the filenames below - or edit
  the CONFIG block to point at wherever your files actually live.
  Then:
  python app.py
  -> open http://127.0.0.1:5000 in a browser
"""

import os
import re
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
from flask import Flask, jsonify, render_template, request

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
# ====================================================================

app = Flask(__name__)


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
    return {
        "cmp": row.get("LAST_PRICE"),
        "low_5day": row.get("5Day_Low"),
        "low_20day": row.get("20Day_Low"),
    }


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

    return jsonify({
        "company": entity["display_name"],
        "symbol": symbol,
        "news": news_records,
        "quote": quote,
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


if __name__ == "__main__":
    # Render (and most hosts) set PORT for you and expect the app to
    # bind 0.0.0.0. debug=True is a security risk on a public server
    # (it exposes a live Python console on error pages), so it's off
    # whenever PORT is set - i.e. whenever this is actually deployed.
    port = int(os.environ.get("PORT", 5000))
    is_hosted = "PORT" in os.environ
    app.run(host="0.0.0.0", port=port, debug=not is_hosted)
