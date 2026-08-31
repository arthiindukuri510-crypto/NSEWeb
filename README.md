# NSE Company Terminal

Search any NSE-listed company by name or symbol and see its latest CMP,
5-day/20-day low, and a merged news feed with both a Category tag
(Penalty / Dividend / AGM or EGM / Financial Results / Others) and a
Sentiment tag per article. The front page also shows 3-Day Rising /
5-Day Rising stock tables and the latest day's news across all
companies.

## Setup (first time only)

1. Install Python 3.9+ if you don't already have it.
2. Clone this repo, then from inside the folder:
   ```
   pip install -r requirements.txt
   ```

## Add your data

Create a `data` folder next to `app.py` (if it isn't there already)
and put your files in it, using these exact names:

| File in `data/`          | Required? | What it's for                                  |
|---------------------------|-----------|-------------------------------------------------|
| `Company_News.xlsx`       | **Yes**   | Company Name, Date, Summary, Sentiment, Category |
| `EQUITY_L.csv`             | Recommended | NSE's symbol master - lets you search by symbol too |
| `Low_Price.xlsx`           | Optional  | Powers the CMP / 5-day low / 20-day low numbers  |
| `uptrend_output.xlsx`      | Optional  | Powers the "3-Day Rising" table on the front page|
| `uptrend_5day.xlsx`        | Optional  | Powers the "5-Day Rising" table on the front page|

If a file is missing, that one feature just doesn't show up - the app
still runs fine without it.

If you'd rather keep your data somewhere else entirely (not in a
`data` folder next to the code), open `app.py` and edit the paths in
the `CONFIG` block near the top to point wherever your files live.

These 5 files are the ones this repo actually commits (see the
Hosting section below for why) - `.gitignore` blocks any other
Excel/CSV so nothing bulkier gets added by accident.


## Run it locally

```
python app.py
```

Then open **http://127.0.0.1:5000** in a browser.

## Hosting it as a real public website (one link, no setup for anyone else)

GitHub itself can only serve static pages, not a running Python app -
so "push to GitHub" alone doesn't give you a live link. The practical
way to get one is a host that deploys straight from your GitHub repo:

**Render (free tier)** - recommended, simplest for this app:
1. Push this repo to GitHub (data files included - see below).
2. Go to [render.com](https://render.com) → sign in with GitHub →
   **New +** → **Web Service** → pick this repo.
3. Set:
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `python app.py`
4. Deploy. Render gives you a URL like
   `https://your-app-name.onrender.com` - that's the link you share.
   Every time you push to GitHub, Render redeploys automatically.

Free-tier note: the app spins down after ~15 minutes of no traffic,
so the first visit after a quiet spell takes 20-30 seconds to wake up
- normal for a free instance, not a bug.

### Data files go in the repo for this to work

Since the app now runs on Render's server, not your PC, it needs its
own copy of the data - so unlike the "clone and run locally" setup,
these 5 files ARE meant to be committed (the `.gitignore` already
allows exactly these five, nothing else):
`data/Company_News.xlsx`, `data/EQUITY_L.csv`, `data/Low_Price.xlsx`,
`data/uptrend_output.xlsx`, `data/uptrend_5day.xlsx`.

Anyone with the link sees whatever data is currently committed - to
update it, replace the files in `data/`, commit, and push; Render
redeploys with the new data automatically. There's no way for a
visitor to search data you haven't pushed.


## Updating with new news day-to-day

`update_company_news.py` (not part of the web app itself, run
separately whenever you have new data) appends a day's raw news file
into `Company_News.xlsx`, classifying Category and trimming each
company back to its most recent 15 dates. Safe to re-run even if a
file accidentally contains news you already added - duplicates are
skipped automatically.

- **Running locally**: restart `app.py` after updating any data file
  - it's loaded once at startup, not re-read live.
- **Running on Render**: commit and push the updated file in `data/`
  - Render redeploys automatically and picks up the change. Restarting
    your own machine's `app.py` does nothing for the hosted version.

