"""
fetch_data.py  (tech DCF project)

Pulls raw "company facts" JSON from the SEC EDGAR XBRL API for 10 SaaS/cloud
software companies and caches each response locally.

Unlike the retail SQL project, a DCF needs more than the income statement --
it needs enough to build a full bridge from revenue down to a per-share
equity value: operating cash flow and capex (to derive free cash flow),
cash and debt (to bridge enterprise value to equity value), and shares
outstanding (to get to a per-share number). fetch_data.py itself doesn't
change which concepts get pulled -- company facts returns everything a
company has ever tagged -- but clean_data.py will need to know which tags
to look for, so this file's COMPANIES list and conventions carry over
directly.

SEC fair-access requirements: descriptive User-Agent + client-side rate
limiting, same as the retail project.

Usage:
    python fetch_data.py
    python fetch_data.py --force
    python fetch_data.py --ticker SNOW
"""

import argparse
import json
import sys
import time
from pathlib import Path

import requests

# --------------------------------------------------------------------------
# Config
# --------------------------------------------------------------------------

USER_AGENT = "Aidan Falcon, student project, aidanfalcon0617@gmail.com"

REQUESTS_PER_SECOND = 5
SLEEP_BETWEEN_REQUESTS = 1.0 / REQUESTS_PER_SECOND

BASE_URL = "https://data.sec.gov/api/xbrl/companyfacts/{cik}.json"

RAW_DATA_DIR = Path(__file__).resolve().parent.parent / "data" / "raw"

# Company universe: SaaS/cloud software, 10 companies, deliberately spanning
# mature/profitable (Adobe, Salesforce) to still-scaling/thin-margin
# (Snowflake, Datadog, CrowdStrike) -- the DCF assumptions have to flex
# across that range, which is the point.
COMPANIES = [
    {"ticker": "CRM",  "company_name": "Salesforce, Inc.",           "cik": "0001108524"},
    {"ticker": "ADBE", "company_name": "Adobe Inc.",                 "cik": "0000796343"},
    {"ticker": "NOW",  "company_name": "ServiceNow, Inc.",           "cik": "0001373715"},
    {"ticker": "WDAY", "company_name": "Workday, Inc.",              "cik": "0001327811"},
    {"ticker": "HUBS", "company_name": "HubSpot Inc",                "cik": "0001404655"},
    {"ticker": "TEAM", "company_name": "Atlassian Corp",             "cik": "0001650372"},
    {"ticker": "DDOG", "company_name": "Datadog, Inc.",              "cik": "0001561550"},
    {"ticker": "SNOW", "company_name": "Snowflake Inc.",             "cik": "0001640147"},
    {"ticker": "ZS",   "company_name": "Zscaler, Inc.",              "cik": "0001713683"},
    {"ticker": "CRWD", "company_name": "CrowdStrike Holdings, Inc.", "cik": "0001535527"},
]


def fetch_company_facts(cik: str, session: requests.Session) -> dict:
    url = BASE_URL.format(cik=f"CIK{cik}")
    response = session.get(url, timeout=30)
    response.raise_for_status()
    return response.json()


def cache_path_for(ticker: str) -> Path:
    return RAW_DATA_DIR / f"{ticker}_companyfacts.json"


def main():
    parser = argparse.ArgumentParser(description="Fetch and cache SEC EDGAR company facts.")
    parser.add_argument("--force", action="store_true", help="Re-download even if a cached file exists.")
    parser.add_argument("--ticker", type=str, default=None, help="Fetch only this ticker (e.g. SNOW).")
    args = parser.parse_args()

    RAW_DATA_DIR.mkdir(parents=True, exist_ok=True)

    companies = COMPANIES
    if args.ticker:
        companies = [c for c in COMPANIES if c["ticker"].upper() == args.ticker.upper()]
        if not companies:
            print(f"Ticker '{args.ticker}' not found in COMPANIES list.", file=sys.stderr)
            sys.exit(1)

    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT})

    fetched, skipped, failed = 0, 0, 0

    for company in companies:
        ticker = company["ticker"]
        cik = company["cik"]
        out_path = cache_path_for(ticker)

        if out_path.exists() and not args.force:
            print(f"[skip]  {ticker}: cached at {out_path}")
            skipped += 1
            continue

        print(f"[fetch] {ticker} (CIK {cik}) ...", end=" ", flush=True)
        try:
            data = fetch_company_facts(cik, session)
        except requests.exceptions.RequestException as e:
            print(f"FAILED ({e})")
            failed += 1
            continue

        with open(out_path, "w") as f:
            json.dump(data, f)

        print(f"saved ({out_path.stat().st_size / 1024:.0f} KB)")
        fetched += 1

        time.sleep(SLEEP_BETWEEN_REQUESTS)

    print(f"\nDone. fetched={fetched} skipped={skipped} failed={failed}")
    if failed:
        sys.exit(1)


if __name__ == "__main__":
    main()
