"""
clean_data.py  (tech DCF project)

Parses cached SEC EDGAR "company facts" JSON (produced by fetch_data.py) into
a clean, per-company annual financial history with everything the DCF model
needs to go from revenue down to a per-share equity value.

Unlike the retail SQL project's clean_data.py, this only needs *annual* (10-K)
figures -- a DCF projects yearly free cash flow, it doesn't need quarterly
granularity -- so there's no Q4-derivation step and no cross-company
fiscal-year normalization (each company's DCF is built off its own reported
fiscal years independently; nothing here compares company A's FY to company
B's FY the way the retail project's ranking queries did).

What carries over from the retail project's approach:

  1. Tag fallback: some companies report a concept under a different XBRL tag
     than others. For each field we try a list of tags, in priority order,
     and merge period-by-period -- a lower-priority tag only fills in dates
     the higher-priority tag never covers, so we don't silently truncate a
     company's history at whatever year it switched tags, but also don't mix
     two different tagged concepts for the same period.

  2. Restatements: if the same period_end_date appears more than once (a
     prior-year figure repeated as a comparative in a later 10-K), we keep
     the value from the most recently *filed* occurrence and log it if the
     value actually changed.

  3. Missing concepts: if a company has no data at all for a field, it's left
     NULL for every year for that company -- nothing is fabricated. This
     matters for debt specifically: several of these companies (Snowflake,
     Datadog, Zscaler, HubSpot) funded themselves through convertible notes
     rather than conventional term debt, so they report under tags like
     ConvertibleDebtNoncurrent instead of LongTermDebtNoncurrent -- real debt
     that a narrower tag list would have missed entirely and left at NULL.

Fields pulled, and why the DCF needs them:
  - revenue, operating_income, net_income        -- income statement
  - income_tax_expense, pretax_income             -- to derive effective tax
                                                      rate for unlevered FCF
  - interest_expense                              -- for a levered/unlevered
                                                      FCF reconciliation
  - depreciation_amortization, capex,
    operating_cash_flow                           -- to derive free cash flow
  - cash_and_equivalents, short_term_debt,
    long_term_debt                                -- to bridge enterprise
                                                      value to equity value
  - diluted_shares_outstanding                    -- to get to a per-share
                                                      number

Two purely-arithmetic fields are derived here (not fabricated, just summed/
subtracted from reported figures), the same way the retail project derives
gross_profit: free_cash_flow (operating_cash_flow - capex), total_debt
(short_term_debt + long_term_debt), and effective_tax_rate (income_tax_expense
/ pretax_income). Everything else -- revenue growth fade, WACC, terminal
value -- is the DCF model script's job, not this one's.

Outputs (all under data/processed/):
  - companies.csv
  - dcf_financials.csv     -- one row per company per fiscal year
  - filings_metadata.csv
  - data_quality_log.csv   -- human-readable audit trail of every fallback,
                               restatement, and missing-concept decision

Usage:
    python clean_data.py
"""

import csv
import json
from collections import defaultdict
from datetime import date
from pathlib import Path

# --------------------------------------------------------------------------
# Config
# --------------------------------------------------------------------------

PROJECT_ROOT = Path(__file__).resolve().parent.parent
RAW_DATA_DIR = PROJECT_ROOT / "data" / "raw"
PROCESSED_DATA_DIR = PROJECT_ROOT / "data" / "processed"

# Must match the COMPANIES list in fetch_data.py.
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
SECTOR = "SaaS / Cloud Software"

# field_name -> ordered list of XBRL tags to try, first match wins per period.
# Debt tags in particular vary a lot across this group: several of these
# companies (SNOW, DDOG, ZS, HUBS) funded themselves with convertible notes
# instead of conventional term debt, so "LongTermDebtNoncurrent" alone would
# leave them all NULL even though they do carry debt.
CONCEPT_TAGS = {
    "revenue": [
        "Revenues",
        "RevenueFromContractWithCustomerExcludingAssessedTax",
        "RevenueFromContractWithCustomerIncludingAssessedTax",
        "SalesRevenueNet",
    ],
    "operating_income": [
        "OperatingIncomeLoss",
    ],
    "net_income": [
        "NetIncomeLoss",
    ],
    "income_tax_expense": [
        "IncomeTaxExpenseBenefit",
    ],
    "pretax_income": [
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments",
    ],
    "interest_expense": [
        "InterestExpense",
        "InterestExpenseDebt",
    ],
    "depreciation_amortization": [
        "DepreciationDepletionAndAmortization",
        "DepreciationAmortizationAndAccretionNet",
        "DepreciationAndAmortization",
    ],
    "capex": [
        "PaymentsToAcquirePropertyPlantAndEquipment",
        "PaymentsForCapitalImprovements",
        "PaymentsToAcquireProductiveAssets",
    ],
    "operating_cash_flow": [
        "NetCashProvidedByUsedInOperatingActivities",
        "NetCashProvidedByUsedInOperatingActivitiesContinuingOperations",
    ],
    "cash_and_equivalents": [
        "CashAndCashEquivalentsAtCarryingValue",
        "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents",
    ],
    "short_term_debt": [
        "LongTermDebtCurrent",
        "ConvertibleDebtCurrent",
        "ConvertibleNotesPayableCurrent",
        "DebtCurrent",
        "ShortTermBorrowings",
    ],
    "long_term_debt": [
        "LongTermDebtNoncurrent",
        "ConvertibleDebtNoncurrent",
        "ConvertibleLongTermNotesPayable",
        "ConvertibleDebt",
        "LongTermDebt",
    ],
    "diluted_shares_outstanding": [
        "WeightedAverageNumberOfDilutedSharesOutstanding",
        "WeightedAverageNumberOfSharesOutstandingBasic",
    ],
}

# Every concept above is reported in USD except share counts, which XBRL
# tags under a "shares" unit instead.
FIELD_UNITS = {"diluted_shares_outstanding": "shares"}

# Point-in-time (instant) balance-sheet concepts have no 'start' date and so
# are exempt from the flow-field duration sanity check below.
POINT_FIELDS = {"cash_and_equivalents", "short_term_debt", "long_term_debt"}
FLOW_FIELDS = set(CONCEPT_TAGS) - POINT_FIELDS

ANNUAL_FORMS = {"10-K", "10-K/A"}
ANNUAL_DURATION_DAYS = (330, 400)

# Accumulates human-readable audit rows across the whole run.
QUALITY_LOG = []


def log_event(ticker, field, issue_type, period_end_date, details):
    QUALITY_LOG.append(
        {
            "ticker": ticker,
            "field": field,
            "issue_type": issue_type,
            "period_end_date": period_end_date or "",
            "details": details,
        }
    )


# --------------------------------------------------------------------------
# Loading raw data
# --------------------------------------------------------------------------

def load_companyfacts(ticker: str) -> dict | None:
    path = RAW_DATA_DIR / f"{ticker}_companyfacts.json"
    if not path.exists():
        return None
    with open(path) as f:
        return json.load(f)


def get_tag_facts(cf_json: dict, tag: str, unit: str = "USD") -> list:
    """Return the list of fact entries for a given us-gaap tag under the
    given unit ("USD" for money, "shares" for share counts), or []."""
    try:
        return cf_json["facts"]["us-gaap"][tag]["units"][unit]
    except KeyError:
        return []


def collect_tag_facts(cf_json: dict, tags: list, ticker: str, field: str) -> list:
    """Gather facts for a field across its whole fallback tag list, filling
    gaps in the highest-priority tag from lower-priority ones on a
    per-period basis (see module docstring, point 1)."""
    unit = FIELD_UNITS.get(field, "USD")
    per_tag_facts = {tag: get_tag_facts(cf_json, tag, unit) for tag in tags}
    used_tags = [tag for tag in tags if per_tag_facts[tag]]
    if not used_tags:
        log_event(ticker, field, "missing_concept", None,
                   f"No data found under any of {tags}; field left NULL for all years.")
        return []

    covered_ends = set()
    merged = []
    fallback_tags_used = []
    for tag in tags:
        facts = per_tag_facts[tag]
        if not facts:
            continue
        new_facts = [f for f in facts if f.get("end") not in covered_ends]
        if new_facts and tag != tags[0]:
            fallback_tags_used.append(tag)
        merged.extend(new_facts)
        covered_ends |= {f.get("end") for f in facts}

    if fallback_tags_used:
        log_event(
            ticker, field, "tag_fallback", None,
            f"Primary tag '{tags[0]}' didn't cover every year; filled gaps "
            f"from {fallback_tags_used} (period-by-period, only for dates "
            f"the higher-priority tag never reports).",
        )
    return merged


def fact_duration_days(fact: dict) -> int | None:
    start, end = fact.get("start"), fact.get("end")
    if not start or not end:
        return None
    try:
        return (date.fromisoformat(end) - date.fromisoformat(start)).days
    except ValueError:
        return None


def is_annual_fact(fact: dict, is_point_in_time: bool = False) -> bool:
    """True if this fact is a real fiscal-year figure from a 10-K.

    XBRL filings occasionally carry stray facts (footnote tables, divested-
    segment disclosures) mistagged with fp='FY' despite spanning a much
    shorter period. Sanity-checking the duration -- not just form/fp -- keeps
    those from silently overwriting the real annual value. Point-in-time
    balance-sheet facts are instants with no 'start' date, so they're exempt.
    """
    if fact.get("form") not in ANNUAL_FORMS or fact.get("fp") != "FY":
        return False
    if is_point_in_time:
        return True
    duration = fact_duration_days(fact)
    return duration is not None and ANNUAL_DURATION_DAYS[0] <= duration <= ANNUAL_DURATION_DAYS[1]


def dedupe_by_end_date(facts: list, ticker: str, field: str) -> dict:
    """Collapse facts to one-per-period_end_date, keeping the value from the
    most recently *filed* occurrence so a true restatement is picked up.

    Deliberately keyed on period_end_date -- the one thing every occurrence
    of a given real period agrees on -- rather than each fact's self-reported
    'fy' label. That label isn't reliable: a tag's very first XBRL appearance
    often bundles several historical years of comparatives into one filing
    with every comparative period carrying *that filing's own*
    DocumentFiscalYearFocus, so the same raw fy value can legitimately point
    at two different real periods. period_end_date has no such ambiguity."""
    groups = defaultdict(list)
    for fact in facts:
        end = fact.get("end")
        val = fact.get("val")
        if end is None or val is None or not fact.get("filed"):
            continue
        groups[end].append(fact)

    by_end = {}
    for end, group in groups.items():
        group.sort(key=lambda f: f["filed"])
        earliest, latest = group[0], group[-1]

        distinct_vals = {f["val"] for f in group}
        if len(distinct_vals) > 1:
            log_event(
                ticker, field, "restatement", end,
                f"Value changed across filings: {earliest['val']} "
                f"(filed {earliest['filed']}) -> {latest['val']} "
                f"(filed {latest['filed']}). Keeping the later filing's value.",
            )

        by_end[end] = latest
    return by_end


# --------------------------------------------------------------------------
# Building the dcf_financials records
# --------------------------------------------------------------------------

def blank_record(cik, period_end_date):
    return {
        "cik": cik,
        "fiscal_year": None,
        "period_end_date": period_end_date,
        "revenue": None,
        "operating_income": None,
        "net_income": None,
        "income_tax_expense": None,
        "pretax_income": None,
        "effective_tax_rate": None,
        "interest_expense": None,
        "depreciation_amortization": None,
        "capex": None,
        "operating_cash_flow": None,
        "free_cash_flow": None,
        "cash_and_equivalents": None,
        "short_term_debt": None,
        "long_term_debt": None,
        "total_debt": None,
        "diluted_shares_outstanding": None,
        "filing_date": None,
    }


def build_records_for_company(cf_json: dict, ticker: str, cik: str) -> dict:
    """Returns dict keyed by period_end_date -> record dict.

    period_end_date, not the self-reported 'fy' label, is the ground truth
    for which real fiscal year a fact belongs to (see dedupe_by_end_date's
    docstring) -- so it's what ties every field's facts together here too.
    fiscal_year is filled in afterward, derived from period_end_date.
    """
    records = {}

    def get_or_create(end):
        if end not in records:
            records[end] = blank_record(cik, end)
        return records[end]

    for field, tags in CONCEPT_TAGS.items():
        facts = collect_tag_facts(cf_json, tags, ticker, field)
        if not facts:
            continue

        is_point_in_time = field in POINT_FIELDS
        annual_facts = [f for f in facts if is_annual_fact(f, is_point_in_time)]
        annual_by_end = dedupe_by_end_date(annual_facts, ticker, field)

        for end, fact in annual_by_end.items():
            rec = get_or_create(end)
            rec[field] = fact.get("val")
            filed = fact.get("filed")
            if filed and (rec["filing_date"] is None or filed > rec["filing_date"]):
                rec["filing_date"] = filed

    for rec in records.values():
        rec["fiscal_year"] = date.fromisoformat(rec["period_end_date"]).year

    compute_derived_fields(records, ticker)
    return records


def compute_derived_fields(records: dict, ticker: str):
    """Purely arithmetic combinations of already-reported figures -- nothing
    fabricated, just summed/subtracted (same spirit as the retail project's
    gross_profit derivation)."""
    for rec in records.values():
        ocf, capex = rec["operating_cash_flow"], rec["capex"]
        rec["free_cash_flow"] = (ocf - capex) if (ocf is not None and capex is not None) else None

        st, lt = rec["short_term_debt"], rec["long_term_debt"]
        if st is None and lt is None:
            rec["total_debt"] = None
        else:
            rec["total_debt"] = (st or 0) + (lt or 0)
            if st is None or lt is None:
                log_event(
                    ticker, "total_debt", "partial_debt_components", rec.get("period_end_date"),
                    f"Only one of short_term_debt ({st}) / long_term_debt ({lt}) was "
                    f"reported for this year; total_debt treats the missing one as 0 "
                    f"rather than leaving the total NULL.",
                )

        tax, pretax = rec["income_tax_expense"], rec["pretax_income"]
        if tax is not None and pretax:
            rec["effective_tax_rate"] = tax / pretax


# --------------------------------------------------------------------------
# Filings metadata
# --------------------------------------------------------------------------

def extract_filings_metadata(cf_json: dict, cik: str) -> list:
    """Distinct (form, filing_date, accession_number) tuples appearing
    anywhere in this company's facts -- lets someone trace any number in
    dcf_financials.csv back to the original 10-K on EDGAR."""
    seen = set()
    rows = []
    us_gaap = cf_json.get("facts", {}).get("us-gaap", {})
    for tag_data in us_gaap.values():
        for fact in tag_data.get("units", {}).get("USD", []):
            form = fact.get("form")
            filed = fact.get("filed")
            accn = fact.get("accn")
            key = (form, filed, accn)
            if form and filed and accn and key not in seen:
                seen.add(key)
                rows.append({
                    "cik": cik,
                    "form_type": form,
                    "filing_date": filed,
                    "accession_number": accn,
                })
    return rows


# --------------------------------------------------------------------------
# CSV writers
# --------------------------------------------------------------------------

def write_companies_csv(path: Path):
    with open(path, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["cik", "ticker", "company_name", "sector"])
        writer.writeheader()
        for c in COMPANIES:
            writer.writerow({
                "cik": c["cik"],
                "ticker": c["ticker"],
                "company_name": c["company_name"],
                "sector": SECTOR,
            })


def write_dcf_financials_csv(path: Path, all_records: list):
    fieldnames = [
        "id", "cik", "fiscal_year", "period_end_date",
        "revenue", "operating_income", "net_income",
        "income_tax_expense", "pretax_income", "effective_tax_rate", "interest_expense",
        "depreciation_amortization", "capex", "operating_cash_flow", "free_cash_flow",
        "cash_and_equivalents", "short_term_debt", "long_term_debt", "total_debt",
        "diluted_shares_outstanding", "filing_date",
    ]
    with open(path, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for i, rec in enumerate(all_records, start=1):
            writer.writerow({"id": i, **rec})


def write_filings_metadata_csv(path: Path, all_filings: list):
    fieldnames = ["id", "cik", "form_type", "filing_date", "accession_number"]
    with open(path, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for i, row in enumerate(all_filings, start=1):
            writer.writerow({"id": i, **row})


def write_quality_log_csv(path: Path):
    fieldnames = ["ticker", "field", "issue_type", "period_end_date", "details"]
    with open(path, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in QUALITY_LOG:
            writer.writerow(row)


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def main():
    PROCESSED_DATA_DIR.mkdir(parents=True, exist_ok=True)

    all_records = []
    all_filings = []
    missing_tickers = []

    for company in COMPANIES:
        ticker = company["ticker"]
        cik = company["cik"]
        cf_json = load_companyfacts(ticker)
        if cf_json is None:
            missing_tickers.append(ticker)
            print(f"[skip]  {ticker}: no cached raw JSON found (run fetch_data.py first)")
            continue

        print(f"[clean] {ticker} ...", end=" ")
        records = build_records_for_company(cf_json, ticker, cik)
        n = len(records)
        all_records.extend(records.values())
        all_filings.extend(extract_filings_metadata(cf_json, cik))
        print(f"{n} fiscal years")

    all_records.sort(key=lambda r: (r["cik"], r["fiscal_year"]))

    write_companies_csv(PROCESSED_DATA_DIR / "companies.csv")
    write_dcf_financials_csv(PROCESSED_DATA_DIR / "dcf_financials.csv", all_records)
    write_filings_metadata_csv(PROCESSED_DATA_DIR / "filings_metadata.csv", all_filings)
    write_quality_log_csv(PROCESSED_DATA_DIR / "data_quality_log.csv")

    print(f"\nDone. {len(all_records)} dcf_financials rows, "
          f"{len(all_filings)} filings_metadata rows, "
          f"{len(QUALITY_LOG)} data quality log entries.")
    if missing_tickers:
        print(f"Missing raw data for: {', '.join(missing_tickers)} (run fetch_data.py)")


if __name__ == "__main__":
    main()
