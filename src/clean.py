"""Clean the raw Shopify + ad-spend exports and write a data-quality log.

Sample project - synthetic data, not client work.

Inputs : data/raw/*.csv
Outputs: data/clean/order_lines_clean.csv
         data/clean/ad_spend_clean.csv
         data/clean/data_quality_log.csv   (also copied into the workbook)

Every rule below appends one row to the log: step, rule, rows affected, orders affected,
USD value affected, action taken. Nothing is dropped silently.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw"
CLEAN = ROOT / "data" / "clean"
STORE_TZ = "America/New_York"
TEST_EMAIL_PATTERN = r"@northwind-test\."

log: list[dict] = []


def record(step: str, rule: str, rows: int, orders: int | None, value_usd: float | None, action: str) -> None:
    dataset = "ad spend" if step.startswith("S") else "orders"
    log.append(dict(step=len(log) + 1, dataset=dataset, rule=rule, rows_affected=int(rows),
                    orders_affected=None if orders is None else int(orders),
                    value_usd=None if value_usd is None else round(float(value_usd), 2),
                    action=action))


def parse_timestamp(s: pd.Series) -> tuple[pd.Series, dict]:
    """Parse mixed timestamp formats to UTC.

    - '...Z'            -> UTC
    - '...+hh:mm/-hh:mm' -> explicit offset, converted to UTC
    - naive 'YYYY-MM-DD HH:MM:SS' (legacy export) -> assumed store-local (America/New_York)
    """
    s = s.str.strip()
    is_z = s.str.endswith("Z")
    has_off = s.str.match(r".*[+-]\d{2}:\d{2}$")
    naive = ~(is_z | has_off)
    out = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns, UTC]")
    out[is_z | has_off] = pd.to_datetime(s[is_z | has_off], utc=True, format="ISO8601")
    local = pd.to_datetime(s[naive], format="%Y-%m-%d %H:%M:%S")
    out[naive] = local.dt.tz_localize(STORE_TZ, ambiguous="NaT", nonexistent="shift_forward").dt.tz_convert("UTC")
    counts = {"utc_z": int(is_z.sum()), "explicit_offset": int(has_off.sum()), "naive_store_local": int(naive.sum())}
    return out, counts


def main() -> None:
    CLEAN.mkdir(parents=True, exist_ok=True)
    raw = pd.read_csv(RAW / "shopify_orders_export.csv", dtype=str, keep_default_na=False)
    products = pd.read_csv(RAW / "products.csv")
    fx = pd.read_csv(RAW / "fx_rates_monthly.csv", dtype={"month": str})
    record("0", "Raw order-line rows loaded", len(raw), raw["order_id"].nunique(), None, "input")

    df = raw.copy()
    # ---- 1. normalise text fields
    cur_bad = (df["currency"] != df["currency"].str.strip().str.upper()).sum()
    df["currency"] = df["currency"].str.strip().str.upper()
    record("1a", "Currency code had stray whitespace / lowercase", cur_bad, None, None, "fixed (trim + upper-case)")
    em_bad = ((df["customer_email"] != df["customer_email"].str.strip().str.lower()) & (df["customer_email"] != "")).sum()
    df["customer_email"] = df["customer_email"].str.strip().str.lower()
    record("1b", "Email had leading spaces / upper-case", em_bad, None, None, "fixed (trim + lower-case)")

    for c in ["order_id", "line_number", "quantity"]:
        df[c] = df[c].astype(int)
    for c in ["unit_price", "line_discount", "refunded_amount"]:
        df[c] = df[c].astype(float)

    # ---- 7. timestamps
    df["created_at_utc"], counts = parse_timestamp(df["created_at"])
    record("7a", "Timestamps in UTC 'Z' format", counts["utc_z"], None, None, "parsed")
    record("7b", "Timestamps with explicit UTC offset (e.g. -04:00)", counts["explicit_offset"], None, None,
           "converted to UTC")
    record("7c", "Naive timestamps (legacy export, no zone)", counts["naive_store_local"], None, None,
           "assumed store-local America/New_York, converted to UTC")
    bad_ts = df["created_at_utc"].isna()
    record("7d", "Unparseable timestamps", bad_ts.sum(), None, None, "removed")
    df = df[~bad_ts].copy()
    local = df["created_at_utc"].dt.tz_convert(STORE_TZ)
    df["order_date"] = local.dt.strftime("%Y-%m-%d")
    df["month"] = local.dt.strftime("%Y-%m")
    # an order must have one timestamp; check consistency
    n_ts = df.groupby("order_id")["created_at_utc"].nunique()
    record("7e", "Orders whose lines disagree on timestamp", (n_ts > 1).sum(), (n_ts > 1).sum(), None,
           "check only (0 expected)")

    fx_map = dict(zip(fx["month"], fx["cad_to_usd"]))
    df["fx"] = df["month"].map(fx_map).where(df["currency"] == "CAD", 1.0)
    df["gross_usd"] = df["quantity"].abs() * df["unit_price"] * df["fx"]

    # ---- 2. duplicated rows (export pagination overlap)
    dup_mask = df.duplicated(subset=["order_id", "line_number"], keep="first")
    full_dup = df.duplicated(keep="first")
    assert (dup_mask == full_dup).all(), "duplicate keys with differing content - needs manual review"
    record("2", "Duplicated order lines (same order_id + line_number, identical content)", dup_mask.sum(),
           df.loc[dup_mask, "order_id"].nunique(), df.loc[dup_mask, "gross_usd"].sum(),
           "removed (value = gross sales it would have double-counted)")
    df = df[~dup_mask].copy()

    # ---- 3. test orders
    test = df["customer_email"].str.contains(TEST_EMAIL_PATTERN, regex=True)
    record("3", "Staff test orders (email @northwind-test.*)", test.sum(), df.loc[test, "order_id"].nunique(),
           df.loc[test, "gross_usd"].sum(), "removed")
    df = df[~test].copy()

    # ---- 4. voided / cancelled orders
    void = df["financial_status"].isin(["voided", "cancelled"])
    record("4", "Voided orders (payment never captured)", void.sum(), df.loc[void, "order_id"].nunique(),
           df.loc[void, "gross_usd"].sum(), "removed")
    df = df[~void].copy()

    # ---- 5. negative quantities
    neg = df["quantity"] <= 0
    record("5", "Non-positive quantity on a sale line (manual adjustment keyed into wrong field)", neg.sum(),
           df.loc[neg, "order_id"].nunique(), df.loc[neg, "gross_usd"].sum(),
           "quarantined: line excluded from all metrics (cannot tell if return or typo); see data/clean/quarantine.csv")
    df.loc[neg].to_csv(CLEAN / "quarantine.csv", index=False)
    df = df[~neg].copy()

    # ---- 6. missing emails: recover from sibling lines on the same order
    blank = df["customer_email"] == ""
    order_email = df[~blank].groupby("order_id")["customer_email"].first()
    recovered = blank & df["order_id"].isin(order_email.index)
    df.loc[recovered, "customer_email"] = df.loc[recovered, "order_id"].map(order_email)
    record("6a", "Missing email recovered from another line of the same order", recovered.sum(),
           df.loc[recovered, "order_id"].nunique(), None, "fixed")
    still = df["customer_email"] == ""
    record("6b", "Missing email, not recoverable (guest checkout / export glitch)", still.sum(),
           df.loc[still, "order_id"].nunique(), None,
           "kept in revenue & order counts; excluded from customer metrics (repeat rate, cohorts, CAC)")

    # ---- 8. currency conversion
    unknown_cur = ~df["currency"].isin(["USD", "CAD"])
    record("8a", "Unknown currency code", unknown_cur.sum(), None, None, "removed")
    df = df[~unknown_cur].copy()
    rate = df["fx"]
    assert rate.notna().all(), "missing FX rate for a month"
    cad = df["currency"] == "CAD"
    record("8b", "CAD lines converted to USD at monthly finance rate", cad.sum(), df.loc[cad, "order_id"].nunique(),
           (df.loc[cad, "quantity"] * df.loc[cad, "unit_price"] * rate[cad]).sum(), "converted (value = USD gross sales)")

    # ---- 9. refunds sanity
    over = df["refunded_amount"] > (df["quantity"] * df["unit_price"] - df["line_discount"]) + 0.01
    record("9", "Refund larger than line net value", over.sum(), None, None, "check only (0 expected)")

    df["unit_price_usd"] = (df["unit_price"] * rate).round(4)
    df["gross_usd"] = (df["quantity"] * df["unit_price"] * rate).round(4)
    df["discount_usd"] = (df["line_discount"] * rate).round(4)
    df["refund_usd"] = (df["refunded_amount"] * rate).round(4)
    df["net_usd"] = df["gross_usd"] - df["discount_usd"] - df["refund_usd"]
    prod = products.set_index("sku")
    unknown_sku = ~df["sku"].isin(prod.index)
    record("10", "SKU not in product catalog", unknown_sku.sum(), None, None, "check only (0 expected)")
    df["category"] = df["sku"].map(prod["category"])
    df["cogs_usd"] = df["quantity"] * df["sku"].map(prod["unit_cost_usd"])

    out_cols = ["order_id", "line_number", "created_at_utc", "order_date", "month", "customer_email",
                "billing_country", "currency", "sku", "product_title", "category", "quantity", "unit_price_usd",
                "gross_usd", "discount_usd", "refund_usd", "net_usd", "cogs_usd", "discount_code",
                "financial_status", "source_channel"]
    df = df.sort_values(["created_at_utc", "order_id", "line_number"])
    df[out_cols].to_csv(CLEAN / "order_lines_clean.csv", index=False)
    record("11", "Clean order lines written", len(df), df["order_id"].nunique(), df["net_usd"].sum(),
           "output (value = net revenue USD)")

    # ---- ad spend
    spend = pd.read_csv(RAW / "ad_spend_daily.csv")
    record("S0", "Raw ad-spend rows loaded", len(spend), None, spend["spend_usd"].sum(), "input")
    sd = spend.duplicated(subset=["date", "channel", "campaign"])
    record("S1", "Duplicated spend rows (date + campaign)", sd.sum(), None, spend.loc[sd, "spend_usd"].sum(), "removed")
    spend = spend[~sd].copy()
    neg_sp = spend["spend_usd"] < 0
    record("S2", "Negative spend", neg_sp.sum(), None, None, "removed")
    spend = spend[~neg_sp].copy()
    spend["month"] = spend["date"].str[:7]
    spend.to_csv(CLEAN / "ad_spend_clean.csv", index=False)
    record("S3", "Clean ad-spend rows written", len(spend), None, spend["spend_usd"].sum(), "output")

    pd.DataFrame(log).to_csv(CLEAN / "data_quality_log.csv", index=False)
    print(pd.DataFrame(log).to_string(index=False))


if __name__ == "__main__":
    main()
