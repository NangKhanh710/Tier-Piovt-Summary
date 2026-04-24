"""
inject_data.py
--------------
Queries SQL Server, aggregates real sales data, and injects it into
index.html, producing index_live.html — a fully standalone dashboard.

Usage:
    export DB_SERVER=<host>
    export DB_NAME=<database>
    export DB_USER=<username>
    export DB_PASS=<password>
    python inject_data.py
"""

import json
import os
import re
import sys

import pandas as pd
import pyodbc


# ── Credentials ──────────────────────────────────────────────────────────────

def get_connection():
    server   = os.environ.get("DB_SERVER", "")
    database = os.environ.get("DB_NAME",   "")
    username = os.environ.get("DB_USER",   "")
    password = os.environ.get("DB_PASS",   "")

    missing = [k for k, v in {"DB_SERVER": server, "DB_NAME": database,
                               "DB_USER": username, "DB_PASS": password}.items() if not v]
    if missing:
        sys.exit(f"Error: missing environment variable(s): {', '.join(missing)}")

    conn_str = (
        "DRIVER={ODBC Driver 17 for SQL Server};"
        f"SERVER={server};DATABASE={database};UID={username};PWD={password};"
    )
    return pyodbc.connect(conn_str)


# ── SQL queries (same logic as database.py, no Streamlit dependency) ─────────

def load_data(conn) -> pd.DataFrame:
    print("Querying customer data…")
    df_customer = pd.read_sql("""
        SELECT T0.CustomerID, T0.CustomerName, T0.Address, T0.RegionID,
               T0.CityID, T0.AccountID, T1.AccountName, T1.AccountDisplay,
               CASE
                   WHEN T0.ChannelName = '21-Gov. Hosp - Dept'    THEN '101'
                   WHEN T0.ChannelName = '22-Gov. Hosp - Phar'    THEN '102'
                   WHEN T0.ChannelName = '23-Private Hosp - Dept' THEN '104'
                   WHEN T0.ChannelName = '24-Private Hosp - Phar' THEN '106'
                   WHEN T0.ChannelName = '25-Polyclinic'          THEN '103'
                   WHEN T0.ChannelName = '26-Pharmacy'            THEN '105'
                   WHEN T0.ChannelName = '27-Wholesalers'         THEN '107'
                   WHEN T0.ChannelName = '28-Monoclinic'          THEN '108'
                   WHEN T0.ChannelName = '29-Pharmacy Chain'      THEN '109'
               END AS ChannelID
        FROM [VNDB_O365].[dbo].[vw_Sipro_Customer] AS T0
        LEFT JOIN [VNDB_O365].[dbo].[vw_ParentAccount] T1
            ON T0.CUSTOMERID = T1.CUSTOMERID
    """, conn)

    print("Querying sales data…")
    df_sales = pd.read_sql("""
        SELECT T0.CustomerID, T0.Qty, T1.TerritoryID, T0.Cust_Brand,
               T0.BrandID, T0.ProductID, T0.ByMonth, T0.ByYear,
               T0.YYMM, T0.AmountUSD, T1.ProductWeight
        FROM (
            SELECT CustomerID, AmountUSD, Qty, BrandID,
                   ProductID, ByMonth, ByYear,
                   YYMM, CONCAT(CustomerID, Brand) AS Cust_Brand
            FROM [VNDB_O365].[dbo].[vw_SIpro_CustomerProductByCutoff]
        ) AS T0
        LEFT JOIN (
            SELECT TeamID, TerritoryID, CustomerID, BrandID,
                   ProductWeight, Brand,
                   CONCAT(CustomerID, Brand) AS Cust_Brand
            FROM [VNDB_O365].[dbo].[vw_SIpro_TerritoryAlignment]
        ) AS T1 ON T0.Cust_Brand = T1.Cust_Brand
    """, conn)

    print(f"  customers: {len(df_customer):,} rows  |  sales: {len(df_sales):,} rows")

    df = df_sales.merge(df_customer, on="CustomerID", how="left")
    print(f"  merged:    {len(df):,} rows")
    return df


# ── Aggregation ───────────────────────────────────────────────────────────────

def aggregate(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # Drop rows with no account name (unresolved customers)
    df = df[df["AccountName"].notna() & (df["AccountName"] != "")]

    # Weighted revenue
    df["ProductWeight"] = pd.to_numeric(df["ProductWeight"], errors="coerce").fillna(1.0)
    df["AmountUSD"]     = pd.to_numeric(df["AmountUSD"],     errors="coerce").fillna(0.0)
    df["WeightedAmount"] = df["AmountUSD"] * df["ProductWeight"]

    # Primary RegionID / ChannelID / TerritoryID per account (most-common value)
    def mode_first(s):
        m = s.mode()
        return str(m.iloc[0]) if len(m) else ""

    meta = (
        df.groupby("AccountName", sort=False)
        .agg(
            RegionID   =("RegionID",    mode_first),
            ChannelID  =("ChannelID",   mode_first),
            TerritoryID=("TerritoryID", mode_first),
        )
        .reset_index()
    )

    # Revenue summed to account × year × month
    agg = (
        df.groupby(["AccountName", "ByYear", "ByMonth"], as_index=False)["WeightedAmount"]
        .sum()
    )
    agg = agg.merge(meta, on="AccountName", how="left")

    # Conform to HTML row format: ProductWeight=1 so AmountUSD*ProductWeight = WeightedAmount
    agg["AmountUSD"]    = agg["WeightedAmount"]
    agg["ProductWeight"] = 1.0
    agg["ByYear"]  = agg["ByYear"].astype(int)
    agg["ByMonth"] = agg["ByMonth"].astype(int)

    print(f"  aggregated: {len(agg):,} account-month rows  |  {agg['AccountName'].nunique():,} unique accounts")
    return agg[["AccountName", "ByYear", "ByMonth", "AmountUSD", "ProductWeight",
                "RegionID", "ChannelID", "TerritoryID"]]


# ── Build derived constants ───────────────────────────────────────────────────

def build_constants(agg: pd.DataFrame):
    years   = sorted(agg["ByYear"].unique().tolist(), reverse=True)
    regions = sorted(agg["RegionID"].dropna().unique().tolist(), key=str)
    terrs   = sorted(agg["TerritoryID"].dropna().unique().tolist(), key=str)

    channel_ids = sorted(agg["ChannelID"].dropna().unique().tolist(), key=str)
    ch_map = {
        "101": "Gov. Hosp – Dept",
        "102": "Gov. Hosp – Phar",
        "103": "Polyclinic",
        "104": "Pvt. Hosp – Dept",
        "105": "Pharmacy",
        "106": "Pvt. Hosp – Phar",
        "107": "Wholesalers",
        "108": "Monoclinic",
        "109": "Pharmacy Chain",
    }
    ch_list = [
        {"id": cid, "n": ch_map.get(cid, cid)}
        for cid in channel_ids if cid
    ]

    return years, regions, terrs, ch_list


# ── HTML patching ─────────────────────────────────────────────────────────────

SAMPLE_DATA_START = "// ─ Sample data ────────────────────────────────────────────────────────────"
SAMPLE_DATA_END   = "const ALL_DATA=(()=>{"

def patch_html(src: str, rows: list, years: list, regions: list,
               terrs: list, ch_list: list) -> str:

    # Serialize rows as compact JSON
    data_json = json.dumps(rows, ensure_ascii=False, separators=(",", ":"))

    replacement = (
        "// ─ SQL Data (generated by inject_data.py) "
        "───────────────────────\n"
        f"const ALL_DATA={data_json};"
    )

    # Replace the entire sample-data block (start comment → end of ALL_DATA line)
    pattern = re.compile(
        re.escape(SAMPLE_DATA_START) + r".*?" + r"const ALL_DATA=\(\(\)=>\{.*?\}\)\(\);",
        re.DOTALL,
    )
    patched, n = pattern.subn(replacement, src)
    if n == 0:
        sys.exit("Error: could not locate the sample-data block in index.html. "
                 "Has the file been modified?")

    # Update YEARS constant
    years_js = json.dumps(years)
    patched = re.sub(r"const YEARS\s*=\s*\[.*?\];", f"const YEARS={years_js};", patched)

    # Update REGIONS constant
    regions_js = json.dumps([str(r) for r in regions])
    patched = re.sub(r'const REGIONS\s*=\s*\[.*?\];', f"const REGIONS={regions_js};", patched)

    # Update TERRS constant
    terrs_js = json.dumps([str(t) for t in terrs])
    patched = re.sub(r'const TERRS\s*=\s*\[.*?\];', f"const TERRS={terrs_js};", patched)

    # Update CH_LIST constant
    ch_js = json.dumps(ch_list, ensure_ascii=False, separators=(",", ":"))
    patched = re.sub(r'const CH_LIST\s*=\s*\[.*?\];', f"const CH_LIST={ch_js};", patched)

    return patched


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    src_path = os.path.join(os.path.dirname(__file__), "index.html")
    out_path = os.path.join(os.path.dirname(__file__), "index_live.html")

    print("Connecting to SQL Server…")
    conn = get_connection()

    df      = load_data(conn)
    agg     = aggregate(df)
    years, regions, terrs, ch_list = build_constants(agg)

    print(f"Years:     {years}")
    print(f"Regions:   {regions}")
    print(f"Channels:  {[c['id'] for c in ch_list]}")
    print(f"Territories: {terrs}")

    print(f"Reading {src_path}…")
    with open(src_path, encoding="utf-8") as f:
        src = f.read()

    # Convert aggregated DataFrame to list of dicts (JSON-serialisable types)
    rows = agg.to_dict(orient="records")
    for r in rows:
        r["ByYear"]       = int(r["ByYear"])
        r["ByMonth"]      = int(r["ByMonth"])
        r["AmountUSD"]    = round(float(r["AmountUSD"]), 4)
        r["ProductWeight"] = 1.0
        r["RegionID"]     = str(r["RegionID"])
        r["ChannelID"]    = str(r["ChannelID"])
        r["TerritoryID"]  = str(r["TerritoryID"])

    print("Patching HTML…")
    patched = patch_html(src, rows, years, regions, terrs, ch_list)

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(patched)

    size_mb = os.path.getsize(out_path) / 1_048_576
    print(f"Written to {out_path}  ({size_mb:.1f} MB)")
    print("Done. Open index_live.html in a browser to verify.")


if __name__ == "__main__":
    main()
