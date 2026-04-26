#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Convert an Excel template to event.json schema, supporting showTimes (open/start),
and normalizing open/start to 'HH:MM' (no seconds).

Usage:
    python excel_to_event_json.py input.xlsx output.json
Requirements:
    pip install pandas openpyxl
"""
import sys, json, re
import pandas as pd
from datetime import datetime, time

def _format_hhmm(val):
    """
    Normalize various time representations to 'HH:MM' (24h).
    Accepts:
      - datetime.time / datetime.datetime
      - Excel serial time as float/int (fraction of day)
      - strings like '14:00', '14:00:00', '2:3', '02:03:59'
    Returns '' if empty/invalid.
    """
    import math
    from datetime import datetime, time, timedelta

    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    # datetime.time
    if isinstance(val, time):
        return val.strftime("%H:%M")
    # datetime.datetime
    if isinstance(val, datetime):
        return val.strftime("%H:%M")
    # numeric: Excel serial time or seconds
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        try:
            # If within a few days range, treat as Excel day fraction; else seconds
            if -2 <= float(val) <= 3:
                total_seconds = int(round(float(val) * 24 * 3600))
            else:
                total_seconds = int(round(float(val)))
            total_seconds = total_seconds % (24*3600)
            hh = total_seconds // 3600
            mm = (total_seconds % 3600) // 60
            return f"{hh:02d}:{mm:02d}"
        except Exception:
            return ""
    # string forms
    s = str(val).strip()
    if s == "":
        return ""
    m = re.match(r'^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$', s)
    if m:
        h = int(m.group(1)) % 24
        mi = int(m.group(2)) % 60
        return f"{h:02d}:{mi:02d}"
    for fmt in ("%H:%M", "%H:%M:%S"):
        try:
            dt = datetime.strptime(s, fmt)
            return dt.strftime("%H:%M")
        except Exception:
            pass
    return ""

def read_event_sheet(df):
    if not set(["Field","Value"]).issubset(df.columns):
        raise ValueError("Event sheet requires columns: Field, Value")
    out = {}
    for _, r in df.iterrows():
        k = str(r["Field"]).strip()
        v = r["Value"]
        if pd.isna(v): v = ""
        out[k] = v
    return out

def read_shows_sheet(df):
    base_required = ["order","dateISO","venue_name","venue_prefecture","buyUrl","buyUrl2"]
    if not set(base_required).issubset(df.columns):
        raise ValueError(f"Shows sheet requires columns: {base_required}")
    has_open = "open" in df.columns
    has_start = "start" in df.columns

    cols_to_fill = base_required + (["open"] if has_open else []) + (["start"] if has_start else [])
    df2 = df.copy()
    for c in cols_to_fill:
        df2[c] = df2[c].fillna("")

    try:
        df2["order_num"] = pd.to_numeric(df2["order"], errors="coerce")
        df2 = df2.sort_values(by=["order_num"], ascending=[True]).drop(columns=["order_num"])
    except Exception:
        pass

    dates, venues, buyUrls, buyUrls2, showTimes = [], [], [], [], []
    for _, r in df2.iterrows():
        if (str(r["dateISO"]).strip()=="" 
            and str(r["venue_name"]).strip()==""
            and str(r["buyUrl"]).strip()==""
            and str(r["buyUrl2"]).strip()==""):
            continue

        dates.append(str(r["dateISO"]).strip())
        venues.append({"name": str(r["venue_name"]).strip(),
                       "prefecture": str(r["venue_prefecture"]).strip()})
        buyUrls.append(str(r["buyUrl"]).strip())
        buyUrls2.append(str(r["buyUrl2"]).strip())
        o = _format_hhmm(r["open"]) if has_open else ""
        s = _format_hhmm(r["start"]) if has_start else ""
        showTimes.append({"open": o, "start": s})
    return dates, venues, buyUrls, buyUrls2, showTimes

def read_sales_sheet(df):
    required = ["label","areas","seats","startISO","endISO"]
    if not set(required).issubset(df.columns):
        raise ValueError(f"Sales sheet requires columns: {required}")
    df2 = df.copy()
    for c in required: df2[c] = df2[c].fillna("")
    sales = []
    for _, r in df2.iterrows():
        if str(r["label"]).strip()=="":
            continue
        sales.append({
            "label": str(r["label"]).strip(),
            "areas": [a.strip() for a in str(r["areas"]).split(";") if a.strip()] if str(r["areas"]).strip()!="" else [],
            "seats": [s.strip() for s in str(r["seats"]).split(";") if s.strip()] if str(r["seats"]).strip()!="" else [],
            "startISO": str(r["startISO"]).strip(),
            "endISO": str(r["endISO"]).strip(),
        })
    return sales

def read_tickets_sheet(df):
    required = ["name","price","buyUrl"]
    if not set(required).issubset(df.columns):
        raise ValueError(f"Tickets sheet requires columns: {required}")
    df2 = df.copy()
    for c in required: df2[c] = df2[c].fillna("")
    tickets = []
    for _, r in df2.iterrows():
        if str(r["name"]).strip()=="":
            continue
        price = r["price"]
        if isinstance(price, float) and price.is_integer():
            price = int(price)
        elif isinstance(price, str):
            try: price = int(float(price))
            except Exception: pass
        tickets.append({
            "name": str(r["name"]).strip(),
            "price": price,
            "buyUrl": str(r["buyUrl"]).strip(),
        })
    return tickets

def main(inp, outp):
    xls = pd.ExcelFile(inp)
    for sn in ["Event","Shows","Sales","Tickets"]:
        if sn not in xls.sheet_names:
            raise ValueError(f"Missing sheet: {sn}")
    df_event = pd.read_excel(inp, sheet_name="Event")
    df_shows = pd.read_excel(inp, sheet_name="Shows")
    df_sales = pd.read_excel(inp, sheet_name="Sales")
    df_tickets = pd.read_excel(inp, sheet_name="Tickets")

    result = read_event_sheet(df_event)
    dates, venues, buyUrls, buyUrls2, showTimes = read_shows_sheet(df_shows)
    result["dates"] = dates
    result["venues"] = venues
    result["buyUrls"] = buyUrls
    result["buyUrls2"] = buyUrls2
    result["showTimes"] = showTimes
    result["sales"] = read_sales_sheet(df_sales)
    result["tickets"] = read_tickets_sheet(df_tickets)

    with open(outp, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python excel_to_event_json.py input.xlsx output.json")
        sys.exit(1)
    main(sys.argv[1], sys.argv[2])
