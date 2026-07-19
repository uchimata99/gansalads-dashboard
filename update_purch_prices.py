#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
עדכון "מחיר קנייה אחרון" בקטלוג הרכש (לשונית PURCHASING) מכרטסת מלאי שבועית.

מיכל מעלה לתיקייה "מחיר קנייה אחרון" (תחת "קניית ירקות") כרטסת מלאי של השבוע
(פורמט מחשבשבת רב-פריטים). לכל פריט מחלצים את **מחיר הרכש האחרון** — שורת
"חשבונית רכש" האחרונה (מחיר נטו) — ומעדכנים את שדה `p`/`pd` בקטלוג. פריטים
בלי חשבונית רכש בשבוע (רק מכירות/משלוחים) נשארים ללא שינוי.

אבטחה: המחירים נכנסים אך ורק לגיליון, לא לקוד הציבורי. הקטלוג נשמר כ-base64
בעמודה A של לשונית PURCHASING (אותו מנגנון כמו ingest_purchasing.py); שאר
הלשוניות (PO_HISTORY, SUP_MAP, ...) לא נגעות. SUP_MAP מוחל בטעינה בדף — לא כאן.

הרצה:
  python3 update_purch_prices.py <כרטסת_שבועית.xlsx> --key <SA.json>          # יבש
  python3 update_purch_prices.py <כרטסת_שבועית.xlsx> --key <SA.json> --apply  # כתיבה
"""
import argparse
import base64
import json
import re
import sys
from datetime import datetime, date

SHEET_ID = "1rWHMhO8zCB8KKzAJwyFYpuKfo_EQ_-rZB8afaiqUv9Q"
TAB = "PURCHASING"
CHUNK = 40000


def last_purchase_prices(path):
    """כרטסת מחשבשבת -> {מפתח פריט: {'name','price','date'}} — חשבונית רכש אחרונה."""
    from openpyxl import load_workbook
    ws = load_workbook(path, read_only=True, data_only=True).active
    ci = None
    cur, out = None, {}
    for row in ws.iter_rows(values_only=True):
        vals = list(row)
        s0 = str(vals[0] or '').strip()
        s1 = str(vals[1] or '').strip() if len(vals) > 1 else ''
        s2 = str(vals[2] or '').strip() if len(vals) > 2 else ''
        if ci is None:
            sv = [str(c).strip() if c is not None else '' for c in vals]
            if 'מחיר נטו' in sv and 'כניסה' in sv:
                ci = {k: sv.index(k) for k in ['סוג מסמך', 'תאריך', 'מחיר נטו'] if k in sv}
            continue
        if s0.startswith('סה'):
            cur = None
            continue
        if s0 and re.match(r'^\d{4,6}$', s1) and re.match(r'^\d{3}$', s2):
            cur = {'key': s1, 'name': s0}
            continue
        if cur is None:
            continue
        if str(vals[ci['סוג מסמך']] or '').strip() != 'חשבונית רכש':
            continue
        price = vals[ci['מחיר נטו']]
        try:
            price = float(price)
        except (TypeError, ValueError):
            continue
        if price <= 0:
            continue
        d = vals[ci['תאריך']]
        dt = d if isinstance(d, (datetime, date)) else None
        prev = out.get(cur['key'])
        if prev is None or (dt and (prev['dt'] is None or dt >= prev['dt'])):
            out[cur['key']] = {'name': cur['name'], 'price': round(price, 3), 'dt': dt,
                               'date': dt.strftime('%d/%m/%y') if dt else ''}
    return out


def _decode_cat(s):
    """קטלוג נשמר כ-base64; נופלים לקריאת JSON גולמי אם אינו base64."""
    try:
        return json.loads(base64.b64decode(s).decode("utf-8"))
    except Exception:
        return json.loads(s)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("ledger", help="כרטסת מלאי שבועית (xlsx)")
    ap.add_argument("--key", required=True, help="מפתח חשבון שירות (JSON)")
    ap.add_argument("--sheet", default=SHEET_ID)
    ap.add_argument("--apply", action="store_true", help="כתיבה בפועל (אחרת יבש)")
    args = ap.parse_args()

    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build
    scope = "https://www.googleapis.com/auth/spreadsheets" if args.apply \
        else "https://www.googleapis.com/auth/spreadsheets.readonly"
    creds = Credentials.from_service_account_file(args.key, scopes=[scope])
    sh = build("sheets", "v4", credentials=creds, cache_discovery=False).spreadsheets()

    prices = last_purchase_prices(args.ledger)
    if not prices:
        sys.exit("לא נמצאו חשבוניות רכש בכרטסת.")

    rows = sh.values().get(spreadsheetId=args.sheet, range=f"{TAB}!A:A").execute().get("values", [])
    cat = _decode_cat("".join(r[0] for r in rows if r))
    bykey = {str(it["k"]): it for it in cat["items"]}

    changed, missing = [], []
    for k, v in prices.items():
        it = bykey.get(str(k))
        if not it:
            missing.append((k, v["name"]))
            continue
        old = it.get("p")
        it["p"] = v["price"]
        it["pd"] = v["date"]
        changed.append((v["name"], old, v["price"]))

    print(f"פריטים שנקנו בכרטסת: {len(prices)} | עודכנו בקטלוג: {len(changed)} | לא בקטלוג: {len(missing)}")
    for n, o, p in changed:
        print(f"  {n[:30]:<32} {str(o):>9} → ₪{p}")
    for k, n in missing:
        print(f"  ⚠️ {k} {n} — לא בקטלוג (דלג)")
    if not args.apply:
        print("\n(הרצה יבשה — להוספת --apply לכתיבה)")
        return
    if not changed:
        print("אין מה לעדכן.")
        return

    cat["pricesUpdatedAt"] = date.today().isoformat()
    js = json.dumps(cat, ensure_ascii=False, separators=(",", ":"))
    b64 = base64.b64encode(js.encode("utf-8")).decode("ascii")   # הקטלוג נשמר כ-base64
    chunks = [b64[i:i + CHUNK] for i in range(0, len(b64), CHUNK)]
    sh.values().clear(spreadsheetId=args.sheet, range=f"{TAB}!A:A").execute()
    sh.values().update(spreadsheetId=args.sheet, range=f"{TAB}!A1",
                       valueInputOption="RAW", body={"values": [[c] for c in chunks]}).execute()
    back = sh.values().get(spreadsheetId=args.sheet, range=f"{TAB}!A:A").execute().get("values", [])
    ok = _decode_cat("".join(r[0] for r in back if r)) == cat
    print(f"\nנכתב לקטלוג ({len(chunks)} שורות). אימות הלוך-ושוב: {'תקין' if ok else 'נכשל'}")
    if not ok:
        sys.exit(1)


if __name__ == "__main__":
    main()
