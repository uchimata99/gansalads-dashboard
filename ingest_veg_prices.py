#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
קליטת ניתוח מחירי הירקות אל הגיליון המוגבל (מקבילה ל-ingest_purchasing.py).

מה זה עושה
----------
מקבל את veg_prices.json (פלט build_veg_prices.py), מקודד ל-base64, מחלק לשורות,
וכותב אותן לעמודה A בלשונית VEG_PRICES שבגיליון הרכש. orders_report.html קורא
את אותה לשונית בזמן ריצה אחרי כניסת גוגל (base64 -> JSON), בדיוק כמו CUSTOMERS.
הלשונית נכתבת מחדש בכל קליטה (עמודה A מנוקה) — אין היסטוריה מצטברת כאן.

אבטחה: מחירי הקנייה נכנסים אך ורק לגיליון (קריאה אחרי כניסת גוגל מורשית),
לעולם לא לקוד הציבורי. מפתח חשבון השירות מסופק בזמן ריצה.

הרצה
----
  python3 ingest_veg_prices.py veg_prices.json --key <service_account.json>

חשבון השירות שכותב: gansalads-sheet-writer@gansalads-dashboards.iam.gserviceaccount.com
(אותו חשבון שכותב לגיליון הרכש/לקוחות).
"""
import argparse
import base64
import json
import os
import sys

# גיליון הרכש הייעודי — זהה למזהה ב-purchasing.html / orders_report.html
SHEET_ID = "1rWHMhO8zCB8KKzAJwyFYpuKfo_EQ_-rZB8afaiqUv9Q"
CHUNK = 40000  # תווי base64 לכל תא (מתחת למגבלת 50,000 של גוגל שיטס)


def main():
    global SHEET_ID
    ap = argparse.ArgumentParser()
    ap.add_argument("payload", help="veg_prices.json / mat_prices.json (פלט build_veg_prices.py)")
    ap.add_argument("--key", default=os.environ.get("GOOGLE_APPLICATION_CREDENTIALS"),
                    help="נתיב למפתח חשבון השירות (JSON)")
    ap.add_argument("--sheet", default=SHEET_ID, help="מזהה גיליון יעד (ברירת מחדל: גיליון הרכש)")
    ap.add_argument("--tab", default=None,
                    help="לשונית יעד; ברירת מחדל לפי סוג הקובץ: veg->VEG_PRICES, items->MAT_PRICES")
    args = ap.parse_args()
    if not args.key:
        sys.exit("חסר מפתח חשבון שירות: --key או GOOGLE_APPLICATION_CREDENTIALS")
    SHEET_ID = args.sheet

    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build

    obj = json.loads(open(args.payload, "r", encoding="utf-8").read())
    if "returns" not in obj or ("veg" not in obj and "items" not in obj):
        sys.exit("הקובץ אינו ניתוח מחירים תקין (חסר veg/items/returns).")
    TAB = args.tab or ("VEG_PRICES" if "veg" in obj else "MAT_PRICES")

    creds = Credentials.from_service_account_file(
        args.key, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    sh = build("sheets", "v4", credentials=creds, cache_discovery=False).spreadsheets()

    # העשרת שמות ספקים + תוויות קבוצות מקטלוג הרכש (לשונית PURCHASING) — כדי
    # שהשמות לא יישמרו בקוד הציבורי אלא ייגזרו בזמן הקליטה מהגיליון עצמו.
    try:
        prows = sh.values().get(spreadsheetId=SHEET_ID, range="PURCHASING!A:A").execute().get("values", [])
        cat = json.loads(base64.b64decode("".join(r[0] for r in prows if r)).decode("utf-8"))
        names = cat.get("suppliers", {})
        enriched = 0
        for code in list(obj.get("suppliers", {})):
            if code in names and names[code]:
                obj["suppliers"][code] = names[code]
                enriched += 1
        print(f"שמות ספקים שהועשרו מהקטלוג: {enriched}/{len(obj.get('suppliers', {}))}")
        if "groups" in obj:
            cats = cat.get("cats", {})
            g_en = 0
            for g in list(obj["groups"]):
                if g in cats and cats[g]:
                    obj["groups"][g] = cats[g]
                    g_en += 1
            print(f"תוויות קבוצות שהועשרו מהקטלוג: {g_en}/{len(obj['groups'])}")
    except Exception as e:
        print(f"אזהרה: לא ניתן היה להעשיר שמות מהקטלוג ({e}); נשמרים קודים.")

    json_str = json.dumps(obj, ensure_ascii=False, separators=(",", ":"))
    b64 = base64.b64encode(json_str.encode("utf-8")).decode("ascii")
    chunks = [b64[i:i + CHUNK] for i in range(0, len(b64), CHUNK)]

    tabs = [s["properties"]["title"] for s in
            sh.get(spreadsheetId=SHEET_ID, fields="sheets.properties").execute()["sheets"]]
    if TAB not in tabs:
        sh.batchUpdate(spreadsheetId=SHEET_ID,
                       body={"requests": [{"addSheet": {"properties": {"title": TAB}}}]}).execute()

    # עמודה A מנוקה ונכתבת מחדש בכל קליטה
    sh.values().clear(spreadsheetId=SHEET_ID, range=f"{TAB}!A:A").execute()
    sh.values().update(spreadsheetId=SHEET_ID, range=f"{TAB}!A1",
                       valueInputOption="RAW", body={"values": [[c] for c in chunks]}).execute()

    # אימות הלוך-ושוב
    rows = sh.values().get(spreadsheetId=SHEET_ID, range=f"{TAB}!A:A").execute().get("values", [])
    back = base64.b64decode("".join(r[0] for r in rows if r)).decode("utf-8")
    ok = json.loads(back) == obj
    n = len(obj.get("veg") or obj.get("items") or [])
    print(f"נכתבו {len(chunks)} שורות ללשונית {TAB} | פריטים: {n} | "
          f"חודשים: {len(obj.get('months', []))} | החזרות: {obj['returns']['totalKg']} ק\"ג | "
          f"אימות זהות: {'תקין' if ok else 'נכשל'}")
    if not ok:
        sys.exit(1)


if __name__ == "__main__":
    main()
