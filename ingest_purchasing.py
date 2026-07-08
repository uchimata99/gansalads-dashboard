#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
קליטת קטלוג הרכש אל הגיליון המוגבל (מקבילה ל-ingest_customers.py).

מה זה עושה
----------
מקבל את purchasing_catalog.json (פלט build_purchasing.py), מקודד ל-base64,
מחלק לשורות, וכותב אותן לעמודה A בלשונית PURCHASING שבגיליון המוגבל.
purchasing.html קורא את אותה לשונית בזמן ריצה אחרי כניסת גוגל.
בנוסף מוודא שקיימת לשונית PO_HISTORY (היסטוריית הזמנות נצברת) עם שורת כותרת —
לשונית זו לעולם לא מנוקה; האפליקציה מוסיפה אליה שורות (append).

אבטחה: מחירי הקנייה נכנסים אך ורק לגיליון המוגבל (קריאה אחרי כניסת גוגל
מורשית), לעולם לא לקוד הציבורי. מפתח חשבון השירות מסופק בזמן ריצה.

הרצה
----
  python3 ingest_purchasing.py purchasing_catalog.json --key <service_account.json>
  # או:  GOOGLE_APPLICATION_CREDENTIALS=sa_key.json python3 ingest_purchasing.py purchasing_catalog.json

חשבון השירות שכותב: gansalads-sheet-writer@gansalads-dashboards.iam.gserviceaccount.com
(אותו חשבון שכותב לגיליון הייצור/לקוחות).
"""
import argparse
import base64
import json
import os
import sys

# גיליון הרכש הייעודי — זהה למזהה ב-purchasing.html (נפרד משאר הדשבורדים)
SHEET_ID = "1rWHMhO8zCB8KKzAJwyFYpuKfo_EQ_-rZB8afaiqUv9Q"
CAT_TAB = "PURCHASING"
HIST_TAB = "PO_HISTORY"
CHUNK = 40000  # תווי base64 לכל תא (מתחת למגבלת 50,000 של גוגל שיטס)
HIST_HEADER = ["תאריך", "מזהה הזמנה", "מפתח ספק", "שם ספק",
               "מפתח פריט", "שם פריט", "כמות", "מחיר אחרון", "הערכת עלות"]


def main():
    global SHEET_ID
    ap = argparse.ArgumentParser()
    ap.add_argument("catalog", help="purchasing_catalog.json (פלט build_purchasing.py)")
    ap.add_argument("--key", default=os.environ.get("GOOGLE_APPLICATION_CREDENTIALS"),
                    help="נתיב למפתח חשבון השירות (JSON)")
    ap.add_argument("--sheet", default=SHEET_ID, help="מזהה גיליון יעד (ברירת מחדל: גיליון הרכש)")
    args = ap.parse_args()
    if not args.key:
        sys.exit("חסר מפתח חשבון שירות: --key או GOOGLE_APPLICATION_CREDENTIALS")
    SHEET_ID = args.sheet

    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build

    obj = json.loads(open(args.catalog, "r", encoding="utf-8").read())
    if "items" not in obj or "suppliers" not in obj:
        sys.exit("הקובץ אינו קטלוג רכש תקין (חסר items/suppliers).")

    creds = Credentials.from_service_account_file(
        args.key, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    sh = build("sheets", "v4", credentials=creds, cache_discovery=False).spreadsheets()

    json_str = json.dumps(obj, ensure_ascii=False, separators=(",", ":"))
    b64 = base64.b64encode(json_str.encode("utf-8")).decode("ascii")
    chunks = [b64[i:i + CHUNK] for i in range(0, len(b64), CHUNK)]

    tabs = [s["properties"]["title"] for s in
            sh.get(spreadsheetId=SHEET_ID, fields="sheets.properties").execute()["sheets"]]
    reqs = []
    if CAT_TAB not in tabs:
        reqs.append({"addSheet": {"properties": {"title": CAT_TAB}}})
    if HIST_TAB not in tabs:
        reqs.append({"addSheet": {"properties": {"title": HIST_TAB}}})
    if reqs:
        sh.batchUpdate(spreadsheetId=SHEET_ID, body={"requests": reqs}).execute()

    # כתיבת הקטלוג (עמודה A מנוקה ונכתבת מחדש בכל קליטה)
    sh.values().clear(spreadsheetId=SHEET_ID, range=f"{CAT_TAB}!A:A").execute()
    sh.values().update(spreadsheetId=SHEET_ID, range=f"{CAT_TAB}!A1",
                       valueInputOption="RAW",
                       body={"values": [[c] for c in chunks]}).execute()

    # שורת כותרת ל-PO_HISTORY רק אם הלשונית ריקה (לא נוגעים בהיסטוריה קיימת)
    hist = sh.values().get(spreadsheetId=SHEET_ID, range=f"{HIST_TAB}!A1:I1").execute().get("values", [])
    if not hist:
        sh.values().update(spreadsheetId=SHEET_ID, range=f"{HIST_TAB}!A1",
                           valueInputOption="RAW", body={"values": [HIST_HEADER]}).execute()

    # אימות הלוך-ושוב של הקטלוג
    rows = sh.values().get(spreadsheetId=SHEET_ID, range=f"{CAT_TAB}!A:A").execute().get("values", [])
    back = base64.b64decode("".join(r[0] for r in rows if r)).decode("utf-8")
    ok = json.loads(back) == obj
    with_price = sum(1 for it in obj["items"] if it.get("p") is not None)
    with_sup = sum(1 for it in obj["items"] if it.get("s"))
    print(f"נכתבו {len(chunks)} שורות ללשונית {CAT_TAB} | פריטים: {len(obj['items'])} | "
          f"עם ספק: {with_sup} | עם מחיר: {with_price} | ספקים: {len(obj['suppliers'])} | "
          f"אימות זהות: {'תקין' if ok else 'נכשל'}")
    if not ok:
        sys.exit(1)


if __name__ == "__main__":
    main()
