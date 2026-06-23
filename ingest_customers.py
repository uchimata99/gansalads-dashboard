#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
קליטת נתוני דשבורד הלקוחות אל הגיליון המוגבל (מקבילה ל-ingest_week.py של הייצור).

מה זה עושה
----------
מקבל קובץ נתונים שבועי (אחד משניים):
  • dashboard_customers.html  — מחלץ ממנו את בלוק `const D={...}`
  • קובץ .json               — אובייקט הנתונים D עצמו
מקודד את ה-JSON ל-base64, מחלק אותו לשורות, וכותב אותם לעמודה A
בלשונית CUSTOMERS שבגיליון המוגבל. הדשבורד (dashboard_customers.html)
קורא את אותה לשונית בזמן ריצה אחרי כניסת גוגל. הקובץ עצמו לא משתנה כל שבוע.

מקור האמת הוא הקובץ שמיכל מעלה; הגיליון מכיל רק את ה-JSON המקודד.
אין סודות בקובץ הזה — מפתח חשבון השירות מסופק בזמן ריצה.

הרצה
----
  python3 ingest_customers.py <data_file> --key <service_account.json>
  # או עם משתנה סביבה:
  GOOGLE_APPLICATION_CREDENTIALS=sa_key.json python3 ingest_customers.py data.html

חשבון השירות שכותב: gansalads-sheet-writer@gansalads-dashboards.iam.gserviceaccount.com
(אותו חשבון שכותב לגיליון הייצור — יש לו הרשאת כתיבה לגיליון).
"""
import argparse
import base64
import json
import os
import re
import sys

# הגיליון המוגבל ולשונית היעד — זהה למזהה ב-dashboard_customers.html
SHEET_ID = "1_rJ8lLYNme8RM83ws1pFOGij0HPAVAUFefBDYl2_B5A"
TAB = "CUSTOMERS"
CHUNK = 40000  # תווי base64 לכל תא (מתחת למגבלת 50,000 של גוגל שיטס)


def extract_json(path):
    """מחזיר מחרוזת JSON תקינה מתוך קובץ HTML (const D={...}) או .json."""
    raw = open(path, "r", encoding="utf-8").read()
    if path.lower().endswith(".json"):
        text = raw
    else:
        m = re.search(r"const\s+D\s*=\s*(\{.*?\})\s*;", raw, re.DOTALL)
        if not m:
            sys.exit("לא נמצא בלוק `const D={...}` בקובץ ה-HTML.")
        text = m.group(1)
    obj = json.loads(text)  # אימות שהוא JSON תקין
    return json.dumps(obj, ensure_ascii=False, separators=(", ", ": "))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("data_file", help="dashboard_customers.html או קובץ .json")
    ap.add_argument("--key", default=os.environ.get("GOOGLE_APPLICATION_CREDENTIALS"),
                    help="נתיב למפתח חשבון השירות (JSON)")
    args = ap.parse_args()
    if not args.key:
        sys.exit("חסר מפתח חשבון שירות: --key או GOOGLE_APPLICATION_CREDENTIALS")

    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build

    json_str = extract_json(args.data_file)
    obj = json.loads(json_str)
    b64 = base64.b64encode(json_str.encode("utf-8")).decode("ascii")
    chunks = [b64[i:i + CHUNK] for i in range(0, len(b64), CHUNK)]

    creds = Credentials.from_service_account_file(
        args.key, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    sh = build("sheets", "v4", credentials=creds, cache_discovery=False).spreadsheets()

    tabs = [s["properties"]["title"] for s in
            sh.get(spreadsheetId=SHEET_ID, fields="sheets.properties").execute()["sheets"]]
    if TAB not in tabs:
        sh.batchUpdate(spreadsheetId=SHEET_ID,
                       body={"requests": [{"addSheet": {"properties": {"title": TAB}}}]}).execute()

    sh.values().clear(spreadsheetId=SHEET_ID, range=f"{TAB}!A:A").execute()
    sh.values().update(spreadsheetId=SHEET_ID, range=f"{TAB}!A1",
                       valueInputOption="RAW",
                       body={"values": [[c] for c in chunks]}).execute()

    # אימות הלוך-ושוב
    rows = sh.values().get(spreadsheetId=SHEET_ID, range=f"{TAB}!A:A").execute().get("values", [])
    back = base64.b64decode("".join(r[0] for r in rows if r)).decode("utf-8")
    ok = json.loads(back) == obj
    n_cust = len(obj.get("customers", []))
    money = obj.get("meta", {}).get("total_money")
    print(f"נכתבו {len(chunks)} שורות ללשונית {TAB} | לקוחות: {n_cust} | "
          f"מחזור: {money} | אימות זהות: {'תקין' if ok else 'נכשל'}")
    if not ok:
        sys.exit(1)


if __name__ == "__main__":
    main()
