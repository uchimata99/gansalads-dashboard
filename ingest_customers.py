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


def load_costs(sh):
    """קריאת מפת העלויות (עלות לאריזה לכל מוצר) מלשונית DATA של הייצור."""
    try:
        rows = sh.values().get(spreadsheetId=SHEET_ID, range="DATA!A:A").execute().get("values", [])
        obj = json.loads("".join(r[0] for r in rows if r))  # לשונית DATA = JSON רגיל (לא base64)
        return obj.get("COSTS", {}) or {}
    except Exception as e:
        print("אזהרה: לא ניתן לקרוא COSTS מלשונית DATA —", e)
        return {}


def enrich_profit(D, costs):
    """הצלבת רכישות הלקוח (top_prod) עם עלות לאריזה → רווח גולמי לכל לקוח ולכל מוצר.
    מדויק על המוצרים המתומחרים; coverage = אחוז ההכנסות המכוסה. אם אין עלויות — דילוג."""
    if not costs:
        print("אזהרה: אין COSTS — דילוג על העשרת רווחיות."); return False
    ikg = D.get("item_kg", {}) or {}
    tot_money = tot_cost = 0.0
    for c in D.get("customers", []):
        pm = pc = pq = pk = 0.0; pr = 0
        for tp in c.get("top_prod", []):
            it = tp.get("item"); cost = costs.get(it); q = tp.get("qty", 0) or 0
            if cost and cost > 0 and q > 0:
                line_cost = cost * q; kg = q * (ikg.get(it, 0) or 0)
                tp["cost"] = round(cost, 4)
                tp["profit"] = round((tp.get("money", 0) or 0) - line_cost, 2)
                if kg: tp["profit_kg"] = round(((tp.get("money", 0) or 0) - line_cost) / kg, 4)
                pm += tp.get("money", 0) or 0; pc += line_cost; pq += q; pk += kg; pr += 1
        prof = pm - pc; money = c.get("money") or 0
        c["gp"] = dict(money_priced=round(pm, 2), cost=round(pc, 2), profit=round(prof, 2),
                       margin=round(prof / pm * 100, 2) if pm else 0,
                       per_box=round(prof / pq, 4) if pq else 0,
                       per_kg=round(prof / pk, 4) if pk else 0,
                       coverage=round(pm / money * 100, 1) if money else 0,
                       priced_items=pr)
        tot_money += pm; tot_cost += pc
    D["gp_meta"] = dict(money_priced=round(tot_money, 2), cost=round(tot_cost, 2),
                        profit=round(tot_money - tot_cost, 2),
                        margin=round((tot_money - tot_cost) / tot_money * 100, 2) if tot_money else 0,
                        priced_costs=len(costs))
    print(f"העשרת רווחיות: {len(D.get('customers', []))} לקוחות | רווח גולמי כולל "
          f"₪{tot_money - tot_cost:,.0f} (על מוצרי הטופ המתומחרים).")
    return True


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("data_file", help="dashboard_customers.html או קובץ .json")
    ap.add_argument("--key", default=os.environ.get("GOOGLE_APPLICATION_CREDENTIALS"),
                    help="נתיב למפתח חשבון השירות (JSON)")
    ap.add_argument("--no-profit", action="store_true", help="לדלג על העשרת רווח גולמי")
    args = ap.parse_args()
    if not args.key:
        sys.exit("חסר מפתח חשבון שירות: --key או GOOGLE_APPLICATION_CREDENTIALS")

    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build

    obj = json.loads(extract_json(args.data_file))

    creds = Credentials.from_service_account_file(
        args.key, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    sh = build("sheets", "v4", credentials=creds, cache_discovery=False).spreadsheets()

    # העשרת רווח גולמי בהצלבה עם עלויות הייצור (לפי שם מוצר)
    if not args.no_profit:
        enrich_profit(obj, load_costs(sh))

    json_str = json.dumps(obj, ensure_ascii=False, separators=(", ", ": "))
    b64 = base64.b64encode(json_str.encode("utf-8")).decode("ascii")
    chunks = [b64[i:i + CHUNK] for i in range(0, len(b64), CHUNK)]

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
    gp = obj.get("gp_meta", {}).get("profit")
    print(f"נכתבו {len(chunks)} שורות ללשונית {TAB} | לקוחות: {n_cust} | מחזור: {money} | "
          f"רווח גולמי: {gp} | אימות זהות: {'תקין' if ok else 'נכשל'}")
    if not ok:
        sys.exit(1)


if __name__ == "__main__":
    main()
