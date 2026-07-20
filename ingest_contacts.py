#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
קליטת ספר הטלפונים של הספקים אל הגיליון המוגבל (לשונית CONTACTS).

מה זה עושה
----------
קורא את קובץ אנשי הקשר שמיכל מילאה (`אנשי_קשר_ספקים_רכש.xlsx` — עמודות:
מפתח ספק, שם ספק, טלפון וואטסאפ, מייל), מנרמל את הטלפון לפורמט בינלאומי
(0→972, בלי תווים לא-ספרתיים), וכותב שורות פשוטות ללשונית CONTACTS שבגיליון
הרכש. purchasing.html קורא את הלשונית בזמן ריצה (loadContacts) וממפה מפתח-ספק
→ {טלפון, מייל}: הטלפון משמש לשליחת וואטסאפ, המייל לשלב המייל האוטומטי.

מבנה הלשונית CONTACTS (שורות רגילות, לא base64 — אין כאן מחירים):
  [מפתח ספק, שם ספק, טלפון (E.164 בלי +), מייל]
עמודה A נכתבת מחדש בכל קליטה. נכתבות רק שורות עם טלפון או מייל.

הרצה
----
  python3 ingest_contacts.py אנשי_קשר_ספקים_רכש.xlsx --key <service_account.json>
  # או:  GOOGLE_APPLICATION_CREDENTIALS=sa_key.json python3 ingest_contacts.py <file.xlsx>

חשבון השירות שכותב: gansalads-sheet-writer@gansalads-dashboards.iam.gserviceaccount.com
"""
import argparse
import os
import re
import sys

SHEET_ID = "1rWHMhO8zCB8KKzAJwyFYpuKfo_EQ_-rZB8afaiqUv9Q"  # גיליון הרכש הייעודי
CONTACTS_TAB = "CONTACTS"
HEADER = ["מפתח ספק", "שם ספק", "טלפון", "מייל"]


def norm_phone(v):
    """מנרמל טלפון לפורמט בינלאומי בלי + (E.164): 0→972, מסיר תווים לא-ספרתיים."""
    if v is None:
        return ""
    digits = re.sub(r"\D", "", str(v).strip())
    if not digits:
        return ""
    if digits.startswith("972"):
        return digits
    if digits.startswith("0"):
        return "972" + digits[1:]
    return digits


def norm_email(v):
    if v is None:
        return ""
    e = str(v).strip()
    return e if "@" in e else ""


def col_index(headers, *names):
    """מאתר עמודה לפי כותרת (מכיל אחד מהשמות), לא לפי מיקום קבוע."""
    for i, h in enumerate(headers):
        hs = str(h or "").strip()
        for n in names:
            if n in hs:
                return i
    return -1


def main():
    global SHEET_ID
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx", help="קובץ אנשי הקשר (xlsx) שמיכל מילאה")
    ap.add_argument("--key", default=os.environ.get("GOOGLE_APPLICATION_CREDENTIALS"),
                    help="נתיב למפתח חשבון השירות (JSON)")
    ap.add_argument("--sheet", default=SHEET_ID, help="מזהה גיליון יעד (ברירת מחדל: גיליון הרכש)")
    ap.add_argument("--dry", action="store_true", help="הרצה יבשה — הצגה בלבד, בלי כתיבה")
    args = ap.parse_args()
    SHEET_ID = args.sheet

    import openpyxl
    wb = openpyxl.load_workbook(args.xlsx, read_only=True, data_only=True)
    ws = wb.active
    rows_in = list(ws.iter_rows(values_only=True))
    if not rows_in:
        sys.exit("הקובץ ריק.")
    headers = [str(c or "").strip() for c in rows_in[0]]
    ci_key = col_index(headers, "מפתח")
    ci_name = col_index(headers, "שם")
    ci_phone = col_index(headers, "טלפון", "וואטסאפ", "נייד")
    ci_mail = col_index(headers, "מייל", "מייל", "אימייל", "דוא")
    if ci_key < 0 or ci_name < 0:
        sys.exit(f"לא נמצאו עמודות 'מפתח ספק'/'שם ספק' בכותרות: {headers}")

    out, n_phone, n_mail = [], 0, 0
    for r in rows_in[1:]:
        key = str(r[ci_key] if ci_key < len(r) and r[ci_key] is not None else "").strip()
        name = str(r[ci_name] if ci_name < len(r) and r[ci_name] is not None else "").strip()
        phone = norm_phone(r[ci_phone]) if 0 <= ci_phone < len(r) else ""
        mail = norm_email(r[ci_mail]) if 0 <= ci_mail < len(r) else ""
        if not key:
            continue
        if not phone and not mail:
            continue  # אין ליצור שורה לספק בלי טלפון ובלי מייל
        out.append([key, name, phone, mail])
        n_phone += 1 if phone else 0
        n_mail += 1 if mail else 0

    print(f"נמצאו {len(out)} ספקים עם פרטי קשר | טלפון: {n_phone} | מייל: {n_mail}")
    if args.dry:
        for row in out[:10]:
            print("  ", row)
        if len(out) > 10:
            print(f"  … ועוד {len(out) - 10}")
        return
    if not args.key:
        sys.exit("חסר מפתח חשבון שירות: --key או GOOGLE_APPLICATION_CREDENTIALS (או השתמש ב---dry)")

    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build
    creds = Credentials.from_service_account_file(
        args.key, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    sh = build("sheets", "v4", credentials=creds, cache_discovery=False).spreadsheets()

    tabs = [s["properties"]["title"] for s in
            sh.get(spreadsheetId=SHEET_ID, fields="sheets.properties").execute()["sheets"]]
    if CONTACTS_TAB not in tabs:
        sh.batchUpdate(spreadsheetId=SHEET_ID,
                       body={"requests": [{"addSheet": {"properties": {"title": CONTACTS_TAB}}}]}).execute()

    sh.values().clear(spreadsheetId=SHEET_ID, range=f"{CONTACTS_TAB}!A:D").execute()
    sh.values().update(spreadsheetId=SHEET_ID, range=f"{CONTACTS_TAB}!A1",
                       valueInputOption="RAW", body={"values": [HEADER] + out}).execute()

    back = sh.values().get(spreadsheetId=SHEET_ID, range=f"{CONTACTS_TAB}!A:D").execute().get("values", [])
    print(f"נכתבו {len(out)} שורות ללשונית {CONTACTS_TAB} | אימות: {len(back) - 1} שורות בגיליון")


if __name__ == "__main__":
    main()
