#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
בניית ניתוח מחירי ירקות מכרטסות המלאי של מחשבשבת (קוד מחויב, לא צ'אט).

מקור הנתונים
------------
תיקיית דרייב "קניית ירקות" — כרטסת מלאי לכל ירק (פורמט מחשבשבת: כרטסת מלאי
לפי תאריכי רישום). כל שורה = תנועת מלאי; אותנו מעניינות שורות "חשבונית רכש":
מחיר נטו (₪ לק"ג), כניסה (ק"ג), תאריך. שורת "החזרת רכש" / ערך ב"יציאה" = החזרה.

מה זה מייצר
-----------
veg_prices.json — אובייקט הניתוח שהדשבורד קורא (אחרי כניסת גוגל, מהגיליון):
  { generatedAt, currency, period:{from,to}, months:[...],
    suppliers:{code:label},
    veg:[ {key,name,totalKg,totalSpend,avgPrice,minPrice,maxPrice,nTxn,
           firstMonth,lastMonth,
           monthly:[{m,kg,spend,avg,min,max,n}],
           bySupplier:[{code,kg,spend,avg,n}]} ],
    returns:{totalKg, rows:[{veg,date,code,kg,price}]},
    totals:{kg,spend,txns,nVeg} }

מחיר חודשי = ממוצע משוקלל בק"ג (spend/kg) — עמיד לקניות קטנות/גדולות.

אבטחה: מחירי הקנייה הם נתון רגיש. veg_prices.json **אינו מחויב למאגר** (מכיל
מחירים) — נבנה ונקלט לגיליון בלבד, בדיוק כמו purchasing_catalog.json.

הרצה
----
  # xlsx ישירות מהתיקייה:
  python3 build_veg_prices.py "כרטסת מלאי *.xlsx" --out veg_prices.json
  # או ייצוא טקסט/CSV של אותה כרטסת:
  python3 build_veg_prices.py veg/*.txt --out veg_prices.json
"""
import argparse
import glob
import json
import re
import sys
from datetime import date, datetime

DOC_PURCHASE = "חשבונית רכש"
RETURN_HINT = "החזר"   # "החזרת רכש" וכו'

# תוויות ידועות לחשבונות ספק (קוד -> שם). אם קוד לא ידוע — מוצג הקוד עצמו.
SUPPLIER_LABELS = {
    "20615": "ספק 20615",
    "20100": "ספק 20100",
}


def _num(x):
    """מחרוזת/מספר -> float, כולל הסרת מפרידי אלפים ומרכאות."""
    if x is None:
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip().replace('"', '').replace(',', '')
    if s == '' or s == '—':
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def _month(datestr):
    """DD/MM/YY (או datetime) -> 'YYYY-MM'. מחזיר None אם לא ניתן לפענח."""
    if isinstance(datestr, (datetime, date)):
        return f"{datestr.year:04d}-{datestr.month:02d}"
    m = re.match(r'\s*(\d{1,2})/(\d{1,2})/(\d{2,4})', str(datestr or ''))
    if not m:
        return None
    yy = int(m.group(3))
    year = yy + 2000 if yy < 100 else yy
    return f"{year:04d}-{int(m.group(2)):02d}"


# ------------------------- קריאת מקור -------------------------
def rows_from_text(path):
    """ייצוא טקסט/CSV של הכרטסת -> (key, name, [ {doc,code,date,price,inn,out} ], total_kg)."""
    key = name = None
    total_kg = None
    out = []
    for line in open(path, encoding='utf-8'):
        f = [c.strip() for c in line.rstrip('\n').split(',')]
        if not f:
            continue
        head = f[0]
        # שורת סיכום פריט (סמכותית לשם/מפתח): "סהכ מפתח פריט", key, name, ... , total_kg, ...
        if head.startswith('סהכ מפתח פריט') or head.startswith('סה"כ מפתח פריט'):
            if len(f) > 1:
                key = f[1]
            if len(f) > 2:
                name = f[2]
            nums = [c for c in f if re.match(r'^-?\d+(\.\d+)?$', c or '')]
            if len(nums) >= 3:
                total_kg = _num(nums[-3])
            continue
        # שורת תנועה: 3 עמודות ריקות ואז מזהה + סוג מסמך
        if len(f) >= 15 and f[0] == '' and f[1] == '' and f[2] == '' and f[4]:
            out.append({'doc': f[4], 'code': f[5], 'date': f[10],
                        'price': _num(f[12]), 'inn': _num(f[14]),
                        'out': _num(f[15]) if len(f) > 15 else 0.0})
    return key, name, out, total_kg


def rows_from_xlsx(path):
    """כרטסת xlsx של מחשבשבת -> אותו מבנה כמו rows_from_text (זיהוי עמודות לפי כותרת)."""
    from openpyxl import load_workbook
    ws = load_workbook(path, read_only=True, data_only=True).active
    rows = [list(r) for r in ws.iter_rows(values_only=True)]

    def cell(r, i):
        return r[i] if (i is not None and i < len(r)) else None

    # מפת עמודות מתוך שורת הכותרת (המכילה "מחיר נטו")
    col = {}
    for r in rows:
        vals = [str(c).strip() if c is not None else '' for c in r]
        if 'מחיר נטו' in vals and 'כניסה' in vals:
            for i, v in enumerate(vals):
                if v:
                    col[v] = i
            break
    need = ['סוג מסמך', 'לקוח/ספק', 'תאריך', 'מחיר נטו', 'כניסה']
    if not all(k in col for k in need):
        sys.exit(f"{path}: לא נמצאה שורת כותרת תקינה של כרטסת מחשבשבת")
    ci = {k: col.get(k) for k in
          ['מזהה', 'סוג מסמך', 'לקוח/ספק', 'תאריך', 'מחיר נטו', 'כניסה', 'יציאה']}

    key = name = None
    total_kg = None
    out = []
    for r in rows:
        c0 = str(cell(r, 0) or '').strip()
        if c0.startswith('סה') and 'מפתח פריט' in c0:      # שורת סיכום
            key = key or (str(cell(r, 1)).strip() if cell(r, 1) is not None else None)
            name = name or (str(cell(r, 2)).strip() if cell(r, 2) is not None else None)
            total_kg = _num(cell(r, ci['כניסה']))
            continue
        doc = cell(r, ci['סוג מסמך'])
        if not doc or not str(doc).strip():
            continue
        idv = cell(r, ci['מזהה']) if ci['מזהה'] is not None else None
        if idv is None or not re.match(r'^\d+$', str(idv).strip()):
            continue
        out.append({'doc': str(doc).strip(), 'code': str(cell(r, ci['לקוח/ספק']) or '').strip(),
                    'date': cell(r, ci['תאריך']), 'price': _num(cell(r, ci['מחיר נטו'])),
                    'inn': _num(cell(r, ci['כניסה'])),
                    'out': _num(cell(r, ci['יציאה'])) if ci['יציאה'] is not None else 0.0})
    return key, name, out, total_kg


# ------------------------- אגרגציה -------------------------
def build_veg(key, name, txns):
    """מבנה ניתוח לירק אחד + רשומות החזרה שנמצאו."""
    purch = [t for t in txns if RETURN_HINT not in t['doc'] and t['out'] == 0 and t['inn'] > 0]
    returns = [t for t in txns if RETURN_HINT in t['doc'] or t['out'] > 0]

    by_month = {}
    by_sup = {}
    by_year = {}          # שנה -> {kg,spend,min,max,n}
    by_year_sup = {}      # (שנה,קוד ספק) -> {kg,spend,n}
    tot_kg = tot_spend = 0.0
    pmin = pmax = None
    for t in purch:
        m = _month(t['date'])
        yr = (m or '?').split('-')[0]
        kg = t['inn']
        sp = t['price'] * kg
        tot_kg += kg
        tot_spend += sp
        pmin = t['price'] if pmin is None else min(pmin, t['price'])
        pmax = t['price'] if pmax is None else max(pmax, t['price'])
        d = by_month.setdefault(m, {'kg': 0.0, 'spend': 0.0, 'min': t['price'],
                                    'max': t['price'], 'n': 0})
        d['kg'] += kg; d['spend'] += sp; d['n'] += 1
        d['min'] = min(d['min'], t['price']); d['max'] = max(d['max'], t['price'])
        s = by_sup.setdefault(t['code'], {'kg': 0.0, 'spend': 0.0, 'n': 0})
        s['kg'] += kg; s['spend'] += sp; s['n'] += 1
        yd = by_year.setdefault(yr, {'kg': 0.0, 'spend': 0.0, 'min': t['price'],
                                     'max': t['price'], 'n': 0})
        yd['kg'] += kg; yd['spend'] += sp; yd['n'] += 1
        yd['min'] = min(yd['min'], t['price']); yd['max'] = max(yd['max'], t['price'])
        ys = by_year_sup.setdefault((yr, t['code']), {'kg': 0.0, 'spend': 0.0, 'n': 0})
        ys['kg'] += kg; ys['spend'] += sp; ys['n'] += 1

    monthly = [{'m': m, 'kg': round(v['kg'], 1), 'spend': round(v['spend'], 1),
                'avg': round(v['spend'] / v['kg'], 3) if v['kg'] else None,
                'min': round(v['min'], 3), 'max': round(v['max'], 3), 'n': v['n']}
               for m, v in sorted(by_month.items())]
    bysup = [{'code': c, 'kg': round(v['kg'], 1), 'spend': round(v['spend'], 1),
              'avg': round(v['spend'] / v['kg'], 3) if v['kg'] else None, 'n': v['n']}
             for c, v in sorted(by_sup.items(), key=lambda kv: -kv[1]['kg'])]
    annual = []
    for y in sorted(by_year):
        v = by_year[y]
        ysup = [{'code': c, 'kg': round(d['kg'], 1), 'spend': round(d['spend'], 1),
                 'avg': round(d['spend'] / d['kg'], 3) if d['kg'] else None, 'n': d['n']}
                for (yy, c), d in sorted(by_year_sup.items(), key=lambda kv: -kv[1]['kg']) if yy == y]
        annual.append({'year': y, 'kg': round(v['kg'], 1), 'spend': round(v['spend'], 1),
                       'avg': round(v['spend'] / v['kg'], 3) if v['kg'] else None,
                       'min': round(v['min'], 3), 'max': round(v['max'], 3),
                       'n': v['n'], 'bySupplier': ysup})

    veg = {
        'key': str(key), 'name': name,
        'totalKg': round(tot_kg, 1), 'totalSpend': round(tot_spend, 1),
        'avgPrice': round(tot_spend / tot_kg, 3) if tot_kg else None,
        'minPrice': round(pmin, 3) if pmin is not None else None,
        'maxPrice': round(pmax, 3) if pmax is not None else None,
        'nTxn': len(purch),
        'firstMonth': monthly[0]['m'] if monthly else None,
        'lastMonth': monthly[-1]['m'] if monthly else None,
        'monthly': monthly, 'bySupplier': bysup, 'annual': annual,
    }
    ret_rows = [{'veg': name, 'date': str(t['date']), 'code': t['code'],
                 'kg': round(t['out'] or t['inn'], 1), 'price': round(t['price'], 3)}
                for t in returns]
    return veg, ret_rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('inputs', nargs='+', help='קבצי כרטסת (xlsx או ייצוא טקסט/csv); תומך בתבניות')
    ap.add_argument('--out', default='veg_prices.json')
    ap.add_argument('--date', default=None, help="חותמת generatedAt (YYYY-MM-DD); ברירת מחדל: היום")
    args = ap.parse_args()

    paths = []
    for pat in args.inputs:
        paths.extend(sorted(glob.glob(pat)) or [pat])

    vegs = []
    ret_all = []
    sup_codes = set()
    ok = True
    for p in paths:
        reader = rows_from_xlsx if p.lower().endswith(('.xlsx', '.xlsm')) else rows_from_text
        key, name, txns, total_kg = reader(p)
        if not txns:
            print(f"⚠️  {p}: לא נמצאו תנועות — מדלגים")
            continue
        veg, rets = build_veg(key, name, txns)
        vegs.append(veg)
        ret_all.extend(rets)
        sup_codes.update(s['code'] for s in veg['bySupplier'])
        # בדיקת הצטלבות מול שורת הסיכום של הכרטסת
        chk = '—'
        if total_kg is not None:
            diff = abs(total_kg - veg['totalKg'])
            chk = 'תקין' if diff < 0.5 else f'סטייה {diff:.1f} ק"ג (דוח {total_kg})'
            ok = ok and diff < 0.5
        print(f"  {veg['name']:<14} | {veg['nTxn']:>3} תנועות | "
              f"{veg['totalKg']:>10.1f} ק\"ג | ממוצע {veg['avgPrice']} ₪ | "
              f"טווח {veg['minPrice']}–{veg['maxPrice']} | הצטלבות: {chk}")

    if not vegs:
        sys.exit("לא נבנה דבר — אין תנועות בקבצים.")

    months = sorted({m['m'] for v in vegs for m in v['monthly'] if m['m']})
    years = sorted({a['year'] for v in vegs for a in v['annual']})
    suppliers = {c: SUPPLIER_LABELS.get(c, f'ספק {c}') for c in sorted(sup_codes)}
    payload = {
        'generatedAt': args.date or date.today().isoformat(),
        'currency': '₪ לק"ג',
        'period': {'from': months[0] if months else None, 'to': months[-1] if months else None},
        'months': months,
        'years': years,
        'suppliers': suppliers,
        'veg': sorted(vegs, key=lambda v: -v['totalSpend']),
        'returns': {'totalKg': round(sum(r['kg'] for r in ret_all), 1), 'rows': ret_all},
        'totals': {
            'kg': round(sum(v['totalKg'] for v in vegs), 1),
            'spend': round(sum(v['totalSpend'] for v in vegs), 1),
            'txns': sum(v['nTxn'] for v in vegs),
            'nVeg': len(vegs),
        },
    }
    json.dump(payload, open(args.out, 'w', encoding='utf-8'),
              ensure_ascii=False, separators=(',', ':'))
    print(f"\nנכתב {args.out} | ירקות: {len(vegs)} | חודשים: {len(months)} | "
          f'סה"כ {payload["totals"]["kg"]:.0f} ק"ג · ₪{payload["totals"]["spend"]:.0f} | '
          f"החזרות: {len(ret_all)} ({payload['returns']['totalKg']} ק\"ג)")
    if not ok:
        print("⚠️  אזהרה: בדיקת ההצטלבות נכשלה לפחות לירק אחד — לבדוק לפני קליטה.")
        sys.exit(2)


if __name__ == '__main__':
    main()
