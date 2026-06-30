#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
reorder_analysis.py — אבחון קצב-הזמנה (reorder cadence) פר לקוח×מוצר מקובץ הלקוחות.

הרעיון: יש מוצרים שלקוח קונה ב"פיק" ואז נגמר לו אחרי שבועיים-שלושה ומזמין שוב.
מודל החציון (8 שבועות) מחליק את הפיקים ולכן מפספס אותם — חוזה נמוך בשבוע הפיק
וגבוה בשבוע החוסר, ושתי השגיאות נספרות ב-WAPE. כאן מזהים אילו מוצרים מתנהגים כך,
לפי הביקוש *לסירוגין* ברמת הלקוח (שלא נראה בנתוני הייצור המצרפיים).

זהו שלב 1 — אבחון בלבד (לא נוגע במודל). הפלט מזין את ההחלטה האם לבנות שכבת
ביקוש-לסירוגין בחיזוי (שלב 2: פרוטוטייפ + בקטסט חוץ-מדגמי).

מקור: אותו קובץ C_xxxx.xlsx של דשבורד הלקוחות (פורמט ציר+פירוט), עם אותו פרסר
מדויק של build_customers (אריזות מכירה נטו; לא מעמודת המשקל הגולמית).

הרצה:  python3 reorder_analysis.py C_1_26.xlsx
       python3 reorder_analysis.py C_1_26.xlsx --min-gap 2 --top 50
"""
import sys
import argparse
import statistics
from collections import defaultdict

import build_customers as bc   # מנצל את אותו פרסר מדויק של דשבורד הלקוחות


def analyse(path, min_gap=2, top=50, min_buys=3):
    rows = bc.load_rows(path)
    weeks = sorted({r["week"] for r in rows})

    # אריזות-מכירה נטו פר (לקוח, מוצר, שבוע) — מתעלמים מזיכויים (pkg<=0)
    series = defaultdict(lambda: defaultdict(float))   # (acc, name) -> week -> pkg
    prod_total = defaultdict(float)
    for r in rows:
        pkg, _money = bc.dirs(r)        # אריזות נטו: מכירה חיובית, זיכוי שלילי
        if pkg <= 0:
            continue
        series[(r["acc"], r["name"])][r["week"]] += pkg
        prod_total[r["name"]] += pkg

    # קצב פר (לקוח, מוצר): שבועות-קנייה, מרווחים, סיווג מחזורי/לסירוגין
    prod = defaultdict(lambda: dict(vol=0.0, inter_vol=0.0, gaps=[], pairs=0, inter_pairs=0))
    for (acc, name), wk_map in series.items():
        buyw = sorted(w for w, v in wk_map.items() if v > 0)
        vol = sum(wk_map.values())
        p = prod[name]
        p["vol"] += vol
        p["pairs"] += 1
        if len(buyw) >= min_buys:
            gaps = [b - a for a, b in zip(buyw, buyw[1:])]
            med = statistics.median(gaps)
            p["gaps"].extend(gaps)
            if med >= min_gap:          # קונה כל ~min_gap+ שבועות → לסירוגין
                p["inter_vol"] += vol
                p["inter_pairs"] += 1

    total_vol = sum(prod_total.values()) or 1.0
    ranked = sorted(prod_total, key=lambda n: -prod_total[n])[:top]

    print(f"קובץ: {path}")
    print(f"שבועות: {weeks[0]}–{weeks[-1]} ({len(weeks)}) | מוצרים: {len(prod_total)} | "
          f"זוגות לקוח×מוצר: {len(series)}\n")
    print(f"{'מוצר':<24}{'נפח%':>6}{'נפח-לסירוגין%':>14}{'מרווח חציוני':>13}  דגל")
    print("-" * 66)
    flagged = []
    for name in ranked:
        p = prod[name]
        share = 100 * p["vol"] / total_vol
        inter = 100 * p["inter_vol"] / p["vol"] if p["vol"] else 0.0
        med = statistics.median(p["gaps"]) if p["gaps"] else 0.0
        # דגל לפי חלק הנפח מלקוחות לסירוגין (האות המשמעותי); המרווח החציוני
        # מאגד בין לקוחות ולכן רק אינפורמטיבי.
        flag = "⟵ מועמד" if inter >= 50 else ""
        if flag:
            flagged.append(name)
        print(f"{name[:24]:<24}{share:5.1f}%{inter:12.0f}%{med:11.1f}   {flag}")

    inter_share = 100 * sum(prod[n]["inter_vol"] for n in ranked) / sum(prod[n]["vol"] for n in ranked)
    print("-" * 66)
    print(f"\nבטופ-{top}: {len(flagged)} מוצרים מועמדים לשכבת ביקוש-לסירוגין "
          f"(≥50% מהנפח מלקוחות שקונים כל {min_gap}+ שבועות).")
    print(f"חלק הנפח ה'לסירוגין' בטופ-{top}: {inter_share:.0f}%")
    if flagged:
        print("מועמדים:", ", ".join(flagged[:15]))


def backtest(path, min_gap=2, train_end=17):
    """בקטסט: האם תחזית מחזורית (פאזה+מרווח) מנצחת מודל שטוח (ממוצע שבועי) על
    הזוגות המחזוריים? תוצאה (שבועות 1-26, אימון≤17): המחזורי *גרוע יותר* — המרווחים
    רועשים מכדי לנבא תזמון, והזוגות המחזוריים הם <5% מהנפח. מסקנה: לא לבנות שכבה כזו."""
    rows = bc.load_rows(path)
    weeks = sorted({r["week"] for r in rows})
    train = [w for w in weeks if w <= train_end]
    test = [w for w in weeks if w > train_end]
    series = defaultdict(lambda: defaultdict(float))
    for r in rows:
        pkg, _ = bc.dirs(r)
        if pkg > 0:
            series[(r["acc"], r["name"])][r["week"]] += pkg
    n_flat = d_flat = n_cad = d_cad = 0.0
    npairs = 0
    vol_cyc = vol_all = 0.0
    for (acc, name), wk in series.items():
        vol_all += sum(wk.values())
        tb = sorted(w for w in train if wk.get(w, 0) > 0)
        if len(tb) < 4:
            continue
        gaps = [y - x for x, y in zip(tb, tb[1:])]
        g = int(round(statistics.median(gaps)))
        if g < min_gap:
            continue
        npairs += 1
        vol_cyc += sum(wk.values())
        m = statistics.median([wk[w] for w in tb])
        mean_wk = sum(wk[w] for w in train) / len(train)
        due = set()
        w = tb[-1] + g
        while w <= test[-1]:
            if w in test:
                due.add(w)
            w += g
        for w in test:
            act = wk.get(w, 0)
            n_flat += abs(mean_wk - act); d_flat += act
            n_cad += abs((m if w in due else 0.0) - act); d_cad += act
    f = lambda n, d: 100 * n / d if d else 0
    print(f"זוגות מחזוריים (מרווח≥{min_gap}, ≥4 קניות): {npairs} | "
          f"נפחם {vol_cyc:.0f}/{vol_all:.0f} ({100*vol_cyc/vol_all:.1f}%)")
    print(f"WAPE על הזוגות המחזוריים (test {test[0]}-{test[-1]}):")
    print(f"   מודל שטוח (ממוצע שבועי):  {f(n_flat,d_flat):.1f}%")
    print(f"   מודל מחזורי (פאזה+מרווח): {f(n_cad,d_cad):.1f}%")
    print("מסקנה: התזמון רועש מכדי לנבא; השכבה המחזורית מזיקה ולא שווה לבנות.")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("path", nargs="?", default="C_1_26.xlsx", help="קובץ C_xxxx.xlsx")
    ap.add_argument("--backtest", action="store_true", help="בקטסט מחזורי מול שטוח")
    ap.add_argument("--min-gap", type=int, default=2, help="מרווח חציוני מינימלי שנחשב 'לסירוגין' (שבועות)")
    ap.add_argument("--top", type=int, default=50, help="כמה מוצרים מובילים לבדוק")
    ap.add_argument("--min-buys", type=int, default=3, help="מינ' קניות כדי לחשב מרווח")
    a = ap.parse_args()
    if a.backtest:
        backtest(a.path, a.min_gap)
    else:
        analyse(a.path, a.min_gap, a.top, a.min_buys)


if __name__ == "__main__":
    main()
