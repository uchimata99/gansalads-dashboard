/*
 * print_plan_pdf.js — מפיק PDF של תוכנית הייצור היומית (לרוחב, צבעוני, עמוד אחד).
 * נועד למיכל שמדפיסה מאייפון (ספארי לא מכבד @page landscape / צבע).
 *
 * מה הוא עושה: קורא את MY_PLAN ישירות מ-forecast.html (מקור אמת יחיד),
 * קורא נתוני חיזוי מ-DATA.json, מחשב את אותה תוכנית כמו forecast.html
 * (התניות כגרעין קבוע + חיזוי מאוזן לשאר), ומייצר plan.html.
 *
 * הרצה:
 *   1) למשוך את לשונית DATA מהגיליון "נתוני דשבורד ייצור" ל-DATA.json
 *      (אובייקט עם PROD_DETAIL / PROD_UNI / WEEKS — דרך Google Drive MCP / Sheets).
 *   2) node tools/print_plan_pdf.js <forecast.html> <DATA.json> <outDir>
 *   3) chrome --headless --no-sandbox --disable-gpu --no-pdf-header-footer \
 *        --print-to-pdf="<outDir>/תוכנית_ייצור.pdf" "file://<outDir>/plan.html"
 *      (Chromium ב: /opt/pw-browsers/chromium-<ver>/chrome-linux/chrome)
 *
 * להוספת התניה: עורכים MY_PLAN ב-forecast.html בלבד — הסקריפט קורא משם.
 */
const fs = require('fs');
const FCSRC = process.argv[2] || 'forecast.html';
const DATA  = process.argv[3] || 'DATA.json';
const OUT   = process.argv[4] || '.';

// --- MY_PLAN מתוך forecast.html (מקור אמת יחיד) ---
const html = fs.readFileSync(FCSRC, 'utf8');
const m = html.match(/const MY_PLAN\s*=\s*(\[[\s\S]*?\]);/);
if (!m) { console.error('MY_PLAN not found in ' + FCSRC); process.exit(1); }
const MY_PLAN = new Function('return ' + m[1])();
const MY_BASES = new Set(MY_PLAN.map(p => p.b));

const o = JSON.parse(fs.readFileSync(DATA, 'utf8'));
const PD = o.PROD_DETAIL, PU = o.PROD_UNI || {};
const weeks = o.WEEKS.map(w => w.w).sort((a, b) => a - b).slice(-4);
const PLAN_DAYS = ['ראשון', 'שני', 'שלישי', 'רביעי', 'חמישי'];
const num = x => { const n = parseFloat(x); return isNaN(n) ? 0 : n; };
const fmt = n => Math.round(n).toLocaleString('en-US');

// חיזוי פר-פריט: ממוצע נע 4 שבועות (ק"ג, אינדקס 1)
const fc = {};
for (const name in PD) { let t = 0, any = false; weeks.forEach(w => { const dd = PD[name][w]; if (dd) for (const di in dd) { const i = +di; if (i >= 0 && i < 6) { t += num(dd[di][1]); any = true; } } }); if (any && t / 4 > 0) fc[name] = t / 4; }
const baseOf = n => PU[n] || String(n).replace(/\s*\d+(\.\d+)?\s*(ק"ג|קג|גרם|גר|ג')?\s*$/, '').trim();
const baseFc = {}; Object.keys(fc).forEach(n => { const b = baseOf(n); baseFc[b] = (baseFc[b] || 0) + fc[n]; });
const planType = n => String(n).replace(/\d+.*$/, '').trim().split(/\s+/)[0] || 'אחר';
const isHT = n => (n.indexOf('חומוס') >= 0 || n.indexOf('טחינה') >= 0) && n.indexOf('גרגרי') < 0 && n.indexOf('חציל') < 0;

// דגל שבועי
const flagOf = {}; MY_PLAN.forEach(p => { const wp = PLAN_DAYS.reduce((a, d) => a + (p.d[d] || 0), 0); const f = Math.round(baseFc[p.b] || 0); const dev = Math.round(wp - f); if (Math.abs(dev) > 20) flagOf[p.n] = { dev }; });

// גרעין קבוע
const load = {}, bins = {}; PLAN_DAYS.forEach(d => { load[d] = 0; bins[d] = []; });
let condKg = 0; MY_PLAN.forEach(p => PLAN_DAYS.forEach(d => { const kg = p.d[d]; if (kg > 0) { bins[d].push({ type: planType(p.n), p: { name: p.n, kg }, fixed: true }); load[d] += kg; condKg += kg; } }));
// חיזוי לשאר
const prods = [], excl = []; Object.keys(fc).forEach(n => { if (MY_BASES.has(baseOf(n))) return; (isHT(n) ? excl : prods).push({ name: n, kg: fc[n] }); });
prods.sort((a, b) => b.kg - a.kg);
const autoKg = prods.reduce((a, p) => a + p.kg, 0); const target = (condKg + autoKg) / 5;
const types = {}; prods.forEach(p => { const t = planType(p.name); (types[t] = types[t] || []).push(p); });
const leastDay = ex => { let b = null; PLAN_DAYS.forEach(d => { if (d === ex) return; if (b === null || load[d] < load[b]) b = d; }); return b; };
const place = (p, t, d) => { load[d] += p.kg; bins[d].push({ type: t, p }); };
const units = Object.keys(types).map(t => ({ t, items: types[t], kg: types[t].reduce((a, p) => a + p.kg, 0) })).sort((a, b) => b.kg - a.kg);
units.forEach(u => { if (u.kg <= target * 1.15 || u.items.length === 1) { const d = leastDay(null); u.items.forEach(p => place(p, u.t, d)); } else { const d1 = leastDay(null), d2 = leastDay(d1); let a = 0, b = 0; u.items.slice().sort((x, y) => y.kg - x.kg).forEach(p => { if (a <= b) { place(p, u.t, d1); a += p.kg; } else { place(p, u.t, d2); b += p.kg; } }); } });

const COLORS = ['#e11d48', '#7c3aed', '#2563eb', '#16a34a', '#d97706'];
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
let colsHtml = '';
PLAN_DAYS.forEach((d, i) => {
  const byType = {}; bins[d].forEach(x => { const g = byType[x.type] = byType[x.type] || { type: x.type, kg: 0, items: [] }; g.kg += x.p.kg; g.items.push({ name: x.p.name, kg: x.p.kg, fixed: !!x.fixed }); });
  const groups = Object.values(byType).sort((a, b) => b.kg - a.kg);
  let body = '';
  groups.forEach(g => {
    let items = '';
    g.items.forEach(it => {
      let mark = '';
      if (it.fixed) { const f = flagOf[it.name]; if (f) { const over = f.dev > 0; mark = '<span style="color:' + (over ? '#16a34a' : '#dc2626') + ';font-weight:700">' + (over ? '▲+' : '▼−') + fmt(Math.abs(f.dev)) + '</span> '; } else { mark = '<span style="color:#16a34a">●</span> '; } }
      const nm = it.fixed ? '<b>' + esc(it.name) + '</b>' : esc(it.name);
      items += '<div class="it"><span>' + mark + nm + '</span><span class="kg">' + fmt(it.kg) + '</span></div>';
    });
    body += '<div class="pg"><div class="pgh"><span>' + esc(g.type) + '</span><span>' + fmt(g.kg) + '</span></div>' + items + '</div>';
  });
  colsHtml += '<div class="col"><div class="hd" style="color:' + COLORS[i] + ';border-color:' + COLORS[i] + '"><span class="dn">' + d + '</span><span class="dk">' + fmt(load[d]) + ' ק"ג</span></div><div class="bd">' + body + '</div></div>';
});
const total = condKg + autoKg;
const html2 = '<!doctype html><html dir="rtl" lang="he"><head><meta charset="utf-8"><style>'
  + '@page{size:A4 landscape;margin:5mm}*{box-sizing:border-box;margin:0;padding:0;-webkit-print-color-adjust:exact;print-color-adjust:exact}'
  + 'body{font-family:"DejaVu Sans",Arial,sans-serif;color:#1e293b}'
  + 'h1{font-size:14px;color:#0d5f2c} .sub{font-size:8.5px;color:#475569;margin:2px 0 6px}'
  + '.grid{display:grid;grid-template-columns:repeat(5,1fr);gap:3px}'
  + '.col{border:1px solid #cbd5e1;border-radius:6px;overflow:hidden}'
  + '.hd{display:flex;justify-content:space-between;align-items:baseline;padding:3px 6px;background:#fff;border-bottom:2.5px solid}'
  + '.dn{font-weight:800;font-size:10px}.dk{font-weight:700;font-size:8px}'
  + '.bd{padding:2px 4px}.pg{margin-bottom:2px}.pgh{display:flex;justify-content:space-between;font-weight:700;font-size:6.8px;color:#1a3a5c;border-bottom:1px dashed #e2e8f0;padding-bottom:1px;margin-bottom:1px}'
  + '.it{display:flex;justify-content:space-between;gap:4px;font-size:6.6px;line-height:1.12;color:#334155;padding:.5px 0}.it .kg{color:#94a3b8;white-space:nowrap}.it span:first-child{overflow-wrap:anywhere}'
  + '.lg{font-size:8px;color:#475569;margin-top:5px}</style></head><body>'
  + '<h1>סלטי גן — תוכנית ייצור שבועית</h1>'
  + '<div class="sub">סה"כ <b>' + fmt(total) + ' ק"ג</b> · מתוכם <b>' + fmt(condKg) + ' ק"ג</b> לפי ההתניות שלך (' + MY_PLAN.length + ' מוצרים) · השאר חיזוי מאוזן · חומוס/טחינה (' + fmt(excl.reduce((a, p) => a + p.kg, 0)) + ' ק"ג) בנפרד</div>'
  + '<div class="grid">' + colsHtml + '</div>'
  + '<div class="lg"><b>●</b> מוצר שלך בתחום · <b style="color:#16a34a">▲</b> עודף · <b style="color:#dc2626">▼</b> חוסר — מול החיזוי השבועי (מעל 20 ק"ג)</div>'
  + '</body></html>';
fs.writeFileSync(OUT + '/plan.html', html2);
console.log('plan.html written → ' + OUT + '/plan.html | total=' + Math.round(total) + ' condKg=' + condKg + ' flags=' + Object.keys(flagOf).length);
