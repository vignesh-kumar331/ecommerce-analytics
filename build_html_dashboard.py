import csv
import json
from datetime import datetime, timedelta
from collections import defaultdict

CSV_PATH = "orders.csv"
HTML_PATH = "index.html"

# ---------- Load & analyze ----------
rows = []
with open(CSV_PATH, newline="", encoding="utf-8") as f:
    for r in csv.DictReader(f):
        r["dt"] = datetime.strptime(r["order_date"], "%m/%d/%Y")
        r["amount_f"] = float(r["order_amount"])
        rows.append(r)

total_revenue = sum(r["amount_f"] for r in rows)
total_orders = len(rows)
aov = total_revenue / total_orders

monthly = defaultdict(lambda: {"revenue": 0.0, "orders": 0})
for r in rows:
    m = r["dt"].strftime("%Y-%m")
    monthly[m]["revenue"] += r["amount_f"]
    monthly[m]["orders"] += 1
months_sorted = sorted(monthly)

cat_rev = defaultdict(float)
cat_ord = defaultdict(int)
for r in rows:
    cat_rev[r["product_category"]] += r["amount_f"]
    cat_ord[r["product_category"]] += 1
cats_sorted = sorted(cat_rev, key=lambda c: cat_rev[c], reverse=True)

cust_orders = defaultdict(int)
cust_last = {}
for r in rows:
    cid = r["customer_id"]
    cust_orders[cid] += 1
    if cid not in cust_last or r["dt"] > cust_last[cid]:
        cust_last[cid] = r["dt"]

analysis_date = max(r["dt"] for r in rows)
cutoff = analysis_date - timedelta(days=90)
repeat_n = sum(1 for n in cust_orders.values() if n >= 2)
one_time_n = len(cust_orders) - repeat_n
churned_n = sum(1 for d in cust_last.values() if d < cutoff)
active_n = len(cust_orders) - churned_n
repeat_rate = repeat_n / len(cust_orders)
churn_pct = churned_n / len(cust_orders)

best_m = max(months_sorted, key=lambda m: monthly[m]["revenue"])
worst_m = min(months_sorted, key=lambda m: monthly[m]["revenue"])
top_cat, bottom_cat = cats_sorted[0], cats_sorted[-1]

def month_label(m):
    names = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    y, mo = m.split("-")
    return f"{names[int(mo)-1]} {y}"

payload = {
    "kpis": {
        "totalRevenue": round(total_revenue),
        "totalOrders": total_orders,
        "aov": round(aov, 2),
        "uniqueCustomers": len(cust_orders),
        "repeatRate": round(repeat_rate * 100, 1),
        "churned": churned_n,
        "churnPct": round(churn_pct * 100, 1),
    },
    "months": [month_label(m) for m in months_sorted],
    "monthlyRevenue": [round(monthly[m]["revenue"]) for m in months_sorted],
    "monthlyOrders": [monthly[m]["orders"] for m in months_sorted],
    "categories": cats_sorted,
    "categoryRevenue": [round(cat_rev[c]) for c in cats_sorted],
    "customerStatus": {"Active": active_n, "Churned": churned_n},
}

html = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>E-Commerce Sales &amp; Customer Retention Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"></script>
<style>
  :root {
    --navy:#1f3864; --blue:#2e75b6; --green:#548235; --purple:#7030a0;
    --red:#c00000; --bg:#f4f6fa; --card:#ffffff; --text:#243447;
  }
  * { box-sizing:border-box; margin:0; padding:0; font-family:'Segoe UI',Arial,sans-serif; }
  body { background:var(--bg); color:var(--text); padding:28px; }
  .wrap { max-width:1200px; margin:0 auto; }
  header { text-align:center; margin-bottom:26px; }
  header h1 { color:var(--navy); font-size:26px; }
  header p { color:#7a8699; margin-top:6px; font-size:14px; }
  .grid-kpi { display:grid; grid-template-columns:repeat(auto-fit,minmax(170px,1fr)); gap:14px; margin-bottom:22px; }
  .kpi { background:var(--card); border-radius:12px; padding:16px; text-align:center;
         box-shadow:0 1px 4px rgba(20,40,80,.08); border-top:4px solid var(--navy); }
  .kpi .label { font-size:11px; letter-spacing:.06em; font-weight:700; color:#8a94a6; }
  .kpi .value { font-size:26px; font-weight:800; color:var(--navy); margin-top:6px; }
  .kpi.rev  { border-color:var(--navy); }
  .kpi.ord  { border-color:var(--blue); }   .kpi.ord .value { color:var(--blue); }
  .kpi.aov  { border-color:var(--green); }  .kpi.aov .value { color:var(--green); }
  .kpi.rep  { border-color:var(--purple);}  .kpi.rep .value { color:var(--purple); }
  .kpi.chr  { border-color:var(--red); }    .kpi.chr .value { color:var(--red); }
  .grid-charts { display:grid; grid-template-columns:1fr 1fr; gap:18px; }
  @media (max-width:860px){ .grid-charts{ grid-template-columns:1fr; } }
  .card { background:var(--card); border-radius:12px; padding:18px; box-shadow:0 1px 4px rgba(20,40,80,.08); }
  .card h2 { font-size:14px; color:var(--navy); margin-bottom:10px; }
  .chart-box { position:relative; height:300px; }
  .insights { background:var(--card); border-radius:12px; padding:20px; margin-top:18px;
              box-shadow:0 1px 4px rgba(20,40,80,.08); }
  .insights h2 { color:var(--navy); font-size:15px; margin-bottom:10px; }
  .insights li { margin:7px 0 7px 18px; font-size:14px; line-height:1.5; }
  footer { text-align:center; color:#9aa5b5; font-size:12px; margin-top:24px; }
</style>
</head>
<body>
<div class="wrap">
  <header>
    <h1>E-Commerce Sales &amp; Customer Retention Dashboard</h1>
    <p>Data period: Jul 2024 – Dec 2025 • __ORDERS__ orders • __CUSTOMERS__ customers</p>
  </header>

  <div class="grid-kpi">
    <div class="kpi rev"><div class="label">TOTAL REVENUE</div><div class="value">$__REV__</div></div>
    <div class="kpi ord"><div class="label">TOTAL ORDERS</div><div class="value">__NORDERS__</div></div>
    <div class="kpi aov"><div class="label">AVG ORDER VALUE</div><div class="value">$__AOV__</div></div>
    <div class="kpi rep"><div class="label">REPEAT PURCHASE RATE</div><div class="value">__REP__%</div></div>
    <div class="kpi chr"><div class="label">CHURNED (90d+)</div><div class="value">__CHURNED__</div></div>
    <div class="kpi"><div class="label">UNIQUE CUSTOMERS</div><div class="value">__CUSTOMERS__</div></div>
  </div>

  <div class="grid-charts">
    <div class="card"><h2>Monthly Revenue Trend</h2><div class="chart-box"><canvas id="revChart"></canvas></div></div>
    <div class="card"><h2>Orders per Month</h2><div class="chart-box"><canvas id="ordChart"></canvas></div></div>
    <div class="card"><h2>Revenue by Product Category</h2><div class="chart-box"><canvas id="catChart"></canvas></div></div>
    <div class="card"><h2>Customer Status (Active vs Churned)</h2><div class="chart-box"><canvas id="pieChart"></canvas></div></div>
  </div>

  <div class="insights">
    <h2>Key Insights</h2>
    <ul>
      <li>Best month: <b>__BESTM__ ($__BESTV__)</b>; weakest month: <b>__WORSTM__ ($__WORSTV__)</b>.</li>
      <li><b>__TOPCAT__</b> leads all categories with $__TOPVAL__ (__TOPPCT__% of revenue); <b>__BOTCAT__</b> is lowest at $__BOTVAL__.</li>
      <li>__REP_N__ of __CUST_TOT__ customers ordered more than once — repeat purchase rate <b>__REP__%</b>.</li>
      <li><b>__CHURNED__ customers (__CHURNPCT__%)</b> haven't ordered in 90+ days — ideal target for a win-back campaign.</li>
      <li>Average order value is <b>$__AOV__</b> across __ORDERS__ orders.</li>
    </ul>
  </div>

  <footer>E-Commerce Analytics Project • Built with Chart.js • Data: orders.csv</footer>
</div>

<script>
const D = __DATA__;
const money = v => '$' + v.toLocaleString('en-US');
const NAVY='#1f3864', BLUE='#2e75b6', GREEN='#548235', RED='#c00000';

new Chart(document.getElementById('revChart'), {
  type:'line',
  data:{ labels:D.months,
    datasets:[{ label:'Revenue ($)', data:D.monthlyRevenue, borderColor:NAVY,
      backgroundColor:'rgba(31,56,100,.12)', fill:true, tension:.25,
      pointRadius:3, pointBackgroundColor:NAVY }]},
  options:{ maintainAspectRatio:false,
    plugins:{ tooltip:{ callbacks:{ label:c=>money(c.parsed.y) } }, legend:{display:false} },
    scales:{ y:{ ticks:{ callback:v=>'$'+(v/1000)+'k' } } } }
});

new Chart(document.getElementById('ordChart'), {
  type:'bar',
  data:{ labels:D.months,
    datasets:[{ label:'Orders', data:D.monthlyOrders, backgroundColor:BLUE, borderRadius:4 }]},
  options:{ maintainAspectRatio:false, plugins:{ legend:{display:false} },
    scales:{ y:{ beginAtZero:true } } }
});

new Chart(document.getElementById('catChart'), {
  type:'bar',
  data:{ labels:D.categories,
    datasets:[{ label:'Revenue ($)', data:D.categoryRevenue,
      backgroundColor:[NAVY,BLUE,GREEN,'#b58b00','#a33b3b'], borderRadius:4 }]},
  options:{ indexAxis:'y', maintainAspectRatio:false,
    plugins:{ legend:{display:false}, tooltip:{ callbacks:{ label:c=>money(c.parsed.x) } } },
    scales:{ x:{ ticks:{ callback:v=>'$'+(v/1000)+'k' } } } }
});

new Chart(document.getElementById('pieChart'), {
  type:'doughnut',
  data:{ labels:Object.keys(D.customerStatus),
    datasets:[{ data:Object.values(D.customerStatus), backgroundColor:[GREEN,RED], borderWidth:2 }]},
  options:{ maintainAspectRatio:false, cutout:'55%',
    plugins:{ legend:{position:'bottom'},
      tooltip:{ callbacks:{ label:c=>c.label+': '+c.parsed.toLocaleString()+' ('+ (c.parsed/(D.kpis.uniqueCustomers)*100).toFixed(1)+'%)' } } } }
});
</script>
</body>
</html>
"""

def fmt_money(x):
    return f"{x:,.0f}"

html = (html
    .replace("__DATA__", json.dumps(payload))
    .replace("__REV__", fmt_money(total_revenue))
    .replace("__NORDERS__", f"{total_orders:,}")
    .replace("__AOV__", f"{aov:,.2f}")
    .replace("__REP__", f"{repeat_rate*100:.1f}")
    .replace("__CHURNED__", f"{churned_n:,}")
    .replace("__CHURNPCT__", f"{churn_pct*100:.1f}")
    .replace("__CUSTOMERS__", f"{len(cust_orders):,}")
    .replace("__ORDERS__", f"{total_orders:,}")
    .replace("__BESTM__", month_label(best_m))
    .replace("__BESTV__", fmt_money(monthly[best_m]["revenue"]))
    .replace("__WORSTM__", month_label(worst_m))
    .replace("__WORSTV__", fmt_money(monthly[worst_m]["revenue"]))
    .replace("__TOPCAT__", top_cat)
    .replace("__TOPVAL__", fmt_money(cat_rev[top_cat]))
    .replace("__TOPPCT__", f"{cat_rev[top_cat]/total_revenue*100:.0f}")
    .replace("__BOTCAT__", bottom_cat)
    .replace("__BOTVAL__", fmt_money(cat_rev[bottom_cat]))
    .replace("__REP_N__", f"{repeat_n:,}")
    .replace("__CUST_TOT__", f"{len(cust_orders):,}")
)

with open(HTML_PATH, "w", encoding="utf-8") as f:
    f.write(html)
print(f"Saved {HTML_PATH} ({len(html):,} bytes)")
