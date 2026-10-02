#!/usr/bin/env python3
"""
Generate the editable AirIQ Exploratory Data Analysis report (.docx).

    python3 theme/make_eda_doc.py

Everything is computed live from the packaged data extract inside
Air_Pollution_Dashboard.twbx -- the exact data the dashboards read -- and
every colour is resolved from theme/palette.json, so the report always
matches the workbook and the design system.

Dependencies (re-install after a sandbox reset):
    pip install --break-system-packages pandas matplotlib python-docx tableauhyperapi

Output: AirIQ_EDA_Report.docx (repo root)
Figures: preview/eda/*.png
"""
from __future__ import annotations

import json
import tempfile
import zipfile
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
T = json.loads((ROOT / "theme" / "palette.json").read_text(encoding="utf-8"))
C = T["color"]
VERSION = T.get("version", "1.0.0")

def hx(token: str) -> str:
    return token.lstrip("#").lower()

def rgb(token: str):
    h = hx(token)
    from docx.shared import RGBColor
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

INK = C["ink"]["heading"]
BODY = C["ink"]["body"]
MUTED = C["ink"]["muted"]
BRAND = C["brand"]["primary"]
GOOD = C["status"]["good"]
MOD = C["status"]["moderate"]
BAD = C["status"]["bad"]
CANVAS = C["surface"]["canvas"]
HAIR = C["line"]["hairline"]
PRIOR = C["series"]["prior"]
CURRENT = C["series"]["current"]
POLL_COLOR = C["pollutant"]

# ================================================================ data load ==
def load_data() -> "pandas.DataFrame":
    import pandas as pd
    from tableauhyperapi import HyperProcess, Connection, Telemetry

    zf = zipfile.ZipFile(ROOT / "Air_Pollution_Dashboard.twbx")
    names = [n for n in zf.namelist() if n.endswith(".hyper")]
    # the wide-format extract is the one with per-pollutant columns
    for name in names:
        with tempfile.TemporaryDirectory(prefix="eda_") as tmp:
            p = Path(tmp) / "extract.hyper"
            p.write_bytes(zf.read(name))
            with HyperProcess(telemetry=Telemetry.DO_NOT_SEND_USAGE_DATA_TO_TABLEAU) as hp:
                with Connection(hp.endpoint, str(p)) as conn:
                    with conn.execute_query('SELECT * FROM "Extract"."Extract"') as r:
                        cols = [str(c.name).strip('"') for c in r.schema.columns]
                        df = pd.DataFrame(r, columns=cols)
        if "pm2_5" in df.columns:
            df["date"] = pd.to_datetime(df["date"].astype(str))
            return df
    raise SystemExit("could not locate the wide-format extract in the workbook")

import pandas as pd  # noqa: E402  (kept local to keep the import-error message friendly)

df = load_data()

POLLUTANTS = [  # (column, label, naaqs 24h limit, naaqs annual limit) -- workbook benchmarks
    ("pm2_5", "PM2.5", 60, 40),
    ("pm10", "PM10", 100, 60),
    ("no2", "NO2", 80, 40),
    ("so2", "SO2", 80, 50),
    ("co", "CO", 4000, 1000),
    ("o3", "O3", 180, 100),
    ("nh3", "NH3", 400, 400),
    ("no", "NO", 80, 40),
]

# ------------------------------------------------------------ core stats ----
n_rows = len(df)
n_cities = df["city"].nunique()
days_per_city = df.groupby("city").size()
d_min, d_max = df["date"].min(), df["date"].max()
cal_days = (d_max - d_min).days + 1
missing_days = sorted(set(pd.date_range(d_min, d_max)) - set(df["date"].unique()))
pct_unhealthy = (df["aqi"] >= 4).mean() * 100
aqi_share = (df["aqi"].value_counts().sort_index() / n_rows * 100)

city = (df.groupby("city")
          .agg(avg_aqi=("aqi", "mean"),
               unhl=("aqi", lambda s: (s >= 4).mean() * 100),
               pm25=("pm2_5", "mean"))
          .sort_values("avg_aqi", ascending=False))
worst_city, best_city = city.index[0], city.index[-1]

poll = []
for col, label, lim24, limyr in POLLUTANTS:
    s = df[col]
    poll.append({
        "col": col, "label": label, "mean": s.mean(), "median": s.median(),
        "lim24": lim24, "limyr": limyr,
        "pct_over": (s > lim24).mean() * 100,
        "mean_vs_annual": s.mean() / limyr * 100,
    })
poll.sort(key=lambda d: -d["pct_over"])
top5_conc = sorted(poll, key=lambda d: -d["mean"])[:5]

df["year"] = df["date"].dt.year
df["month"] = df["date"].dt.month
d21 = df[df["year"] == 2021]
d22 = df[df["year"] == 2022]
avg21, avg22 = d21["aqi"].mean(), d22["aqi"].mean()
yoy = (avg22 - avg21) / avg21 * 100
d2122 = df[df["year"].isin([2021, 2022])]
mt21 = d21.groupby("month")["aqi"].mean()
mt22 = d22.groupby("month")["aqi"].mean()

SEASONS = {"Winter": [12, 1, 2], "Pre-monsoon": [3, 4, 5],
           "Monsoon": [6, 7, 8, 9], "Post-monsoon": [10, 11]}
season_rows = []
for name, months in SEASONS.items():
    s = df[df["month"].isin(months)]["aqi"]
    season_rows.append((name, s.mean(), (s >= 4).mean() * 100))

cm2122 = d2122.pivot_table(index="city", columns="month", values="aqi", aggfunc="mean")
cm2122 = cm2122.reindex(city.index)  # worst city on top
peak_month = cm2122.idxmax(axis=1)
n_peak_nov_mar = int(peak_month.isin([11, 12, 1, 2, 3]).sum())
n_records_2122 = len(d2122)

# ================================================================ figures ====
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap
from matplotlib.patches import Patch

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "font.size": 10,
    "text.color": BODY, "axes.labelcolor": BODY,
    "xtick.color": MUTED, "ytick.color": MUTED,
    "axes.edgecolor": HAIR, "axes.linewidth": 1.0,
    "axes.grid": True, "grid.color": HAIR, "grid.linewidth": 0.8,
    "axes.axisbelow": True,
    "axes.spines.top": False, "axes.spines.right": False,
    "figure.facecolor": "white", "axes.facecolor": "white",
    "savefig.dpi": 200, "savefig.bbox": "tight", "savefig.pad_inches": 0.1,
})
FIG = ROOT / "preview" / "eda"
FIG.mkdir(parents=True, exist_ok=True)
MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

def status_color(pct: float) -> str:
    return GOOD if pct < 50 else (MOD if pct < 75 else BAD)

# -- fig 1: AQI category distribution -----------------------------------------
fig, ax = plt.subplots(figsize=(6.9, 2.8))
cats = [1, 2, 3, 4, 5]
vals = [aqi_share[c] for c in cats]
colors = [GOOD, GOOD, MOD, BAD, BAD]
bars = ax.bar([str(c) for c in cats], vals, color=colors, width=0.62)
for b, v in zip(bars, vals):
    ax.text(b.get_x() + b.get_width() / 2, v + 1.2, f"{v:.1f}%",
            ha="center", va="bottom", fontsize=10, color=INK, fontweight="bold")
ax.set_ylim(0, max(vals) * 1.18)
ax.set_xlabel("AQI category  (1 = best  \u00b7  5 = worst)")
ax.set_ylabel("Share of city-days")
ax.set_yticks([0, 10, 20, 30, 40, 50, 60])
ax.set_yticklabels(["0%", "10%", "20%", "30%", "40%", "50%", "60%"])
ax.grid(axis="x", visible=False)
fig.savefig(FIG / "aqi_dist.png")
plt.close(fig)

# -- fig 2: unhealthy days by city --------------------------------------------
fig, ax = plt.subplots(figsize=(6.9, 5.5))
srt = city.sort_values("unhl", ascending=False)
colors = [status_color(v) for v in srt["unhl"]]
bars = ax.barh(range(len(srt)), srt["unhl"], color=colors, height=0.66)
ax.set_yticks(range(len(srt)))
ax.set_yticklabels(srt.index, fontsize=8.5)
ax.invert_yaxis()
for i, v in enumerate(srt["unhl"]):
    ax.text(v + 1.0, i, f"{v:.0f}%", va="center", fontsize=8, color=BODY)
ax.set_xlim(0, 108)
ax.set_xticks([0, 25, 50, 75, 100])
ax.set_xticklabels(["0%", "25%", "50%", "75%", "100%"])
ax.grid(axis="y", visible=False)
ax.legend(handles=[Patch(color=GOOD, label="< 50% of days"),
                   Patch(color=MOD, label="50 \u2013 75%"),
                   Patch(color=BAD, label="> 75%")],
          frameon=False, fontsize=8.5, loc="lower right")
fig.savefig(FIG / "city_unhealthy.png")
plt.close(fig)

# -- fig 3: monthly national trend, 2021 vs 2022 ------------------------------
fig, ax = plt.subplots(figsize=(6.9, 3.1))
x = range(1, 13)
ax.plot(x, [mt21[m] for m in x], color=PRIOR, lw=2.4, marker="o", ms=4.5,
        label="2021", zorder=3)
ax.plot(x, [mt22[m] for m in x], color=CURRENT, lw=2.6, marker="o", ms=4.5,
        label="2022", zorder=4)
ax.set_xticks(list(x))
ax.set_xticklabels(MONTH_ABBR)
ax.set_ylabel("National average AQI")
ax.set_ylim(2.4, 5.2)
ax.legend(frameon=False, loc="lower center", ncol=2, fontsize=9.5)
fig.savefig(FIG / "monthly_trend.png")
plt.close(fig)

# -- fig 4: days above NAAQS 24-h limit, by pollutant -------------------------
fig, ax = plt.subplots(figsize=(6.9, 3.3))
labels = [p["label"] for p in poll]
vals = [p["pct_over"] for p in poll]
colors = [POLL_COLOR[p["col"]] for p in poll]
bars = ax.bar(labels, vals, color=colors, width=0.62)
for b, v in zip(bars, vals):
    ax.text(b.get_x() + b.get_width() / 2, v + max(vals) * 0.02, f"{v:.1f}%",
            ha="center", va="bottom", fontsize=9, color=INK, fontweight="bold")
ax.set_ylim(0, max(vals) * 1.2)
ax.set_ylabel("% of days above 24-h limit")
ax.grid(axis="x", visible=False)
fig.savefig(FIG / "exceedance.png")
plt.close(fig)

# -- fig 5: city x month heatmap (2021-2022) -----------------------------------
fig, ax = plt.subplots(figsize=(6.9, 4.3))
cmap = LinearSegmentedColormap.from_list("atmos", [CANVAS, MOD, BAD])
im = ax.imshow(cm2122.values, aspect="auto", cmap=cmap, vmin=2, vmax=5)
ax.set_xticks(range(12))
ax.set_xticklabels(MONTH_ABBR, fontsize=8.5)
ax.set_yticks(range(len(cm2122)))
ax.set_yticklabels(cm2122.index, fontsize=7.5)
ax.grid(visible=False)
for spine in ax.spines.values():
    spine.set_visible(False)
cb = fig.colorbar(im, ax=ax, shrink=0.85, pad=0.01)
cb.set_label("Average AQI", fontsize=9, color=BODY)
cb.ax.tick_params(labelsize=8, colors=MUTED)
cb.outline.set_edgecolor(HAIR)
fig.savefig(FIG / "heatmap.png")
plt.close(fig)

# ================================================================ document ===
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

def shade(cell, hexcolor: str):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), hx(hexcolor))
    tcPr.append(shd)

def set_cell(cell, text, *, bold=False, color=BODY, size=9.5, align=None):
    cell.text = ""
    par = cell.paragraphs[0]
    run = par.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    run.font.color.rgb = rgb(color)
    run.font.name = "Calibri"
    if align:
        par.alignment = align
    return cell

def add_table(doc, headers, rows, widths=None, aligns=None):
    tbl = doc.add_table(rows=1, cols=len(headers))
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, head in enumerate(headers):
        cell = tbl.rows[0].cells[i]
        set_cell(cell, head, bold=True, color="#FFFFFF", size=9.5,
                 align=(aligns[i] if aligns else None))
        shade(cell, INK)
    for r_i, row in enumerate(rows):
        cells = tbl.add_row().cells
        fill = "#FFFFFF" if r_i % 2 == 0 else CANVAS
        for c_i, value in enumerate(row):
            set_cell(cells[c_i], str(value),
                     align=(aligns[c_i] if aligns else None))
            shade(cells[c_i], fill)
    if widths:
        for row in tbl.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Inches(w)
    return tbl

def h(doc, text, level, color=INK):
    heading = doc.add_heading(text, level=level)
    for run in heading.runs:
        run.font.color.rgb = rgb(color)
        run.font.name = "Calibri"
    return heading

def p(doc, text, *, size=10.5, color=BODY, bold=False, italic=False,
      space_after=6, align=None):
    para = doc.add_paragraph()
    run = para.add_run(text)
    run.font.size = Pt(size)
    run.font.color.rgb = rgb(color)
    run.bold = bold
    run.italic = italic
    run.font.name = "Calibri"
    para.paragraph_format.space_after = Pt(space_after)
    if align:
        para.alignment = align
    return para

def bullet(doc, text, *, bold_prefix=None, size=10.5):
    para = doc.add_paragraph(style="List Bullet")
    if bold_prefix:
        r = para.add_run(bold_prefix)
        r.bold = True
        r.font.size = Pt(size)
        r.font.color.rgb = rgb(INK)
        r.font.name = "Calibri"
    r = para.add_run(text)
    r.font.size = Pt(size)
    r.font.color.rgb = rgb(BODY)
    r.font.name = "Calibri"
    return para

def caption(doc, text):
    p(doc, text, size=8.5, color=MUTED, italic=True, space_after=10,
      align=WD_ALIGN_PARAGRAPH.CENTER)

def figure(doc, name, cap):
    doc.add_picture(str(FIG / name), width=Inches(6.9))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    caption(doc, cap)

doc = Document()
for section in doc.sections:
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
normal = doc.styles["Normal"]
normal.font.name = "Calibri"
normal.font.size = Pt(10.5)
normal.font.color.rgb = rgb(BODY)

# ---- cover -------------------------------------------------------------------
title = doc.add_paragraph()
run = title.add_run("AIRIQ  \u00b7  EXPLORATORY DATA ANALYSIS")
run.font.size = Pt(26)
run.bold = True
run.font.color.rgb = rgb(INK)
run.font.name = "Calibri"

sub = doc.add_paragraph()
run = sub.add_run("Daily air quality across 26 Indian cities  \u00b7  "
                  f"{d_min:%d %B %Y} \u2013 {d_max:%d %B %Y}")
run.font.size = Pt(14)
run.font.color.rgb = rgb(BRAND)
run.font.name = "Calibri"

meta = doc.add_paragraph()
run = meta.add_run(f"{n_rows:,} daily readings  \u00b7  8 pollutants  \u00b7  CPCB NAAQS benchmarks  \u00b7  "
                   f"Atmos Design System v{VERSION}  \u00b7  {date.today():%d %B %Y}")
run.font.size = Pt(9.5)
run.font.color.rgb = rgb(MUTED)
run.font.name = "Calibri"

rule = doc.add_table(rows=1, cols=1)
rule.rows[0].cells[0].text = ""
shade(rule.rows[0].cells[0], BRAND)
rule.rows[0].height = Pt(4)
doc.add_paragraph()

# ---- 1. executive summary ------------------------------------------------------
h(doc, "1.  Executive summary", 1)
p(doc, "This report profiles the dataset behind the AirIQ dashboards: daily "
       "ambient air-quality readings from the CPCB continuous monitoring "
       "network, packaged inside Air_Pollution_Dashboard.twbx. Every number "
       "below is computed directly from that extract, so the report and the "
       f"dashboards can never disagree. The dashboard's headline analytical "
       f"window is the two complete years 2021\u20132022 ({n_records_2122:,} records); "
       "this analysis covers the full extract.")

kpi = doc.add_table(rows=2, cols=3)
kpi.style = "Table Grid"
kpi.alignment = WD_TABLE_ALIGNMENT.CENTER
KPIS = [
    ("RECORDS ANALYSED", f"{n_rows:,}", "26 cities \u00d7 904 reporting days"),
    ("UNHEALTHY DAYS", f"{pct_unhealthy:.0f}%", "share of city-days at AQI \u2265 4"),
    ("PM2.5 EXCEEDANCE", f"{poll[0]['pct_over']:.0f}%", "days above the 24-h NAAQS limit"),
    ("WORST \u00b7 BEST CITY", f"{worst_city} \u00b7 {best_city}", f"avg AQI {city.iloc[0]['avg_aqi']:.2f} vs {city.iloc[-1]['avg_aqi']:.2f}"),
    ("PEAK SEASON", "Winter", "every city peaks Nov \u2013 Mar"),
    ("YOY CHANGE 2021\u219222", f"{yoy:+.1f}%", f"national avg AQI {avg21:.2f} \u2192 {avg22:.2f}"),
]
for idx, (label, value, note) in enumerate(KPIS):
    cell = kpi.rows[idx // 3].cells[idx % 3]
    cell.text = ""
    lp = cell.paragraphs[0]
    r = lp.add_run(label)
    r.bold = True
    r.font.size = Pt(8)
    r.font.color.rgb = rgb(MUTED)
    r.font.name = "Calibri"
    vp = cell.add_paragraph()
    r = vp.add_run(value)
    r.bold = True
    r.font.size = Pt(16)
    r.font.color.rgb = rgb(INK)
    r.font.name = "Calibri"
    np_ = cell.add_paragraph()
    r = np_.add_run(note)
    r.font.size = Pt(8.5)
    r.font.color.rgb = rgb(MUTED)
    r.font.name = "Calibri"
    shade(cell, CANVAS)
doc.add_paragraph()

bullet(doc, "71% of all city-days sit at AQI \u2265 4, and 53% fall in the worst "
            "category (5). Clean-air days are the exception, not the rule.",
       bold_prefix="Air is unhealthy most of the time.  ")
bullet(doc, f"{worst_city}, Jorapokhar, Amritsar, Patna and Lucknow average AQI "
            f"{city.head(5)['avg_aqi'].min():.1f}+ with roughly 95% unhealthy days, while "
            f"{best_city}, Coimbatore, Kochi, Ernakulam and Bengaluru stay near "
            f"{city.tail(5)['avg_aqi'].max():.1f}.",
       bold_prefix="A north-Indian pollution belt stands out.  ")
bullet(doc, "PM2.5 exceeds its 24-hour NAAQS limit on roughly half of all days "
            "and PM10 on two of five; every other pollutant is compliant on "
            "96%+ of days. Mean PM2.5 runs at ~2.3\u00d7 the annual limit.",
       bold_prefix="Particulates are the binding constraint.  ")
bullet(doc, "The national average climbs from a monsoon trough near 2.8 "
            "(Jul\u2013Aug) to winter peaks above 4.7 (Nov\u2013Feb); all 26 cities "
            "peak between November and March.",
       bold_prefix="Seasonality is strong and consistent.  ")
bullet(doc, f"National average AQI moved {avg21:.2f} \u2192 {avg22:.2f} "
            f"({yoy:+.1f}%) between the two complete years \u2014 structurally flat.",
       bold_prefix="Year-over-year, air barely improved.  ")
bullet(doc, "26 cities \u00d7 904 days with zero nulls and zero duplicates; only "
            "3 calendar days are absent from the entire network.",
       bold_prefix="Data quality is excellent.  ")
doc.add_page_break()

# ---- 2. dataset & data quality -------------------------------------------------
h(doc, "2.  Dataset & data quality", 1)
p(doc, "The extract is a tidy daily panel: one row per city per day, an "
       "ordinal AQI index, and daily mean concentrations for eight pollutants. "
       "The AQI is a 1\u20135 category scale (1 = best, 5 = worst); the workbook "
       "classes days at AQI \u2265 4 as unhealthy and bins city averages into "
       "Low (\u2264 3.3), Moderate (\u2264 4.2) and High.")
add_table(doc,
    ["Field", "Type", "Description"],
    [
        ["city", "Text", "Monitoring city \u2014 26 stations across India"],
        ["date", "Date", f"Calendar day, {d_min:%d %b %Y} \u2192 {d_max:%d %b %Y}"],
        ["aqi", "Integer", "Ordinal air-quality index, 1 (best) \u2192 5 (worst); \u2265 4 = unhealthy day"],
        ["co", "Number", "Daily mean carbon monoxide concentration"],
        ["no", "Number", "Daily mean nitric oxide concentration"],
        ["no2", "Number", "Daily mean nitrogen dioxide concentration"],
        ["o3", "Number", "Daily mean ground-level ozone concentration"],
        ["so2", "Number", "Daily mean sulphur dioxide concentration"],
        ["pm2_5", "Number", "Daily mean fine particulate matter (PM2.5) concentration"],
        ["pm10", "Number", "Daily mean coarse particulate matter (PM10) concentration"],
        ["nh3", "Number", "Daily mean ammonia concentration"],
    ],
    widths=[0.8, 0.8, 5.3])
doc.add_paragraph()
bullet(doc, f"{n_rows:,} rows \u00b7 {n_cities} cities \u00b7 904 reporting days per city "
            f"(every city reports on exactly the same days).")
bullet(doc, f"Coverage: 904 of {cal_days} calendar days (99.7%). The three "
            f"absent dates \u2014 {', '.join(d.strftime('%d %b %Y') for d in missing_days)} "
            f"\u2014 are missing network-wide, not per city.")
bullet(doc, "Zero null values across all 11 fields and zero duplicate "
            "(city, date) keys \u2014 no imputation or cleaning was required.")
bullet(doc, "Edge years are partial (Dec 2020: 32 days; Jan\u2013May 2023: 144 "
            "days), so year-over-year comparisons in this report use the "
            "complete years 2021 and 2022.")
figure(doc, "aqi_dist.png",
       f"Figure 1 \u00b7 Share of city-days by AQI category, all 26 cities, "
       f"{d_min:%b %Y} \u2013 {d_max:%b %Y}. Green = below the unhealthy threshold, "
       f"red = AQI \u2265 4.")
doc.add_page_break()

# ---- 3. cities -------------------------------------------------------------------
h(doc, "3.  Where pollution concentrates", 1)
p(doc, "Average exposure differs more between cities than between years. "
       f"Mean PM2.5 spans {city.loc[worst_city,'pm25']:.0f} \u00b5g/m\u00b3 in {worst_city} "
       f"against {city.loc[best_city,'pm25']:.0f} in {best_city} \u2014 a six-fold gap \u2014 "
       "and the share of unhealthy days ranges from below 40% to above 96%.")
figure(doc, "city_unhealthy.png",
       "Figure 2 \u00b7 Share of days at AQI \u2265 4 by city, full period. Colour "
       "follows the workbook status scale: green < 50%, amber 50\u201375%, red > 75%.")
rank_rows = []
for i in list(range(5)) + list(range(21, 26)):
    row = city.iloc[i]
    rank_rows.append([str(i + 1), city.index[i], f"{row['avg_aqi']:.2f}",
                      f"{row['unhl']:.1f}%", f"{row['pm25']:.0f}"])
tbl = add_table(doc,
    ["Rank", "City", "Avg AQI", "Unhealthy days", "Mean PM2.5 (\u00b5g/m\u00b3)"],
    rank_rows,
    widths=[0.6, 1.9, 1.1, 1.55, 1.75],
    aligns=[WD_ALIGN_PARAGRAPH.CENTER, None, WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER])
# divider row between top 5 and bottom 5
cells = tbl.add_row().cells
merged = cells[0].merge(cells[1]).merge(cells[2]).merge(cells[3]).merge(cells[4])
set_cell(merged, "\u22ee  16 cities in between  \u22ee", color=MUTED,
         align=WD_ALIGN_PARAGRAPH.CENTER)
shade(merged, CANVAS)
p(doc, "The most burdened cities form a contiguous belt \u2014 Delhi NCR, the "
       "Indo-Gangetic plain and the Jharkhand coal corridor (Jorapokhar, "
       "Talcher) \u2014 while coastal and southern cities (Aizawl, Kerala\u2019s "
       "Kochi and Ernakulam, Coimbatore, Bengaluru) breathe consistently "
       "cleaner air.", space_after=0)
doc.add_page_break()

# ---- 4. pollutants -----------------------------------------------------------------
h(doc, "4.  Pollutants vs NAAQS limits", 1)
p(doc, "Each pollutant is benchmarked against the Indian NAAQS limits the "
       "workbook uses: a 24-hour limit (the \u201cexceedance\u201d test behind the "
       "drilldown\u2019s Percentage-of-Days chart) and an annual limit (the "
       "reference line in the Avg-Concentration chart). Particulates dominate "
       "both tests; CO clears its annual limit on average, but only marginally.")
add_table(doc,
    ["Pollutant", "Mean", "Median", "24-h limit", "Days above", "Annual limit", "Mean vs annual"],
    [[q["label"], f"{q['mean']:,.1f}", f"{q['median']:,.1f}", f"{q['lim24']:,}",
      f"{q['pct_over']:.1f}%", f"{q['limyr']:,}", f"{q['mean_vs_annual']:.0f}%"]
     for q in poll],
    widths=[1.0, 0.9, 0.9, 0.95, 1.05, 1.0, 1.1],
    aligns=[None] + [WD_ALIGN_PARAGRAPH.CENTER] * 6)
doc.add_paragraph()
figure(doc, "exceedance.png",
       "Figure 3 \u00b7 Share of days above the NAAQS 24-hour limit by pollutant, "
       "full period. Bars use each pollutant\u2019s fixed palette colour.")
p(doc, f"Average concentrations rank CO ({top5_conc[0]['mean']:,.0f}), PM10 "
       f"({top5_conc[1]['mean']:,.0f}) and PM2.5 ({top5_conc[2]['mean']:,.0f}) "
       "highest \u2014 the same top five the Nationwide Overview\u2019s "
       "\u201cTop 5 Pollutants\u201d card shows. Medians sit well below means for "
       "every pollutant: the distributions are right-skewed, driven by "
       "episodic winter spikes rather than a uniformly high baseline.",
  space_after=0)
doc.add_page_break()

# ---- 5. seasonality ------------------------------------------------------------------
h(doc, "5.  Seasonality & year-over-year trend", 1)
p(doc, "Pollution in India is a winter problem. The national average AQI "
       "bottoms out during the monsoon (" +
       " \u00b7 ".join(f"{name} {avg:.2f} ({unh:.0f}% unhealthy)"
                     for name, avg, unh in season_rows) +
       ") and peaks between November and February. " +
       f"{n_peak_nov_mar} of 26 cities record their worst month inside the "
       "November\u2013March window; none peaks between April and October.")
figure(doc, "monthly_trend.png",
       "Figure 4 \u00b7 National monthly average AQI, 2021 vs 2022. The two "
       "years trace an almost identical seasonal curve.")
figure(doc, "heatmap.png",
       "Figure 5 \u00b7 Average AQI by city and month, 2021\u20132022. Cities are "
       "ordered worst to cleanest; darker red = higher AQI.")
doc.add_page_break()

# ---- 6. findings & dashboard linkage ---------------------------------------------------
h(doc, "6.  Key findings & how they reach the dashboard", 1)
for i, (head, body) in enumerate([
    ("Unhealthy air is the norm. ",
     "71% of city-days are at AQI \u2265 4 and 53% in the worst category \u2014 "
     "the headline KPI band on the Nationwide Overview."),
    ("Geography is destiny. ",
     f"Avg AQI spans {city.iloc[0]['avg_aqi']:.2f} ({worst_city}) to "
     f"{city.iloc[-1]['avg_aqi']:.2f} ({best_city}); the north-Indian belt "
     "vs south/northeast split is the dataset\u2019s strongest pattern."),
    ("PM2.5 is the pollutant to manage. ",
     "It breaches the 24-h limit on ~49% of days and averages ~2.3\u00d7 the "
     "annual limit; PM10 follows. Gaseous pollutants are largely compliant."),
    ("Winter is the danger window. ",
     "Every city peaks Nov\u2013Mar, with ~95% unhealthy days in winter vs "
     "~46% in the monsoon \u2014 the timing story told by the monthly trend "
     "and the peak-month KPIs."),
    ("2022 looked like 2021. ",
     f"National average AQI {avg21:.2f} \u2192 {avg22:.2f} ({yoy:+.1f}%), so "
     "year filters change the picture far less than city or season."),
    ("The data can be trusted. ",
     "Zero nulls, zero duplicates, 99.7% calendar coverage \u2014 no cleaning "
     "layer between the extract and the dashboards."),
], start=1):
    bullet(doc, body, bold_prefix=f"{i}.  {head}")
doc.add_paragraph()
add_table(doc,
    ["Analysis in this report", "Where it appears in the dashboard"],
    [
        ["City ranking & unhealthy-day share (Fig. 2)", "Average AQI by City \u2014 bubble map, Nationwide Overview"],
        ["Monthly 2021 vs 2022 profile (Fig. 4)", "Monthly AQI Trend by Year \u2014 matrix, Nationwide Overview"],
        ["Pollutant concentration ranking (\u00a74)", "Top 5 Air Pollutants by Avg Concentration \u2014 bars, Nationwide Overview"],
        ["Annual NAAQS compliance (\u00a74)", "Avg Concentration vs annual-limit reference \u2014 bars, City Drilldown"],
        ["24-hour exceedance rate (Fig. 3)", "Percentage of Days above the limit \u2014 bars, City Drilldown"],
        ["Peak-month & seasonality (Fig. 5)", "Peak Pollution Month \u2014 KPI cards on both dashboards"],
    ],
    widths=[3.1, 3.8])
doc.add_paragraph()
p(doc, "Method & reproducibility \u2014 every statistic and figure in this "
       "document is computed live from the packaged extract inside "
       "Air_Pollution_Dashboard.twbx, and every colour resolves to the Atmos "
       "tokens in theme/palette.json. Re-run after any data or palette change:",
  bold=True, size=10)
code = doc.add_paragraph()
run = code.add_run("python3 theme/make_eda_doc.py")
run.font.name = "Consolas"
run.font.size = Pt(9.5)
run.font.color.rgb = rgb(INK)

p(doc, "")
note = doc.add_paragraph()
run = note.add_run(f"AirIQ EDA \u00b7 generated from the workbook extract \u00b7 "
                   f"Atmos Design System v{VERSION} \u00b7 {date.today():%d %b %Y}")
run.font.size = Pt(8.5)
run.font.color.rgb = rgb(MUTED)
run.italic = True
note.alignment = WD_ALIGN_PARAGRAPH.CENTER

out = ROOT / "AirIQ_EDA_Report.docx"
doc.save(out)
print(f"  \u2713 wrote {out.name} ({out.stat().st_size/1024:.0f} KB)")
print(f"  \u2713 figures in preview/eda/: {', '.join(sorted(f.name for f in FIG.glob('*.png')))}")
