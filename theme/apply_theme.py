#!/usr/bin/env python3
"""
Atmos theme builder — Air Pollution Dashboard (Tableau .twbx)
=============================================================

This script is the single entry point that applies the modular color system
(theme/palette.json) to the Tableau workbook:

  1. Restyles every worksheet  (surfaces, ink, fonts, mark colors, map style)
  2. Pins every discrete mark color to a palette token (bucket -> color maps)
  3. Rebuilds both dashboards with a professional, card-based layout
  4. Regenerates the navigation button assets from the same tokens
  5. Repackages the .twbx and renders design previews of both dashboards

Usage:
    python3 theme/apply_theme.py            # apply theme + previews
    python3 theme/apply_theme.py --no-preview

Idempotent: running it twice produces the same result.
"""
from __future__ import annotations

import io
import json
import re
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TWBX = ROOT / "Air_Pollution_Dashboard.twbx"
TWB_NAME = "Air_Pollution_Dashboard.twb"
PREVIEW_DIR = ROOT / "preview"
TOKENS = json.loads((ROOT / "theme" / "palette.json").read_text(encoding="utf-8"))
C = TOKENS["color"]
T = TOKENS["type"]

# ---------------------------------------------------------------- constants --
DS = "federated.0r5ypuv12cy9kq13t07co0hzn2wb"          # air_pollution_data
YEAR = f"[{DS}].[yr:date:ok]"
CITY = f"[{DS}].[none:city:nk]"
PARAM1 = "[Parameters].[Parameter 1]"
UUID_D1 = "{202F36E4-D400-48B5-B650-479690FDBBBC}"     # Nationwide Overview
UUID_D2 = "{87EC700B-A4F4-46F3-B4DF-14B41884149F}"     # City Drilldown

def hx(token: str) -> str:
    return token.lower()

# ------------------------------------------------------------- layout specs --
# Zone coordinates are in Tableau's 1/1000-percent space (0..100000 of parent).
# These specs drive BOTH the .twb zone XML and the rendered previews, so the
# preview is always a faithful picture of the generated dashboard.

KPI_CACHE = ("<layout-cache cell-count-h='1' non-cell-size-h='79' "
             "type-h='cell' type-w='fixed' />")
WS_CACHES = {
    "Map": "<layout-cache minwidth='202' type-h='fixed' type-w='scalable' />",
    "Monthly Y Variation": ("<layout-cache cell-count-w='12' minheight='263' "
                            "non-cell-size-w='45' type-h='scalable' type-w='cell' />"),
    "Top 5 Pollutants": ("<layout-cache cell-count-w='5' minheight='162' "
                         "non-cell-size-w='66' type-h='scalable' type-w='cell' />"),
    "Total Records Monitored": "<layout-cache cell-count-w='1' type-h='cell' type-w='cell' />",
    "Avg Conc Year": ("<layout-cache cell-count-w='2' minheight='162' "
                      "non-cell-size-w='66' type-h='scalable' type-w='cell' />"),
    "Monthly Avg Concentration": ("<layout-cache cell-count-w='12' minheight='162' "
                                  "non-cell-size-w='66' type-h='scalable' type-w='cell' />"),
    "Percentage of Days": ("<layout-cache cell-count-w='24' minheight='228' "
                           "non-cell-size-w='48' type-h='scalable' type-w='cell' />"),
}

MARGIN = 800          # page side margin
GUTTER = 800          # gutter between cards
CONTENT_W = 100000 - 2 * MARGIN

def row(n: int, y: int, h: int, gap: int = GUTTER):
    """Evenly spaced x-positions for n cards across the content width."""
    w = (CONTENT_W - (n - 1) * gap) // n
    return [(MARGIN + i * (w + gap), y, w, h) for i in range(n)]

DASH_NATIONWIDE = {
    "name": "Nationwide Overview",
    "title1": "INDIA AIR QUALITY INTELLIGENCE",
    "title2": "Nationwide Overview  \u00b7  26 cities  \u00b7  Jan 2021 \u2013 Dec 2022",
    "header": (0, 0, 100000, 9500),
    "header_children": {
        "title": (1200, 1700, 58000, 6400),
        "year_filter": (63400, 2350, 15000, 4800),
        "button": (81000, 2350, 17000, 4800),
    },
    "accent": (0, 9500, 100000, 650),
    "kpis": [
        ("National AQI ", "NATIONAL AVERAGE AQI", "3.6"),
        ("% Change", "YOY CHANGE IN AQI", "-4.8%\n\u25bc"),
        ("Most Polluted City", "MOST POLLUTED CITY", "Delhi"),
        ("Cleanest City", "CLEANEST CITY", "Aizawl"),
        ("Peak Pollution Month", "PEAK POLLUTION MONTH", "November"),
    ],
    "kpi_row": (10600, 15300),
    "main": {
        "map": (800, 27000, 35000, 57000),
        "monthly_y": (36600, 27000, 62600, 28400),
        "top5": (36600, 56400, 62600, 27600),
    },
    "footer": {
        "records": (800, 85100, 30000, 14100),
        "note": (31600, 85100, 67600, 14100),
    },
    "button_kind": "drilldown",
}

DASH_CITY = {
    "name": "City Drilldown",
    "title1": "CITY-LEVEL POLLUTION DRILLDOWN",
    "title2": "Benchmark any of 26 cities against NAAQS pollutant limits",
    "header": (0, 0, 100000, 9500),
    "header_children": {
        "title": (1200, 1700, 62000, 6400),
        "button": (81000, 2350, 17000, 4800),
    },
    "accent": (0, 9500, 100000, 650),
    "control": (800, 10600, 98400, 9600),
    "control_children": {
        "label": (2600, 12200, 42000, 6200),
        "city_filter": (47600, 12200, 15600, 6200),
        "pollutant_param": (64800, 12200, 15600, 6200),
        "legend": (82200, 12200, 15200, 6200),
    },
    "kpis": [
        ("City Selector", "SELECTED CITY", "< Delhi >"),
        ("Avg AQI ", "AVG AQI \u00b7 SELECTED CITY", "4.3"),
        ("C Peak Pollution Month", "PEAK POLLUTION MONTH", "November"),
        ("Unhealthy AQI Days", "% UNHEALTHY AQI DAYS", "62.7%"),
    ],
    "kpi_row": (21400, 15300),
    "main": {
        "avg_conc": (800, 37800, 29000, 45400),
        "monthly_avg": (29800, 37800, 68600, 26000),
        "pct_days": (29800, 64800, 68600, 18400),
    },
    "footer": {"note": (800, 84300, 98400, 14900)},
    "button_kind": "home",
}

def to_parent(child, parent):
    """Convert a ROOT-space rect (x, y, w, h) into parent-fraction coords."""
    px, py, pw, ph = parent
    cx, cy, cw, ch = child
    return (round((cx - px) * 100000 / pw), round((cy - py) * 100000 / ph),
            round(cw * 100000 / pw), round(ch * 100000 / ph))

# ------------------------------------------------------------ xml builders --
def zone_style(bg=None, border=None, bstyle="solid", bwidth=1,
               margin=None, padding=None) -> str:
    fmts = []
    if border is not None:
        fmts.append(f"<format attr='border-color' value='{border}' />")
    fmts.append(f"<format attr='border-style' value='{bstyle}' />")
    fmts.append(f"<format attr='border-width' value='{bwidth}' />")
    if margin is not None:
        fmts.append(f"<format attr='margin' value='{margin}' />")
    if padding is not None:
        fmts.append(f"<format attr='padding' value='{padding}' />")
    if bg is not None:
        fmts.append(f"<format attr='background-color' value='{bg}' />")
    body = "\n            ".join(fmts)
    return f"<zone-style>\n            {body}\n          </zone-style>"

ATTR_ORDER = ["h", "id", "mode", "name", "pane-specification-id", "param",
              "type-v2", "w", "x", "y", "custom-title", "forceUpdate"]

def attrs(kw) -> str:
    keys = sorted(kw, key=lambda k: ATTR_ORDER.index(k) if k in ATTR_ORDER else 99)
    return " ".join(f"{k}='{kw[k]}'" for k in keys)

def ws_zone(zid: int, name: str, x: int, y: int, w: int, h: int,
            cache: str | None = None, style: str | None = None) -> str:
    inner = ""
    if cache:
        inner += f"\n          {cache}"
    if style:
        inner += f"\n          {style}"
    return (f"<zone {attrs(dict(h=h, id=zid, name=name, w=w, x=x, y=y))}>"
            f"{inner}\n        </zone>")

def container(zid: int, x: int, y: int, w: int, h: int, style: str,
              children: str) -> str:
    return (f"<zone {attrs(dict(h=h, id=zid, **{'type-v2': 'layout-basic'}, w=w, x=x, y=y))}>\n"
            f"        {children}\n"
            f"          {style}\n        </zone>")

def text_zone(zid: int, x: int, y: int, w: int, h: int, runs: str,
              style: str | None = None, align: str = "0") -> str:
    style = style or zone_style(bstyle="none", bwidth=0)
    return (f"<zone {attrs(dict(forceUpdate='true', h=h, id=zid, **{'type-v2': 'text'}, w=w, x=x, y=y))}>\n"
            f"          <formatted-text>\n            {runs}\n          </formatted-text>\n"
            f"          {style}\n        </zone>")

def run(text: str, size: int, color: str, bold=False, italic=False,
        font=None, align=None) -> str:
    a = []
    if bold: a.append("bold='true'")
    if italic: a.append("italic='true'")
    if align is not None: a.append(f"fontalignment='{align}'")
    a.append(f"fontcolor='{color}'")
    a.append(f"fontname='{font or T['family']}'")
    a.append(f"fontsize='{size}'")
    return f"<run {' '.join(a)}>{text}</run>"

def esc(text: str) -> str:
    return (text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))

def filter_zone(zid, name, param, x, y, w, h, mode="dropdown") -> str:
    head = attrs(dict(h=h, id=zid, mode=mode, name=name, param=param,
                 **{'type-v2': 'filter'}, w=w, x=x, y=y))
    style = zone_style(bg=hx(C['surface']['card']), border=hx(C['line']['hairline']))
    return f"<zone {head}>\n          {style}\n        </zone>"

def param_zone(zid, x, y, w, h, title) -> str:
    title_run = run(title, 9, hx(C["ink"]["muted"]), bold=True)
    head = attrs(dict(**{'custom-title': 'true'}, h=h, id=zid, mode='compact',
                      param=PARAM1, **{'type-v2': 'paramctrl'}, w=w, x=x, y=y))
    return (f"<zone {head}>\n"
            f"          <formatted-text>\n            {title_run}\n          </formatted-text>\n"
            f"          {zone_style(bg=hx(C['surface']['card']), border=hx(C['line']['hairline']))}\n"
            f"        </zone>")

def legend_zone(zid, x, y, w, h) -> str:
    head = attrs(dict(h=h, id=zid, name='Monthly Avg Concentration',
                      **{'pane-specification-id': 1}, param=YEAR,
                      **{'type-v2': 'color'}, w=w, x=x, y=y))
    return (f"<zone {head}>\n"
            f"          {zone_style(bg=hx(C['surface']['card']), border=hx(C['line']['hairline']))}\n"
            f"        </zone>")

def button_zone(zid, image, tooltip, x, y, w, h, target_uuid) -> str:
    return (f"<zone {attrs(dict(h=h, id=zid, **{'type-v2': 'dashboard-object'}, w=w, x=x, y=y))}>\n"
            f"          <button action='tabdoc:goto-sheet window-id=&quot;{target_uuid}&quot;'>\n"
            f"            <button-visual-state>\n"
            f"              <image-path>Image/{image}</image-path>\n"
            f"              <tooltip-text>{tooltip}</tooltip-text>\n"
            f"            </button-visual-state>\n"
            f"          </button>\n        </zone>")

# -------------------------------------------------------- dashboard builder --
def build_dashboard_1() -> str:
    d = DASH_NATIONWIDE
    hdr = d["header"]
    hc = d["header_children"]

    title_runs = (
        run(esc(d["title1"]), T["dashboardTitle"], hx(C["ink"]["inverse"]),
            bold=True, font=T["familyBold"])
        + run("&#10;" + esc(d["title2"]), T["dashboardSubtitle"],
              hx(C["ink"]["inverseMuted"])))

    header_children = [
        text_zone(11, *to_parent(hc["title"], hdr), title_runs),
        filter_zone(12, "Total Records Monitored", YEAR,
                    *to_parent(hc["year_filter"], hdr)),
        button_zone(13, "btn-drilldown.png", "City-Level Pollution Drilldown",
                    *to_parent(hc["button"], hdr), UUID_D2),
    ]
    header = container(10, *hdr,
                       zone_style(bg=hx(C["surface"]["header"]), bstyle="none", bwidth=0),
                       "\n        ".join(header_children))
    accent = simple_zone(15, d["accent"], hx(C["brand"]["primary"]))

    card = lambda: zone_style(bg=hx(C["surface"]["card"]), border=hx(C["line"]["hairline"]))
    ky, kh = d["kpi_row"]
    kpi_zones = []
    for zid, (x, y, w, h) in enumerate(row(len(d["kpis"]), ky, kh), start=20):
        kpi_zones.append(ws_zone(zid, d["kpis"][zid - 20][0], x, y, w, h,
                                 cache=KPI_CACHE, style=card()))

    m = d["main"]
    main_zones = [
        ws_zone(30, "Map", *m["map"], cache=WS_CACHES["Map"], style=card()),
        ws_zone(31, "Monthly Y Variation", *m["monthly_y"],
                cache=WS_CACHES["Monthly Y Variation"], style=card()),
        ws_zone(32, "Top 5 Pollutants", *m["top5"],
                cache=WS_CACHES["Top 5 Pollutants"], style=card()),
    ]
    f = d["footer"]
    footer_zones = [
        ws_zone(40, "Total Records Monitored", *f["records"],
                cache=WS_CACHES["Total Records Monitored"], style=card()),
        text_zone(41, *f["note"],
                  run("HOW TO READ", 9, hx(C["ink"]["muted"]), bold=True, font=T["familyBold"])
                  + run("&#10;" + esc("Bubble size and color track average AQI by city (green = low, amber = moderate, red = high). "
                                      "Use the Year filter in the header to compare 2021 vs 2022 \u2014 the matrix and bars respond instantly."),
                        T["footnote"], hx(C["ink"]["body"])),
                  style=card()),
    ]

    children = "\n        ".join([header, accent] + kpi_zones + main_zones + footer_zones)
    deps = f"""      <datasources>
        <datasource caption='air_pollution_data (air_pollution_data)' name='{DS}' />
      </datasources>
      <datasource-dependencies datasource='{DS}'>
        <column caption='Date' datatype='date' name='[date]' role='dimension' type='ordinal' />
        <column-instance column='[date]' derivation='Year' name='[yr:date:ok]' pivot='key' type='ordinal' />
      </datasource-dependencies>"""
    return dashboard_xml(d["name"], "Dashboard1", deps, children)

def build_dashboard_2() -> str:
    d = DASH_CITY
    hdr = d["header"]
    hc = d["header_children"]
    ctl, cc = d["control"], d["control_children"]

    title_runs = (
        run(esc(d["title1"]), T["dashboardTitle"], hx(C["ink"]["inverse"]),
            bold=True, font=T["familyBold"])
        + run("&#10;" + esc(d["title2"]), T["dashboardSubtitle"],
              hx(C["ink"]["inverseMuted"])))

    header_children = [
        text_zone(11, *to_parent(hc["title"], hdr), title_runs),
        button_zone(12, "btn-home.png", "Nationwide Air Quality Overview",
                    *to_parent(hc["button"], hdr), UUID_D1),
    ]
    header = container(10, *hdr,
                       zone_style(bg=hx(C["surface"]["header"]), bstyle="none", bwidth=0),
                       "\n        ".join(header_children))
    accent = simple_zone(15, d["accent"], hx(C["brand"]["primary"]))

    card = lambda: zone_style(bg=hx(C["surface"]["card"]), border=hx(C["line"]["hairline"]))
    control_children = [
        text_zone(21, *to_parent(cc["label"], ctl),
                  run("FILTERS &amp; CONTROLS", 9, hx(C["ink"]["muted"]), bold=True,
                      font=T["familyBold"])
                  + run("&#10;" + esc("Choose a city and pollutant \u2014 every chart on this page updates instantly."),
                        T["footnote"], hx(C["ink"]["body"])),
                  style=zone_style(bstyle="none", bwidth=0)),
        filter_zone(22, "City Selector", CITY, *to_parent(cc["city_filter"], ctl)),
        param_zone(23, *to_parent(cc["pollutant_param"], ctl), "POLLUTANT"),
        legend_zone(24, *to_parent(cc["legend"], ctl)),
    ]
    control = container(20, *ctl, card(), "\n        ".join(control_children))

    ky, kh = d["kpi_row"]
    kpi_zones = []
    for zid, (x, y, w, h) in enumerate(row(len(d["kpis"]), ky, kh), start=30):
        kpi_zones.append(ws_zone(zid, d["kpis"][zid - 30][0], x, y, w, h,
                                 cache=KPI_CACHE, style=card()))

    m = d["main"]
    main_zones = [
        ws_zone(40, "Avg Conc Year", *m["avg_conc"], cache=WS_CACHES["Avg Conc Year"], style=card()),
        ws_zone(41, "Monthly Avg Concentration", *m["monthly_avg"],
                cache=WS_CACHES["Monthly Avg Concentration"], style=card()),
        ws_zone(42, "Percentage of Days", *m["pct_days"],
                cache=WS_CACHES["Percentage of Days"], style=card()),
    ]
    footer = text_zone(50, *d["footer"]["note"],
                       run("HOW TO READ", 9, hx(C["ink"]["muted"]), bold=True, font=T["familyBold"])
                       + run("&#10;" + esc("Left \u2014 annual average concentration of the selected pollutant (2021 vs 2022) with the NAAQS annual limit as the dashed reference line. "
                                           "Right \u2014 monthly concentration trend by year, and the share of days above the NAAQS 24-hour limit for the selected city. "
                                           "Source: CPCB ambient air quality monitoring, 26 Indian cities, 2021\u20132022."),
                             T["footnote"], hx(C["ink"]["body"])),
                       style=card())

    children = "\n        ".join([header, accent, control] + kpi_zones + main_zones + [footer])
    deps = f"""      <datasources>
        <datasource name='Parameters' />
        <datasource caption='air_pollution_data (air_pollution_data)' name='{DS}' />
      </datasources>
      <datasource-dependencies datasource='Parameters'>
        <column caption='Pollutant Selector' datatype='string' name='[Parameter 1]' param-domain-type='list' role='measure' type='nominal' value='&quot;PM2.5&quot;'>
          <calculation class='tableau' formula='&quot;PM2.5&quot;' />
          <members>
            <member value='&quot;CO&quot;' />
            <member value='&quot;NO&quot;' />
            <member value='&quot;NO2&quot;' />
            <member value='&quot;O3&quot;' />
            <member value='&quot;SO2&quot;' />
            <member value='&quot;PM2.5&quot;' />
            <member value='&quot;PM10&quot;' />
            <member value='&quot;NH3&quot;' />
          </members>
        </column>
      </datasource-dependencies>
      <datasource-dependencies datasource='{DS}'>
        <column caption='City' datatype='string' name='[city]' role='dimension' semantic-role='[City].[Name]' type='nominal'>
          <semantic-values semantic-role='[Geographical].[Latitude]'>
            <semantic-value key='&quot;Brajrajnagar&quot;' value='21.816700000000001' />
            <semantic-value key='&quot;Gurugram&quot;' value='28.457522999999998' />
            <semantic-value key='&quot;Jorapokhar&quot;' value='22.422454999999999' />
          </semantic-values>
          <semantic-values semantic-role='[Geographical].[Longitude]'>
            <semantic-value key='&quot;Brajrajnagar&quot;' value='83.916700000000006' />
            <semantic-value key='&quot;Gurugram&quot;' value='77.026343999999995' />
            <semantic-value key='&quot;Jorapokhar&quot;' value='85.760650999999996' />
          </semantic-values>
        </column>
        <column-instance column='[city]' derivation='None' name='[none:city:nk]' pivot='key' type='nominal' />
      </datasource-dependencies>"""
    return dashboard_xml(d["name"], "Dashboard2", deps, children)

def simple_zone(zid, rect, bg) -> str:
    x, y, w, h = rect
    head = attrs(dict(h=h, id=zid, **{'type-v2': 'layout-basic'}, w=w, x=x, y=y))
    return (f"<zone {head}>\n"
            f"          {zone_style(bg=bg, bstyle='none', bwidth=0)}\n        </zone>")

DASHBOARD_IDS = {
    "Nationwide Overview": "{60D55F28-9148-4111-ACAE-82318A366B4A}",
    "City Drilldown": "{D08F4275-C266-4185-994A-C4A12A4F42FD}",
}

def dashboard_xml(name, repo_id, deps, children) -> str:
    root_head = attrs(dict(h=100000, id=4, **{'type-v2': 'layout-basic'},
                           w=100000, x=0, y=0))
    return f"""    <dashboard enable-sort-zone-taborder='true' name='{name}'>
      <repository-location id='{repo_id}' path='/workbooks/Air_Pollution_Dashboard' revision='' />
      <style>
        <style-rule element='dashboard'>
          <format attr='background-color' value='{hx(C['surface']['canvas'])}' />
        </style-rule>
      </style>
      <size sizing-mode='automatic' />
{deps}
      <zones>
        <zone {root_head}>
        {children}
          {zone_style(bg=hx(C['surface']['canvas']), bstyle='none', bwidth=0)}
        </zone>
      </zones>
      <simple-id uuid='{DASHBOARD_IDS[name]}' />
    </dashboard>"""

# ---------------------------------------------------- worksheet restyling --
KPI_SHEETS = {
    "National AQI ": ("NATIONAL AVERAGE AQI", 26),
    "% Change": ("YOY CHANGE IN AQI", 24),
    "Most Polluted City": ("MOST POLLUTED CITY", 18),
    "Cleanest City": ("CLEANEST CITY", 18),
    "Peak Pollution Month": ("PEAK POLLUTION MONTH", 16),
    "Total Records Monitored": ("RECORDS MONITORED", 24),
    "Avg AQI ": ("AVG AQI \u00b7 SELECTED CITY", 26),
    "C Peak Pollution Month": ("PEAK POLLUTION MONTH", 16),
    "Unhealthy AQI Days": ("% UNHEALTHY AQI DAYS", 24),
    "City Selector": ("SELECTED CITY", 20),
}
CHART_TITLES = {
    "Map": "Average AQI by City",
    "Monthly Y Variation": "Monthly AQI Trend by Year",
    "Top 5 Pollutants": "Top 5 Air Pollutants by Avg Concentration",
}

def sheet_span(xml: str, name: str):
    i = xml.find(f"<worksheet name='{name}'>")
    if i < 0:
        raise RuntimeError(f"worksheet not found: {name}")
    j = xml.find("</worksheet>", i)
    return i, j

def sub_count(pattern, repl, s, flags=re.S, expect=None, label="", already=None):
    """re.subn that tolerates re-runs: if nothing matched but the intended
    result is already present (idempotent re-theming), treat as success."""
    new, n = re.subn(pattern, repl, s, flags=flags)
    if expect is not None and n != expect:
        if n == 0 and already and already in s:
            return s, 0
        raise RuntimeError(f"[{label}] expected {expect} replacement(s), got {n}")
    return new, n

def restyle_worksheet(xml: str, name: str, title: str | None, kind: str,
                      size: int | None = None) -> str:
    """kind: 'kpi' | 'chart'"""
    i, j = sheet_span(xml, name)
    blk = xml[i:j]

    # 1. title run ------------------------------------------------------------
    if title is None:
        # parameter-driven titles: restyle runs in place, keep the dynamic text
        k = blk.find("<title>")
        l = blk.find("</title>", k)
        if k >= 0 and l >= 0:
            def _fix_run(m):
                is_param = "Parameters].[Parameter 1" in m.group(1)
                color = hx(C["brand"]["primary"] if is_param else C["ink"]["heading"])
                return (f"<run bold='true' fontalignment='1' fontcolor='{color}' "
                        f"fontname='{T['familyBold']}' fontsize='{T['chartTitle']}'>"
                        f"{m.group(1)}</run>")
            span, _ = sub_count(r"<run[^>]*>(.*?)</run>", _fix_run, blk[k:l],
                                label=f"{name}:title-runs")
            blk = blk[:k] + span + blk[l:]
    if title is not None:
        color = hx(C["ink"]["muted"] if kind == "kpi" else C["ink"]["heading"])
        font = T["familyBold"] if kind == "chart" else T["family"]
        bold = "bold='true' " if kind == "chart" else ""
        fsize = T["chartTitle"] if kind == "chart" else T["kpiLabel"]
        new_run = (f"<run {bold}fontcolor='{color}' fontname='{font}' "
                   f"fontsize='{fsize}'>{esc(title)}</run>")
        blk, n = sub_count(
            r"(<title>\s*<formatted-text>\s*)<run[^>]*>.*?</run>(\s*</formatted-text>)",
            lambda m: m.group(1) + new_run + m.group(2), blk,
            expect=1, label=f"{name}:title")
    # 2. surfaces -------------------------------------------------------------
    blk, _ = sub_count(r"(<style-rule element='table'>\s*<format attr='background-color' value=')#(?:[0-9a-fA-F]{6})(')",
                       lambda m: m.group(1) + hx(C["surface"]["card"]) + m.group(2), blk,
                       label=f"{name}:table-bg")
    blk, _ = sub_count(r"(<style-rule element='title'>\s*<format attr='background-color' value=')#(?:[0-9a-fA-F]{6})(')",
                       lambda m: m.group(1) + hx(C["surface"]["card"]) + m.group(2), blk,
                       label=f"{name}:title-bg")

    ink_value = hx(C["ink"]["heading"] if kind == "kpi" else C["ink"]["body"])
    # 3. cell-level value formats (numeric KPIs) ------------------------------
    if kind == "kpi":
        blk = blk.replace("value='Yu Gothic UI Semibold'", f"value='{T['familyBold']}'")  # no-op if themed
        blk, _ = sub_count(r"(<format attr='color' field='[^']+' value=')#(?:000000|132a43)(')",
                           lambda m: m.group(1) + hx(C["ink"]["heading"]) + m.group(2), blk,
                           label=f"{name}:cell-color")
        if size:
            blk, _ = sub_count(r"(<format attr='font-size' field='[^']+' value=')\d+(')",
                               lambda m: m.group(1) + str(size) + m.group(2), blk,
                               label=f"{name}:cell-size")
    # 4. worksheet-level formats ----------------------------------------------
    if kind == "kpi":
        fam = T["familyBold"] if size else T["family"]
        blk, _ = sub_count(r"(<style-rule element='worksheet'>.*?)</style-rule>",
                           lambda m: _fix_worksheet_rule(m.group(0), fam, hx(C["ink"]["heading"]), size),
                           blk, label=f"{name}:ws-rule")
    else:
        blk, _ = sub_count(r"(<style-rule element='worksheet'>.*?)</style-rule>",
                           lambda m: _fix_worksheet_rule(m.group(0), T["family"], hx(C["ink"]["body"]), None),
                           blk, label=f"{name}:ws-rule")
        # muted dashed reference lines
        blk = blk.replace("<format attr='stroke-color' id='refline0' value='#000000' />",
                          f"<format attr='stroke-color' id='refline0' value='{hx(C['ink']['muted'])}' />")
    return xml[:i] + blk + xml[j:]

def _fix_worksheet_rule(rule: str, family: str, color: str, size: int | None) -> str:
    rule = re.sub(r"<format attr='font-family' value='[^']*' />",
                  f"<format attr='font-family' value='{family}' />", rule)
    rule = re.sub(r"<format attr='color' value='#(?:[0-9a-fA-F]{6})' />",
                  f"<format attr='color' value='{color}' />", rule)
    if size:
        rule = re.sub(r"<format attr='font-size' value='\d+' />",
                      f"<format attr='font-size' value='{size}' />", rule)
    return rule

def bucket_encoding(field: str, mapping: dict) -> str:
    maps = []
    for value, color in mapping.items():
        bucket = value if value.isdigit() else f"&quot;{value}&quot;"
        maps.append(f"<map to='{hx(color)}'>\n              <bucket>{bucket}</bucket>\n            </map>")
    return (f"<encoding attr='color' field='{field}' type='palette'>\n            "
            + "\n            ".join(maps) + "\n          </encoding>")

def year_encoding() -> str:
    maps = "\n            ".join(
        f"<map to='{hx(v)}'>\n              <bucket>{k}</bucket>\n            </map>"
        for k, v in C["year"].items())
    return (f"<encoding attr='color' field='[yr:date:ok]' type='palette'>\n            "
            + maps + "\n          </encoding>")

def apply_theme_to_twb(xml: str) -> str:
    report = []

    # ---- worksheets ----------------------------------------------------------
    for name, (title, size) in KPI_SHEETS.items():
        xml = restyle_worksheet(xml, name, title, "kpi", size)
    for name in ("Map", "Monthly Y Variation", "Top 5 Pollutants",
                 "Avg Conc Year", "Monthly Avg Concentration", "Percentage of Days"):
        xml = restyle_worksheet(xml, name, CHART_TITLES.get(name), "chart")
    report.append("worksheets: titles, surfaces, ink and fonts restyled")

    # ---- mark colors (panes) -------------------------------------------------
    amber = hx(C['categorical']['amber']); primary = hx(C['brand']['primary'])
    xml, _ = sub_count(r"<format attr='mark-color' value='#4e79a7' />",
                       f"<format attr='mark-color' value='{amber}' />",
                       xml, expect=1, label="monthlyY-delta-color",
                       already=f"<format attr='mark-color' value='{amber}' />")
    xml, _ = sub_count(r"<format attr='mark-color' value='#f28e2b' />",
                       f"<format attr='mark-color' value='{primary}' />",
                       xml, expect=1, label="monthlyY-aqi-color",
                       already=f"<format attr='mark-color' value='{primary}' />")

    # Avg Conc Year: single gray bars -> colored by year + add color encoding
    i, j = sheet_span(xml, "Avg Conc Year")
    blk = xml[i:j]
    blk, _ = sub_count(r"<format attr='mark-color' value='#898989' />",
                       year_encoding(), blk, expect=1, label="avgconc-year-encoding",
                       already=year_encoding())
    def _add_color(m):
        return m.group(1) + f"<color column='{YEAR}' />\n              " + m.group(2)
    blk, _ = sub_count(r"(<encodings>\s*)(<text column=')", _add_color,
                       blk, expect=1, label="avgconc-color-pill",
                       already=f"<color column='{YEAR}' />")
    xml = xml[:i] + blk + xml[j:]

    # Monthly Avg Concentration: gray mark color -> year palette
    i, j = sheet_span(xml, "Monthly Avg Concentration")
    blk = xml[i:j]
    blk, _ = sub_count(r"<format attr='mark-color' value='#606b76' />",
                       year_encoding(), blk, expect=1, label="monthlyavg-year-encoding",
                       already=year_encoding())
    xml = xml[:i] + blk + xml[j:]

    # ---- bucket -> color maps (pin every discrete color to the palette) ------
    def replace_encoding(xml, field, mapping, expect):
        pat = (r"<encoding attr='color' field='" + re.escape(field)
               + r"' type='palette'>.*?</encoding>")
        return sub_count(pat, bucket_encoding(field, mapping), xml,
                         expect=expect, label=f"encoding:{field}")

    xml, _ = replace_encoding(xml, "[none:Attribute:nk]", C["pollutant"], 1)
    xml, _ = replace_encoding(xml, "[usr:Calculation_481885172091969556:nk]",
                              C["aqiCategory"], 1)
    xml, _ = replace_encoding(xml, "[usr:Calculation_481885172375638083:nk]",
                              C["exceedanceFlag"], 1)
    report.append("mark colors: pollutants, AQI categories, exceedance flags, "
                  "years pinned to palette tokens")

    # ---- map style -----------------------------------------------------------
    xml, _ = sub_count(r"<format attr='map-style' value='dark' />",
                       "<format attr='map-style' value='normal' />", xml,
                       expect=1, label="map-style",
                       already="<format attr='map-style' value='normal' />")
    xml, _ = sub_count(r"<format attr='washout' value='0' />",
                       "<format attr='washout' value='30' />", xml,
                       expect=1, label="map-washout",
                       already="<format attr='washout' value='30' />")
    report.append("map: light basemap with 30% washout")

    # ---- City Selector sheet extras -----------------------------------------
    i, j = sheet_span(xml, "City Selector")
    blk = xml[i:j]
    if "fontname='Yu Gothic UI Semibold'" in blk:
        blk = blk.replace("fontname='Yu Gothic UI Semibold'", f"fontname='{T['familyBold']}'")
    # customized "<City>" label: bold ink numerals
    blk, _ = sub_count(r"<run bold='true'( fontcolor='[0-9a-f]{6}')? fontname='Tableau Bold' fontsize='20'>",
                       f"<run bold='true' fontcolor='{hx(C['ink']['heading'])}' "
                       f"fontname='{T['familyBold']}' fontsize='20'>",
                       blk, label="cityselector-label")
    blk, _ = sub_count(r"(<style-rule element='quick-filter'>.*?<run)[^>]*>(City)</run>",
                       lambda m: (m.group(1)
                                  + f" bold='true' fontcolor='{hx(C['ink']['heading'])}' "
                                    f"fontname='{T['familyBold']}' fontsize='11'>"
                                    + m.group(2) + "</run>"),
                       blk, expect=1, label="cityselector-qf-title")
    xml = xml[:i] + blk + xml[j:]

    # ---- tooltip labels + global legacy-hex sweep ----------------------------
    xml = xml.replace("fontcolor='#787878'", f"fontcolor='{hx(C['ink']['muted'])}'")
    legacy = {
        "#4e79a7": hx(C["categorical"]["blue"]),   "#f28e2b": hx(C["categorical"]["amber"]),
        "#76b7b2": hx(C["categorical"]["teal"]),   "#9c755f": hx(C["categorical"]["brown"]),
        "#c3ce3d": hx(C["categorical"]["olive"]),  "#e15759": hx(C["categorical"]["rose"]),
        "#5fbb68": hx(C["status"]["good"]),        "#51b364": hx(C["status"]["good"]),
        "#eb1e2c": hx(C["status"]["bad"]),         "#e03531": hx(C["status"]["bad"]),
        "#f9a729": hx(C["status"]["moderate"]),
    }
    for old, new in legacy.items():
        xml = xml.replace(f"to='{old}'", f"to='{new}'")
    report.append("legacy colors: any remaining legacy hex values remapped to tokens")

    # ---- dashboards ------------------------------------------------------------
    d1, d2 = build_dashboard_1(), build_dashboard_2()
    i = xml.find("<dashboards>")
    j = xml.find("</dashboards>") + len("</dashboards>")
    if i < 0 or j < len("</dashboards>"):
        raise RuntimeError("dashboards section not found")
    xml = xml[:i] + "<dashboards>\n" + d1 + "\n" + d2 + "\n  </dashboards>" + xml[j:]
    report.append("dashboards: rebuilt with professional card layout "
                  "(header, controls, KPI band, chart grid, footnotes)")

    # ---- windows ---------------------------------------------------------------
    xml, _ = sub_count(r"<window class='dashboard' maximized='true' name='Dashboard 1'>",
                       "<window class='dashboard' maximized='true' name='Nationwide Overview'>",
                       xml, expect=1, label="win-d1",
                       already="name='Nationwide Overview'")
    xml, _ = sub_count(r"<window class='dashboard' name='Dashboard 2'>",
                       "<window class='dashboard' name='City Drilldown'>",
                       xml, expect=1, label="win-d2", already="name='City Drilldown'")
    xml, _ = sub_count(r"<active id='26' />", "<active id='-1' />", xml,
                       expect=1, label="win-active", already="<active id='-1' />")
    report.append("windows: dashboards renamed, stale zone references cleared")

    for line in report:
        print(f"  \u2713 {line}")
    return xml

# ------------------------------------------------------------ button assets --
def make_button(label: str, arrow: str) -> bytes:
    from PIL import Image, ImageDraw, ImageFont
    W, H = 1360, 200
    img = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    def rgb(hexstr):
        hexstr = hexstr.lstrip("#")
        return tuple(int(hexstr[k:k+2], 16) for k in (0, 2, 4))
    fill = rgb(hx(C["surface"]["header"]))
    border = rgb(hx(C["line"]["headerRule"]))
    textc = rgb(hx(C["ink"]["inverse"]))
    r = 28
    d.rounded_rectangle([2, 2, W - 3, H - 3], radius=r, fill=fill,
                        outline=border, width=4)
    font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 62)
    text = f"{arrow}  {label}" if arrow == "\u2190" else f"{label}  {arrow}"
    bbox = d.textbbox((0, 0), text, font=font)
    d.text(((W - (bbox[2] - bbox[0])) / 2 - bbox[0], (H - (bbox[3] - bbox[1])) / 2 - bbox[1]),
           text, font=font, fill=textc)
    out = io.BytesIO()
    img.save(out, format="PNG")
    return out.getvalue()

def make_assets() -> dict[str, bytes]:
    return {
        "Image/btn-drilldown.png": make_button("CITY DRILLDOWN", "\u2192"),
        "Image/btn-home.png": make_button("NATIONWIDE VIEW", "\u2190"),
    }

# ----------------------------------------------------------------- previews --
def render_preview(spec: dict, path: Path, buttons: dict):
    from PIL import Image, ImageDraw, ImageFont
    W, H = 1600, 900
    sx, sy = W / 100000, H / 100000

    def rgb(hexstr):
        hexstr = hexstr.lstrip("#")
        return tuple(int(hexstr[k:k+2], 16) for k in (0, 2, 4))
    t = lambda k: rgb(hx(C[k[0]][k[1]])) if isinstance(k, tuple) else rgb(hx(C["surface"][k]))
    F = lambda s, b=False: ImageFont.truetype(
        f"/usr/share/fonts/truetype/dejavu/DejaVuSans{'-Bold' if b else ''}.ttf", s)

    img = Image.new("RGB", (W, H), t("canvas"))
    d = ImageDraw.Draw(img)

    def rect(zone, fill=None, outline=None, width=1, pad=0):
        x, y, w, h = zone
        box = [x * sx + pad, y * sy + pad, (x + w) * sx - pad, (y + h) * sy - pad]
        d.rectangle(box, fill=fill, outline=outline, width=width)

    def text(zone, s, font, fill, anchor="la", dx=0, dy=0):
        x, y, w, h = zone
        d.text((x * sx + dx, y * sy + dy + h * sy / 2), s, font=font, fill=fill, anchor=anchor)

    card = dict(fill=t("card"), outline=t(("line", "hairline")), width=1)

    # header + accent -----------------------------------------------------------
    rect(spec["header"], fill=t("header"))
    rect(spec["accent"], fill=t(("brand", "primary")))
    tx, ty, tw, th = spec["header_children"]["title"]
    d.text((tx * sx_ + 6, (ty + th * 0.28) * sy_), spec["title1"], font=F(26, True),
           fill=t(("ink", "inverse")), anchor="lm")
    d.text((tx * sx_ + 8, (ty + th * 0.78) * sy_), spec["title2"], font=F(14),
           fill=t(("ink", "inverseMuted")), anchor="lm")
    btn_key = ("Image/btn-drilldown.png" if spec["button_kind"] == "drilldown"
               else "Image/btn-home.png")
    bx, by, bw, bh = spec["header_children"]["button"]
    btn = Image.open(io.BytesIO(buttons[btn_key])).resize(
        (int(bw * sx_), int(bh * sy_)), Image.LANCZOS)
    img.paste(btn, (int(bx * sx_), int(by * sy_)), btn)

    # KPI cards -------------------------------------------------------------------
    zones = row(len(spec["kpis"]), *spec["kpi_row"])
    for zone, (ws, label, value) in zip(zones, spec["kpis"]):
        rect(zone, **card)
        x, y, w, h = zone
        cx = x * sx + w * sx / 2
        d.text((cx, y * sy + 26), label, font=F(13), fill=t(("ink", "muted")), anchor="ma")
        d.line([x * sx + 14, y * sy + 44, x * sx + w * sx - 14, y * sy + 44],
               fill=t(("line", "hairline")), width=1)
        big = 34 if len(value) < 8 else 30
        vy = y * sy + h * sy / 2 + 14
        for k, line in enumerate(value.split("\n")):
            d.text((cx, vy + k * 30), line, font=F(big, True), fill=t(("ink", "heading")), anchor="mm")

    if spec["name"] == "Nationwide Overview":
        render_nationwide(img, d, spec, t, F, rect, card, buttons)
    else:
        render_city(img, d, spec, t, F, rect, card, buttons)

    d.text((W - 12, H - 10), "Design preview generated from theme/palette.json \u2014 not a Tableau screenshot",
           font=F(11), fill=t(("ink", "muted")), anchor="rd")
    path.parent.mkdir(parents=True, exist_ok=True)
    img.save(path)
    print(f"  \u2713 preview: {path.relative_to(ROOT)}")

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

def chart_title(zone, s, t, F):
    x, y, w, h = zone
    from PIL import ImageDraw
    return s  # placeholder (title drawn by caller)

def _title(d, zone, s, t, F, pad=12):
    x, y, w, h = zone
    d.text((x + pad, y + pad), s, font=F(15, True), fill=t(("ink", "heading")))

def _caption(d, zone, s, t, F, pad=12, dy=34):
    x, y, w, h = zone
    d.text((x + pad, y + pad + dy), s, font=F(11), fill=t(("ink", "muted")))

def render_nationwide(img, d, spec, t, F, rect, card, buttons):
    m = spec["main"]
    # Map ---------------------------------------------------------------------------
    z = m["map"]; rect(z, **card)
    _title(d, z, "Average AQI by City", t, F)
    _caption(d, z, "bubble size = avg AQI \u00b7 color = category", t, F)
    x, y, w, h = z
    mx, my, mw, mh = x * sx_ + 20, y * sy_ + 66, w * sx_ - 40, h * sy_ - 96
    d.rectangle([mx, my, mx + mw, my + mh], fill="#E9EEF3", outline=t(("line", "hairline")))
    india = [(0.44,0.04),(0.38,0.10),(0.33,0.20),(0.30,0.34),(0.24,0.42),(0.22,0.55),
             (0.28,0.68),(0.33,0.80),(0.36,0.93),(0.44,0.97),(0.50,0.90),(0.55,0.78),
             (0.62,0.66),(0.70,0.55),(0.78,0.47),(0.88,0.42),(0.96,0.34),(0.88,0.30),
             (0.76,0.30),(0.62,0.26),(0.52,0.16)]
    pts = [(mx + a * mw, my + b * mh) for a, b in india]
    d.polygon(pts, fill="#F3F6F9", outline="#C7D2DD")
    cities = [  # (x, y, r, color)
        (0.42,0.22,11,"bad"),(0.40,0.26,10,"bad"),(0.33,0.20,8,"moderate"),(0.36,0.30,9,"moderate"),
        (0.52,0.32,10,"bad"),(0.60,0.36,10,"bad"),(0.70,0.44,9,"bad"),(0.58,0.42,8,"moderate"),
        (0.64,0.40,8,"moderate"),(0.56,0.45,8,"moderate"),(0.24,0.46,9,"moderate"),(0.26,0.56,8,"moderate"),
        (0.37,0.47,8,"moderate"),(0.44,0.66,8,"moderate"),(0.43,0.79,7,"good"),(0.50,0.83,7,"good"),
        (0.37,0.87,6,"good"),(0.37,0.92,5,"good"),(0.55,0.73,7,"good"),(0.43,0.76,6,"good"),
        (0.85,0.37,6,"good"),(0.80,0.41,6,"good"),(0.89,0.43,5,"good"),(0.50,0.77,6,"good"),
        (0.37,0.84,6,"good"),(0.60,0.34,7,"moderate")]
    for a, b, r, c in cities:
        cx, cy = mx + a * mw, my + b * mh
        col = t(("status", c))
        d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=col, outline="#FFFFFF", width=2)

    # Monthly Y Variation ------------------------------------------------------------
    z = m["monthly_y"]; rect(z, **card)
    _title(d, z, "Monthly AQI Trend by Year", t, F)
    x, y, w, h = z
    gx, gy = x * sx_ + 26, y * sy_ + 62
    gw, gh = w * sx_ - 52, h * sy_ - 84
    cols = ["2021", "2022", "\u0394 %"]
    colw = gw / 3
    for k, cname in enumerate(cols):
        d.text((gx + k * colw + colw / 2, gy), cname, font=F(12, True),
               fill=t(("ink", "muted")), anchor="ma")
    import random
    rng = random.Random(7)
    rowh = gh / 13
    for r in range(12):
        yy = gy + 22 + r * rowh
        d.text((gx - 6, yy + rowh / 2), MONTHS[r], font=F(11), fill=t(("ink", "body")), anchor="rm")
        v21 = 3.2 + rng.random() * 1.2
        v22 = v21 + (rng.random() - 0.62) * 0.8
        delta = (v22 - v21) / v21 * 100
        d.text((gx + colw / 2, yy + rowh / 2), f"{v21:.1f}", font=F(11),
               fill=t(("series", "prior")), anchor="mm")
        d.text((gx + colw + colw / 2, yy + rowh / 2), f"{v22:.1f}", font=F(11),
               fill=t(("series", "current")), anchor="mm")
        dcol = t(("categorical", "green")) if delta < 0 else t(("categorical", "rose"))
        d.text((gx + 2 * colw + colw / 2, yy + rowh / 2), f"{delta:+.1f}%", font=F(11),
               fill=dcol, anchor="mm")
        if r < 11:
            d.line([gx, yy + rowh, gx + gw, yy + rowh], fill="#F0F3F7", width=1)

    # Top 5 Pollutants ------------------------------------------------------------------
    z = m["top5"]; rect(z, **card)
    _title(d, z, "Top 5 Air Pollutants by Avg Concentration", t, F)
    x, y, w, h = z
    bx, by = x * sx_ + 60, y * sy_ + h * sy_ - 46
    bw, bh = w * sx_ - 120, h * sy_ - 110
    data = [("pm2_5", 0.92, "105"), ("pm10", 0.78, "89"), ("co", 0.42, "48"),
            ("no2", 0.36, "41"), ("o3", 0.28, "32")]
    n = len(data)
    slot = bw / n
    for k, (pol, frac, val) in enumerate(data):
        cx = bx + k * slot + slot / 2
        barw = slot * 0.44
        top = by - bh * frac
        d.rectangle([cx - barw / 2, top, cx + barw / 2, by],
                    fill=t(("pollutant", pol)))
        d.text((cx, top - 18), val, font=F(12, True), fill=t(("ink", "heading")), anchor="ma")
        d.text((cx, by + 10), pol.replace("_", ".").replace("pm", "PM"), font=F(12),
               fill=t(("ink", "body")), anchor="ma")
    d.line([bx - 10, by, bx + bw + 10, by], fill=t(("ink", "muted")), width=1)

    # footer ----------------------------------------------------------------------------
    z = spec["footer"]["records"]; rect(z, **card)
    x, y, w, h = z
    d.text((x * sx_ + 16, y * sy_ + 18), "RECORDS MONITORED", font=F(13),
           fill=t(("ink", "muted")))
    d.text((x * sx_ + 16, y * sy_ + h * sy_ / 2 + 6), "18,980", font=F(34, True),
           fill=t(("ink", "heading")), anchor="lm")
    d.text((x * sx_ + 16, y * sy_ + h * sy_ - 26), "daily readings \u00b7 26 cities \u00b7 2 years",
           font=F(11), fill=t(("ink", "muted")))
    z = spec["footer"]["note"]; rect(z, **card)
    _title(d, z, "HOW TO READ", t, F)
    _caption(d, z, "Bubble size and color track average AQI by city (green = low, amber = moderate, red = high).", t, F, dy=32)
    _caption(d, z, "Use the Year filter in the header to compare 2021 vs 2022 \u2014 the matrix and bars respond instantly.", t, F, dy=54)

def render_city(img, d, spec, t, F, rect, card, buttons):
    # control strip ----------------------------------------------------------------------
    rect(spec["control"], **card)
    cc = spec["control_children"]
    lx, ly, lw, lh = cc["label"]
    d.text((lx * sx_ + 4, (ly + lh * 0.22) * sy_), "FILTERS & CONTROLS", font=F(13, True),
           fill=t(("ink", "muted")), anchor="lm")
    d.text((lx * sx_ + 5, (ly + lh * 0.72) * sy_), "Choose a city and pollutant \u2014 every chart updates instantly.",
           font=F(12), fill=t(("ink", "body")), anchor="lm")
    def chip(zrect, label, value):
        zx, zy, zw, zh = zrect
        box = [zx * sx_, zy * sy_, (zx + zw) * sx_, (zy + zh) * sy_]
        d.rounded_rectangle(box, radius=6, fill=t("card"), outline=t(("line", "hairline")))
        d.text((box[0] + 12, (box[1] + box[3]) / 2), label, font=F(11),
               fill=t(("ink", "muted")), anchor="lm")
        d.text((box[0] + 68, (box[1] + box[3]) / 2), value, font=F(13, True),
               fill=t(("ink", "heading")), anchor="lm")
        d.text((box[2] - 12, (box[1] + box[3]) / 2), "\u25be", font=F(13),
               fill=t(("ink", "muted")), anchor="rm")
    chip(cc["city_filter"], "City", "Delhi")
    chip(cc["pollutant_param"], "Pollutant", "PM2.5")
    zx, zy, zw, zh = cc["legend"]
    cy = (zy + zh / 2) * sy_
    lx = zx * sx_
    d.text((lx + 6, cy), "Year", font=F(11), fill=t(("ink", "muted")), anchor="lm")
    d.rectangle([lx + 52, cy - 8, lx + 72, cy + 8], fill=t(("series", "prior")))
    d.text((lx + 80, cy), "2021", font=F(12), fill=t(("ink", "body")), anchor="lm")
    d.rectangle([lx + 130, cy - 8, lx + 150, cy + 8], fill=t(("series", "current")))
    d.text((lx + 158, cy), "2022", font=F(12), fill=t(("ink", "body")), anchor="lm")

    m = spec["main"]
    # Avg Conc Year ------------------------------------------------------------------------
    z = m["avg_conc"]; rect(z, **card)
    _title(d, z, "Average Concentration of PM2.5", t, F)
    _caption(d, z, "2021 vs 2022 \u00b7 \u03bcg/m\u00b3", t, F)
    x, y, w, h = z
    bx, by = x * sx_ + 56, y * sy_ + h * sy_ - 60
    bw, bh = w * sx_ - 112, h * sy_ - 150
    ref = by - bh * 0.55
    for k, (lab, frac, colk, val) in enumerate(
            [("2021", 0.42, ("series", "prior"), "44.6"),
             ("2022", 0.36, ("series", "current"), "38.2")]):
        cx = bx + bw * (0.28 + 0.44 * k)
        barw = bw * 0.22
        d.rectangle([cx - barw / 2, by - bh * frac, cx + barw / 2, by], fill=t(colk))
        d.text((cx, by - bh * frac - 20), val, font=F(13, True), fill=t(("ink", "heading")), anchor="ma")
        d.text((cx, by + 12), lab, font=F(12), fill=t(("ink", "body")), anchor="ma")
    for dash in range(int(bx - 20), int(bx + bw + 20), 12):
        d.line([dash, ref, dash + 6, ref], fill=t(("ink", "muted")), width=2)
    d.text((bx + bw + 24, ref), "NAAQS\nannual limit", font=F(10), fill=t(("ink", "muted")), anchor="lm")

    # Monthly Avg Concentration ---------------------------------------------------------------
    z = m["monthly_avg"]; rect(z, **card)
    _title(d, z, "Monthly Avg Concentration Trend \u2014 PM2.5", t, F)
    x, y, w, h = z
    bx, by = x * sx_ + 46, y * sy_ + h * sy_ - 52
    bw, bh = w * sx_ - 92, h * sy_ - 120
    import random
    rng = random.Random(3)
    vals = [(.30 + .45 * (1 if i in (10, 11, 0, 1) else 0) + rng.random() * .18) for i in range(12)]
    slot = bw / 12
    for i in range(12):
        cx = bx + i * slot + slot / 2
        v1, v2 = vals[i], vals[i] * (0.82 + rng.random() * 0.2)
        bwid = slot * 0.26
        d.rectangle([cx - bwid - 2, by - bh * v1, cx - 2, by], fill=t(("series", "prior")))
        d.rectangle([cx + 2, by - bh * v2, cx + 2 + bwid, by], fill=t(("series", "current")))
        if i % 2 == 0:
            d.text((cx, by + 10), MONTHS[i], font=F(10), fill=t(("ink", "muted")), anchor="ma")
    d.line([bx - 10, by, bx + bw + 10, by], fill=t(("ink", "muted")), width=1)
    lgx = bx + bw - 220
    d.rectangle([lgx, y * sy_ + 60, lgx + 16, y * sy_ + 76], fill=t(("series", "prior")))
    d.text((lgx + 22, y * sy_ + 68), "2021", font=F(11), fill=t(("ink", "body")), anchor="lm")
    d.rectangle([lgx + 80, y * sy_ + 60, lgx + 96, y * sy_ + 76], fill=t(("series", "current")))
    d.text((lgx + 102, y * sy_ + 68), "2022", font=F(11), fill=t(("ink", "body")), anchor="lm")

    # Percentage of Days ------------------------------------------------------------------------
    z = m["pct_days"]; rect(z, **card)
    _title(d, z, "% of Days Above NAAQS 24-Hour Limit: PM2.5", t, F)
    x, y, w, h = z
    bx, by = x * sx_ + 46, y * sy_ + h * sy_ - 48
    bw, bh = w * sx_ - 92, h * sy_ - 112
    import random
    rng = random.Random(11)
    slot = bw / 24
    for i in range(24):
        frac = 0.15 + rng.random() * 0.8
        col = t(("status", "bad")) if frac > 0.55 else t(("status", "good"))
        cx = bx + i * slot + slot / 2
        d.rectangle([cx - slot * 0.3, by - bh * frac, cx + slot * 0.3, by], fill=col)
        if i % 4 == 0:
            d.text((cx, by + 10), MONTHS[i // 2], font=F(10), fill=t(("ink", "muted")), anchor="ma")
    ref = by - bh * 0.5
    for dash in range(int(bx - 10), int(bx + bw + 10), 12):
        d.line([dash, ref, dash + 6, ref], fill=t(("ink", "muted")), width=2)
    d.text((bx + bw + 10, ref), "threshold", font=F(10), fill=t(("ink", "muted")), anchor="rm")

    # footer --------------------------------------------------------------------------------------
    z = spec["footer"]["note"]; rect(z, **card)
    _title(d, z, "HOW TO READ", t, F)
    _caption(d, z, "Left \u2014 annual average concentration of the selected pollutant (2021 vs 2022) with the NAAQS annual limit as the dashed reference line.", t, F, dy=32)
    _caption(d, z, "Right \u2014 monthly concentration trend by year, and the share of days above the NAAQS 24-hour limit for the selected city.", t, F, dy=52)
    _caption(d, z, "Source: CPCB ambient air quality monitoring, 26 Indian cities, 2021\u20132022.", t, F, dy=72)

# module-level scale factors used by the render helpers
sx_, sy_ = 1600 / 100000, 900 / 100000

# --------------------------------------------------------------------- main --
def main():
    print("Atmos theme builder")
    print("=" * 60)
    if not TWBX.exists():
        sys.exit(f"workbook not found: {TWBX}")

    with zipfile.ZipFile(TWBX) as zf:
        names = zf.namelist()
        contents = {n: zf.read(n) for n in names}
    twb = contents[TWB_NAME].decode("utf-8")

    print("\n[1/4] Applying theme to workbook XML")
    twb = apply_theme_to_twb(twb)

    # XML well-formedness check
    import xml.etree.ElementTree as ET
    ET.fromstring(twb.encode("utf-8"))
    print("  \u2713 generated XML is well-formed")

    print("\n[2/4] Generating navigation button assets from palette")
    buttons = make_assets()

    print("\n[3/4] Repackaging .twbx")
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(TWB_NAME, twb)
        for n in names:
            if n == TWB_NAME or n.startswith("Image/"):
                continue          # old clipart buttons replaced by themed assets
            zf.writestr(n, contents[n])
        for n, data in buttons.items():
            zf.writestr(n, data)
    TWBX.write_bytes(out.getvalue())
    print(f"  \u2713 {TWBX.relative_to(ROOT)} ({TWBX.stat().st_size/1024:.0f} KB)")

    if "--no-preview" not in sys.argv:
        print("\n[4/4] Rendering design previews")
        render_preview(DASH_NATIONWIDE, PREVIEW_DIR / "Nationwide_Overview.png", buttons)
        render_preview(DASH_CITY, PREVIEW_DIR / "City_Drilldown.png", buttons)
    print("\nDone. Open the .twbx in Tableau to see the themed dashboards.")

if __name__ == "__main__":
    main()
