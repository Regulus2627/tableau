# Atmos Design System — Air Pollution Dashboard

A modular color code and layout system for `Air_Pollution_Dashboard.twbx`.
**Every color in the workbook is derived from a single source of truth:
[`theme/palette.json`](theme/palette.json).**

```
┌─────────────────────────────────────────────────────────────┐
│  theme/palette.json        ← edit tokens here (the only     │
│                              place colors are defined)      │
│         │                                                   │
│         ▼                                                   │
│  theme/apply_theme.py       ← builds: restyled worksheets,  │
│         (python3 theme/apply_theme.py)   pinned mark colors,│
│                              rebuilt dashboards, buttons,   │
│                              previews                       │
│         │                                                   │
│         ▼                                                   │
│  Air_Pollution_Dashboard.twbx  +  preview/*.png             │
└─────────────────────────────────────────────────────────────┘
```

## Quick start

```bash
python3 theme/apply_theme.py             # apply theme + render previews
python3 theme/apply_theme.py --no-preview
```

The build is **idempotent** — running it twice produces the same workbook, so
you can change a token and re-run at any time to re-skin everything.

## Color tokens

| Group | Token | Hex | Applied to |
|---|---|---|---|
| **Surface** | `canvas` | `#F2F5F9` | Dashboard background |
| | `card` | `#FFFFFF` | Every worksheet/text card + worksheet table & title backgrounds |
| | `header` | `#17293F` | Dashboard header band |
| **Ink** | `heading` | `#16283C` | Chart titles, KPI values, axis text emphasis |
| | `body` | `#42566C` | Worksheet body text, footnotes |
| | `muted` | `#75879D` | KPI labels, captions, reference lines, tooltip labels |
| | `inverse` / `inverseMuted` | `#FFFFFF` / `#B9C8DB` | Title / subtitle on the navy header |
| **Line** | `hairline` | `#E3EAF2` | Card borders |
| | `headerRule` | `#33517A` | Navigation-button border on the header |
| **Brand** | `primary` | `#2C7DA0` | Accent strip + KPI card accent rules, Avg-AQI marks, 2022 series, parameter highlight in dynamic titles |
| | `primaryDark` | `#1F5F80` | Reserve (hover / emphasis) |
| **Series** | `prior` / `current` | `#B3C5D6` / `#2C7DA0` | Year comparison — **2021 vs 2022** everywhere |
| **Status** | `good` / `moderate` / `bad` | `#4E9D77` / `#E5AF4E` / `#D26A66` | AQI categories (map bubbles), NAAQS exceedance flags |
| **Categorical** | 8 tokens | see palette | Distinct series (blue, amber, green, rose, purple, teal, olive, brown) |
| **Pollutant** | 8 tokens | see palette | Fixed color per pollutant — pm2.5 rose, pm10 amber, co blue, no2 green, o3 purple, so2 teal, nh3 olive, no brown |

### Semantic color locks

Colors are not just consistent, they are **pinned in the workbook XML** as
bucket → color maps at the datasource level, so every sheet renders the same
color for the same value no matter what:

- `2021` → `series.prior`, `2022` → `series.current`
- AQI category `Low / Moderate / High` → `status.good / moderate / bad`
- Exceedance flag `Normal / High` → `status.good / bad`
- Each pollutant (`pm2_5`, `pm10`, `co`, `no2`, `o3`, `so2`, `nh3`, `no`) → its own categorical token

## Typography

Tableau's own font stack is used so the workbook renders identically on every
machine: **Tableau** for body/labels, **Tableau Bold** for titles and KPI
values. Sizes: dashboard title 22, subtitle 10, chart title 12, KPI label 10,
KPI value 26 (numeric) / 18 (text), footnote 9.

## Dashboard architecture

Both dashboards share the same professional scaffold
(header → accent rule → content bands → footnote), laid out on a card grid
with an 0.8 % margin and gutters:

**Nationwide Overview**
1. **Header band** — title + subtitle, Year filter, drill-down button
2. **KPI band** — 5 cards: National Avg AQI · YoY Change · Most Polluted City · Cleanest City · Peak Pollution Month
3. **Chart grid** — city map (bubble = avg AQI, semantic colors) · monthly AQI matrix by year · top-5 pollutant bars
4. **Footnote band** — records monitored stat + "How to read" guidance

**City Drilldown**
1. **Header band** — title + back-to-overview button
2. **Control strip** — city filter, pollutant parameter, year legend
3. **KPI band** — 4 cards: Selected City · Avg AQI · Peak Month · % Unhealthy Days
4. **Chart grid** — annual concentration vs NAAQS limit · **locator map of the selected city** (new `City Map` worksheet, wired to the city filter and auto-zooms to the selection) · monthly trend by year · % days above 24-h limit
5. **Footnote band** — reading guidance + source note

Every stat card carries a slim brand-colored accent rule across its top; the
layout uses 1 % margins and gutters for an airy, balanced grid.

All interactivity from the original workbook is preserved (city filter,
pollutant parameter, year filter, cross-dashboard navigation buttons), and the
new locator map responds to the city filter like every other drilldown sheet.

## Files

| Path | Purpose |
|---|---|
| `Air_Pollution_Dashboard.twbx` | The themed Tableau workbook |
| `theme/palette.json` | Design tokens — the single source of truth (v1.1, softer aesthetic) |
| `theme/apply_theme.py` | Theme builder (worksheet restyle, color pinning, dashboard rebuild, button assets, previews) |
| `preview/Nationwide_Overview.png` | Design preview rendered from the same layout spec + tokens |
| `preview/City_Drilldown.png` | Design preview rendered from the same layout spec + tokens |

## Re-theming example

Want a warmer brand? Change two tokens in `theme/palette.json`:

```json
"brand": { "primary": "#B3541E", "primaryDark": "#8C3F16" },
"surface": { "canvas": "#F7F2EC", "card": "#FFFFFF", "header": "#3B2A1D" }
```

…run `python3 theme/apply_theme.py`, and the accent strip, buttons, 2022
series, parameter highlights, KPI ink and dashboard chrome all update
together — nothing else to touch.
