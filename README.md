# KPI Card — Tableau Viz Extension

A configurable KPI card for Tableau that displays a value, delta badges, goal progress bars, a sparkline chart, and more — all driven by your worksheet data.

## Features

- Big number display with custom prefix/suffix and abbreviation (K/M/B)
- Delta badge vs previous period (green/red, reversible)
- Secondary comparison badge (any calculated field)
- Primary and secondary goal progress bars
- Sparkline chart with data labels, period labels, and legend
- Custom link with optional icon
- Full settings panel (layout, formatting, visibility toggles, labels)
- Copy/paste settings between instances
- Works in Tableau Desktop authoring mode (settings gear hidden in dashboards/published views)

## Files

| File | Purpose |
|------|---------|
| `KPICard.trex` | Tableau extension manifest |
| `KPICard.html` | HTML for the card |
| `KPICard.css` | Styles for the card |
| `KPICard.js` | All logic: data, rendering, settings |
| `tableau.extensions.1.latest.min.js` | Tableau Extensions API (bundled) |

## Setup

1. Host `KPICard.html`, `KPICard.css`, `KPICard.js`, and `tableau.extensions.1.latest.min.js` on any HTTPS server.
2. Update the `<url>` in `KPICard.trex` to your hosted URL.
3. In Tableau Desktop, use a **Viz Extension** mark type and load the `.trex` file.
4. Drag measures onto the **Value**, **Goal**, **Date**, and **Secondary Goal** encoding tiles.
5. Click the gear icon to open the settings panel.

## Local Development

```bash
npx http-server -p 8765 -c-1
```

Then load `KPICard.trex` in Tableau Desktop pointing to `http://localhost:8765/KPICard.html`.

## Hosting (GitHub Pages)

1. Push files to a GitHub repo.
2. Enable Pages under Settings > Pages (deploy from main branch).
3. Update the URL in `KPICard.trex` to `https://<user>.github.io/<repo>/KPICard.html`.

## Security

- No data leaves Tableau. All processing is client-side.
- The hosted files contain only rendering logic — no business data, no credentials.
- Tableau Server admins must allowlist the extension URL for published dashboards.
