# DoN Congressional Map — rough draft

Interactive map of 120th Congress districts with Department of the Navy installations.

## Update the data

Edit **`site/data/map_data.xlsx`**, save it, and refresh the page. You don't need to rebuild anything.

| Tab | What it holds |
|---|---|
| House | One row per district (`VA-2`, `AK-AL`). Name, Party, Status. |
| Senate | Two rows per state. |
| Installations | 387 Navy/USMC sites (226 active, 161 reserve). `Component` = Active or Reserve (reserve sites have their own layer). `Show = No` hides a row. |
| Poster Check (118th) | Every site on the 2023 NAVFAC poster, matched to the map, with its 118th and 120th districts. Rows shaded yellow need review. |
| Contractors | 131 private plants building ships, aircraft, weapons and systems for the Navy and USMC (orange dots). Each district panel lists the plants inside it. Locations are site-level estimates. |

Any column you **add** to the House or Senate tabs (Committees, Military Service, etc.) automatically shows up in that member's pop-up. The README tab inside the workbook has the full rules.

## Folders

- `site/` — the website. Upload this whole folder to any static host (GitHub Pages, Netlify, Cloudflare Pages).
- `site/geo/` — boundary files (districts, states, installation outlines). Generated once; you won't need to edit these.
- `prep/` — scripts that generated the boundary files and the first version of the spreadsheet (for reference).

## Data sources

- Districts: U.S. Census TIGER/Line 2026, 120th Congress (`cd120`), clipped to the 2025 shoreline.
- Installations: DoD MIRTA dataset, active Navy/USMC sites, plus Parris Island from Census TIGER.
- Election results: Wikipedia results pages as of Nov 9, 2026. **Verify these before publishing.**
- Optional street background: CARTO Positron tiles (free for low-volume use; check their terms before a high-traffic public launch).
