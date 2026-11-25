# Brazil TE — database & golden template analysis (OCT 2026 / P10 UK)

Inputs analysed:

- `databases/` — 4 Power BI exports (`Detalhe True Efficiency - APP`, `- ITA`, `Detalle True Efficiency  - MES`, `Detalle True Efficiency`)
- `templates/TE_blank_templates/` — 8 blank templates
- `templates/TE_completed_templates/` — 8 golden templates (passed Measure Up QC)

Reproduce the numbers below with `python analysis/compare_db_vs_golden.py`.

---

## 1. Headline

1. **The template is the same as Mexico's.** Same sheets, same 8-column blocks, same row layout (Default / TC +1 row / PC +10 rows). Brazil also uses the **`Stax`** sheet (Sorocaba Canister), which Mexico never used.
2. **The databases are completely different from Mexico's.** There is no `Reporte Detalle Plantas` or `Pond` sheet and no `.xlsm`. Instead there are four flat Power BI exports:
   - three **downtime pivots** (rows = loss category, columns = line, values = hours), one per source system: **APP**, **ITA**, **MES**;
   - one **production list** (Plant, Line, Cut, Flavor, kg).
   So the reading half of `engine.py` has to be rewritten. The writing half (cell layout, xlwings, OneDrive staging, mtime check) still applies.
3. **The line ↔ template-block mapping and the category ↔ cell mapping are confirmed** (tables §3, §4). SKU counts per line match the goldens almost everywhere.
4. **The DB exports don't cover the same period as the goldens.** Every export is filtered *"Data is not 10/27 … 10/31/2026"* (26 days = 624 h), while the goldens are the full month (31 days = 744 h). Golden values are mostly the DB value or larger. Production is ≈ DB × 31/26. Only **217 of 465** downtime values match exactly. Re-exporting full October should close most of the gap.
5. **MES lines don't have the detail the template needs** (see §5). The **ITU** MES lines don't match the goldens at all (0 exact), so they were filled from some other source.
6. **Product master data is not in any DB.** That covers SKU name as written in the template, product type, % seasoning / oil, and rated throughput. We need a per-line SKU master (proposal in §6).

---

## 2. Database structure

| File | Rows | Columns | Value | Plants covered |
| --- | --- | --- | --- | --- |
| `Detalhe True Efficiency - APP.xlsx` | 29 loss categories (`DWT_GLOBAL_CLASSIFICACAO`) | 34 lines + Total | `TEMPO_DBL` hours | everything not in MES (filter *En MES is No*) |
| `Detalhe True Efficiency - ITA.xlsx` | 14 categories | FOFURA, TORCIDA | `HORAS ITA` hours | Lucky (planta ITA) |
| `Detalle True Efficiency  - MES.xlsx` | 18 categories (`Reason2`) — **coarser** | 9 lines | `DT Horas` hours | Sete Lagoas, ITU (filter *En App is No*) |
| `Detalle True Efficiency.xlsx` | 1 row per Plant × Line × Corte × Sabor | Planta, Linha, Corte, Sabor, Produção Real KG | kg | all plants |

Notes:

- Category names are already **English, TE-standard** ("Planned sanitation hours", "Dark hour Loss Tree - Staffling", …), so no Portuguese header translation is needed. MES uses slightly different labels (`Dark hrs: …`, `Other Hours`).
- APP also has columns `FOFURA` (4 992 h = 8 × 624) and `TORCIDA` (2 496 h = 4 × 624). The Lucky golden uses the **ITA** numbers, so the APP FOFURA/TORCIDA columns must be ignored (or they belong to something else — to confirm).
- `Productive time` and `Total` are not written anywhere; the template calculates them.
- Plant codes in the production file: `CIC, FSA, ITA, ITU, PET, REC, SOR, STL`.

## 3. Plant / line → template block (confirmed)

| Template file | Sheet / block | Header in golden | Source | DB line |
| --- | --- | --- | --- | --- |
| Amacoco Petrolina (PET) | Other 0/1/2 | 1 - TBA 19 / 2 - TBA 21 / 3 - TBA 22 | APP | TBA19 / TBA21 / TBA22 |
| CIC | PC 0 | 1 - PC60 | APP | PC60 |
| CIC | TC 0 | 2 - TC | APP | DORITOS_TC2 |
| CIC | Extruded 0 / 1 | 4 - Clextral / 8 - Clextral 7 | APP | CLEXTRAL III / CLEXTRAL VII |
| CIC | Other 0/1/2/3 | 3 - Stiksy / 5 - Pingo / 6 - Bacon / 7 - Popcorners | APP | STIKSY / PINGO CIC / BACON CIC / POPCORNERS |
| Feira de Santana (FSA) | Other 0/1/2/3/4 | 3 - BOSH / 1 - PLF / 2 - MASIPACK 1 / 4 - SIG 1 / 5 - SIG 2 | APP | BOSCH / PLF / MASIPACK / SIG 1 / SIG 2 |
| ITU | PC 0 | 1 - PC 50 | **MES** | PC50 |
| ITU | PC 1 | 8 - KETTLE ITU | APP | KETTLE (no production rows in DB) |
| ITU | TC 0 | 2 - TC | **MES** | DORITOS_TC1 |
| ITU | Extruded 0 / 1 | 3 - Clextral 1 / 4 - Clextral 2 | **MES** | CLEXTRAL I / CLEXTRAL II |
| ITU | Other 0 / 1 | 9 - Crunchy ITU / 5 - Pingo | APP | CRUNCHY ITU / PINGO ITU |
| Lucky (ITA) | Extruded 0 | 3 - FOFURA | **ITA** | FOFURA |
| Lucky (ITA) | Other 0 | 2 - TORCIDA | **ITA** | TORCIDA |
| NE Recife (REC) | Extruded 0/1/2 | 2 - Clextral / 1 - SCHAAF / 4 - SCHAAF 2 | APP | CLEXTRAL VI / SCHAAF_1 / SCHAAF_2 |
| NE Recife (REC) | Other 0 | 3 - Torcida Recife | APP | TORCIDA REC |
| Sete Lagoas (STL) | PC 0 | 1 - PC 37 | **MES** | PC37 |
| Sete Lagoas (STL) | TC 0 / 1 | 2 - TC / 5 - TC-04 | **MES** | Doritos 3 / Doritos 4 |
| Sete Lagoas (STL) | Extruded 0 / 1 | 3 - Clextral / 4 - Clextral 5 | **MES** | CLX IV / CLX V |
| Sorocaba (SOR) | Extruded 0 | 3 - SCHAAF | APP | SCHAAF_3 |
| Sorocaba (SOR) | Other 0/1/2/3 | 10 - OPTIMA 1 SOR / 11 - OPTIMA 2 SOR / 12 - BOSCH SOR / 13 - SYNTHEGON 9 SOR | APP | **TODDY_2 / TODDY_1 / TODDY_3 / TODDY_4** |
| Sorocaba (SOR) | Other 4/5/6 | 7 - L02 / 8 - L03 / 9 - L04 | APP | SPEED_2 / SPEED_3 / SPEED_4 |
| Sorocaba (SOR) | **Stax** 0 | 5 - Canister | APP | CANISTER_1 |

Line-type cell (row 7 col C, the `type_value` in `Job`) is hand-picked in the goldens: Extruded = `Baked`; Other = `Other` / `Fried Wheat` / `Pellets` / `Popcorn` / `Fritos`; Stax = `STAX750`. PC/TC are formulas, so they are left alone.

## 4. Loss category → template cell (Default layout; TC = row +1, PC = row +10)

| DB category (APP / ITA) | MES equivalent | Cell |
| --- | --- | --- |
| *(days in month × 24)* | — | **C18** Total calendar time (744) |
| Dark hours: line available to run if needed | Dark hrs: line available to run if needed | C19 |
| Planned maintenance hours | same | C22 |
| Planned sanitation hours | same | C23 |
| Planned startup & shutdown hours | same | C24 |
| Planned changeovers hours | same | C25 |
| Planned network outages hours | same | C26 |
| Team meetings hours | — | C27 |
| Other scheduled downtime hours | — | C28 |
| Maintenance overrun hours | same | C30 |
| Sanitation overrun hours | same | C31 |
| Changeover overrun hours | same | C32 |
| Raw material shortages hours | same | C35 |
| Quality issues hours | same | C36 |
| Other hours | Other Hours | C39 |
| Dark hour Loss Tree - RM shortage | ⚠ inside *Dark hrs: unavailable due to constraints* | G20 |
| Dark hour Loss Tree - FG warehouse | ⚠ same | G21 |
| Dark hour Loss Tree - Holiday & Religious reasons | ⚠ same | G22 |
| Dark hour Loss Tree - CAPEX & Test (R&D) | ⚠ same | G23 |
| Dark hour Loss Tree - Staffling | ⚠ same | G24 |
| Dark hour Loss Tree - Inventory | ⚠ same | G27 |
| Dark hour Loss Tree - Other | ⚠ same | G28 |
| Dark hour Loss Tree - Corn Cook Hours | — | G30 (**TC only**, extra row) |
| Processing Breakdown - Technician Failure | ⚠ *Processing Breakdown* (no split) | G32 |
| Processing Breakdown - Operational Failure | ⚠ same | G33 |
| Packaging Breakdown - Technician Failure | *Packaging Breakdown* → goldens put it all here | G36 |
| Packaging Breakdown - Operational Failure | — | G37 |
| Utilities - Internal | — | G40 |
| Utilities - External | *Utilities hours* → goldens put it all here | G41 |
| — | FG warehouse (unscheduled) | ? (probably E32/E33 "FG warehouse" concurrent box; unfilled in goldens) |
| — | Unspecified | ? (not visible in goldens) |

Product / production cells (as in Mexico):

| Sheet | Product table (first row) | Production column (first row) |
| --- | --- | --- |
| Default (Extruded / Other / Stax / Cookies / Wafer) | B44 name, C44 type, D44 % seasoning, F44, G44 rated throughput | C60 kg |
| TC | B45 name, C45 type, D45 % seasoning, E45 % oil, F45 non-IPS max throughput | C69 kg |
| PC | B54 name, C54 type, D54 % seasoning, E54 % oil, F54 non-IPS throughput | C88 kg |

Cells filled by hand in the goldens that are **not in any DB**: TE target (G80 / TC G97 / PC G126), rate/T-put losses (G106–G111), PC fryer params (C10 wash water 20, C11 moisture 0.015, C12 surface water 0.08, C17 gross solids), and the "# FG SKUs" cells (TC G124/G127, PC G156).

## 5. Gaps & open questions for the Brazil team

1. **Period.** Can the four exports be re-run for the **full fiscal month** that the template covers (OCT 2026 = 744 h)? With the current *excluding 10/27–10/31* filter, the automation can't reproduce the goldens.
2. **MES detail (Sete Lagoas, ITU PC50 / TC1 / Clextral I & II).** The MES export has a single *"Dark hrs: unavailable due to constraints"* bucket. The goldens split it into Holiday / CAPEX / Staffing / Other (e.g. Sete Lagoas Clextral: 18.86 / 1.70 / 130 / 52.69). Processing Breakdown likewise isn't split into Technical / Operational. Is there an MES export with the `Dark hour Loss Tree - …` and Tech/Op breakdown, like APP has? If not, what rule do analysts use (e.g. "all into Other / Technical")?
3. **ITU MES lines** match **0** golden values. For example, Clextral 1 sanitation is 64.06 h in MES but 0.08 h in the golden. Where did the ITU golden numbers come from?
4. **MES "FG warehouse" and "Unspecified"** — where should they go (if anywhere)?
5. **SKU master data** — see §6. Who owns it, and how often does it change?
6. **KETTLE (ITU)** has downtime in APP but **no production rows** in `Detalle True Efficiency.xlsx`, while the golden has 2 SKUs (8 013.82 / 1 240.32 kg).
7. **APP FOFURA / TORCIDA columns** (8× and 4× 624 h) — confirm they should be ignored in favour of ITA.
8. **Manual cells** (TE targets, rate losses, PC fryer params, line types). Should they be fixed values in config, carried over from the previous month's template, or left for the analyst?

## 6. SKU master — the one piece the DB can't supply

The production file identifies a SKU by **Corte + Sabor** in mixed Spanish/Portuguese (`CORTE ACANALADO` + `SAL`). The template wants the commercial name (`RUFFLES SAL`) plus product type, % seasoning and rated throughput. The goldens show that each line has a fixed SKU list (often with SKUs that had no production that month).

Proposal: keep a `sku_master` per line in `plants.py` (or a CSV the analysts can edit), e.g.

```python
# line PC60 (CIC)
SKU("CORTE ACANALADO", "SAL", name="RUFFLES SAL", ptype="Wavy", seasoning=0.0121, oil=0, tput=0),
```

The engine then writes the master rows into the product table in a fixed order. Each SKU's production goes on the matching row, and SKUs with no production stay at 0 / blank, as in the goldens. I can generate a first draft of this master from the OCT goldens; the Brazil team then corrects it.

## 7. What changes in `modular_architecture/` (implemented — see CLAUDE.md §4)

| Module | Change |
| --- | --- |
| `engine.py` — readers | Replace `_rdp_header_map` / `_load_pond_df` / `pond_*` with a reader for the 4 exports. Plain pandas/openpyxl, no COM needed: build `{source: {line: {category: hours}}}` and `{line: [(corte, sabor, kg)]}`. |
| `engine.py` — constants | Replace `MANDATORY_HEADERS` / `OPTIONAL_HEADERS` / `ROWS_*` with one `CATEGORY_CELLS` table (§4) plus the kind offsets (TC +1, PC +10). Add `"Stax"` to `_PRODUCT_SHEETS`. Calendar hours = days in month × 24 (from the filename period or a UI field). |
| `engine.py` — writers | `_write_header` is unchanged. New `_write_hours(ws, kind, cons, cats)` and `_write_products(ws, kind, cons, master, production)`. Mexico's Pond-specific SKU logic (aggregation, PC top-table lookup) goes away. |
| `engine.py` — runner | `fill_plant` / OneDrive staging / mtime check unchanged. `_fill_plant_core` opens only the template in Excel; the DBs are read up front by pandas. |
| `plants.py` | Brazil registry from §3 (8 plants, 43 blocks). `Job` gets `source` ("APP"/"ITA"/"MES") + `db_line` in place of `rdp_row` / `pond_plant` / `pond_line`. `match_filename()` for `Snk_TE_<Plant>_P11 2026 OCT(P10 UK).xlsm`. |
| `app.py` | Branding (title, "Brazil Sites"). The DB picker selects the **databases folder** (4 files) instead of one `.xlsm`. |
| `build.spec` | `name="TE_Filler_Brazil"`. |
