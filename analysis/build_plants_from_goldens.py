"""Generate modular_architecture/plants.py from the golden templates.

    python analysis/build_plants_from_goldens.py            # writes plants.py
    python analysis/build_plants_from_goldens.py --check    # only report unmatched SKUs

One-off bootstrap: it reads each golden block (header, line type, TE target,
PC fryer params, product table) and attaches the DB (Corte, Sabor) keys that
feed each product row, using ALIASES / LINE_ALIASES below. After that,
plants.py is the source of truth and analysts edit it directly. Re-run only
to rebuild the registry from a new set of goldens (hand edits are lost).
"""
from __future__ import annotations

import os
import sys
import unicodedata

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
GOLDEN = os.path.join(ROOT, "templates", "TE_completed_templates")
DBDIR = os.path.join(ROOT, "databases")
OUT = os.path.join(ROOT, "modular_architecture", "plants.py")


def norm(s):
    s = "".join(c for c in unicodedata.normalize("NFKD", str(s or "")) if not unicodedata.combining(c))
    return " ".join(s.upper().split())


# ---------------------------------------------------------------------------
# Plants: (REGISTRY key, filename keyword, [(sheet, cons, source, DB line), ...])
# ---------------------------------------------------------------------------
PLANTS = [
    ("AMACOCO_PETROLINA", "amacoco", [
        ("Other", 0, "APP", "TBA19"), ("Other", 1, "APP", "TBA21"), ("Other", 2, "APP", "TBA22")]),
    ("CIC", "_cic_", [
        ("PC", 0, "APP", "PC60"), ("TC", 0, "APP", "DORITOS_TC2"),
        ("Extruded", 0, "APP", "CLEXTRAL III"), ("Extruded", 1, "APP", "CLEXTRAL VII"),
        ("Other", 0, "APP", "STIKSY"), ("Other", 1, "APP", "PINGO CIC"),
        ("Other", 2, "APP", "BACON CIC"), ("Other", 3, "APP", "POPCORNERS")]),
    ("FEIRA_DE_SANTANA", "feira de santana", [
        ("Other", 0, "APP", "BOSCH"), ("Other", 1, "APP", "PLF"), ("Other", 2, "APP", "MASIPACK"),
        ("Other", 3, "APP", "SIG 1"), ("Other", 4, "APP", "SIG 2")]),
    ("ITU", "_itu_", [
        ("PC", 0, "MES", "PC50"), ("PC", 1, "APP", "KETTLE"), ("TC", 0, "MES", "DORITOS_TC1"),
        ("Extruded", 0, "MES", "CLEXTRAL I"), ("Extruded", 1, "MES", "CLEXTRAL II"),
        ("Other", 0, "APP", "CRUNCHY ITU"), ("Other", 1, "APP", "PINGO ITU")]),
    ("LUCKY", "_lucky_", [
        ("Extruded", 0, "ITA", "FOFURA"), ("Other", 0, "ITA", "TORCIDA")]),
    ("NE_RECIFE", "ne recife", [
        ("Extruded", 0, "APP", "CLEXTRAL VI"), ("Extruded", 1, "APP", "SCHAAF_1"),
        ("Extruded", 2, "APP", "SCHAAF_2"), ("Other", 0, "APP", "TORCIDA REC")]),
    ("SETE_LAGOAS", "sete lagoas", [
        ("PC", 0, "MES", "PC37"), ("TC", 0, "MES", "Doritos 3"), ("TC", 1, "MES", "Doritos 4"),
        ("Extruded", 0, "MES", "CLX IV"), ("Extruded", 1, "MES", "CLX V")]),
    ("SOROCABA", "sorocaba", [
        ("Extruded", 0, "APP", "SCHAAF_3"),
        ("Other", 0, "APP", "TODDY_2"), ("Other", 1, "APP", "TODDY_1"),
        ("Other", 2, "APP", "TODDY_3"), ("Other", 3, "APP", "TODDY_4"),
        ("Other", 4, "APP", "SPEED_2"), ("Other", 5, "APP", "SPEED_3"), ("Other", 6, "APP", "SPEED_4"),
        ("Stax", 0, "APP", "CANISTER_1")]),
]

# ---------------------------------------------------------------------------
# (Corte, Sabor) -> template product names it can feed (first one present on the line wins)
# ---------------------------------------------------------------------------
ALIASES = {
    # Potato chips (PC)
    ("CORTE ACANALADO", "SAL"): ["RUFFLES SAL"],
    ("CORTE ACANALADO", "CHURRASCO"): ["RUFFLES CHURRASCO"],
    ("CORTE ACANALADO", "CEBOLA E SALSA"): ["RUFFLES CEBOLA E SALSA"],
    ("CORTE ACANALADO", "PIMIENTA/PIMENTA"): ["RUFFLES PIMENTA MEXICANA"],
    ("LAYS LISA", "SAL"): ["LAYS CLASSICA"],
    ("LAYS LISA", "VINAGRE"): ["LAYS VINEGAR"],
    ("LAYS LISA", "CREMA AGRIA"): ["LAYS SOUR CREAM"],
    ("LAYS LISA", "PICANHA"): ["LAYS PICANHA"],
    ("LAYS LISA", "QUESO"): ["LAYS QUEIJO"],
    ("PALHA", "SAL"): ["NA MESA", "PALHA"],
    ("PALHA EXTRA FINA", "SAL"): ["NA MESA EXTRA FINA", "PALHA EXTRA FINA"],
    ("SENSACOES", "PEITO DE PERU"): ["SENSACOES PEITO DE PERU", "SENSACOES PERU"],
    ("SENSACOES II", "PEITO DE PERU"): ["SENSACOES PEITO DE PERU", "SENSACOES PERU"],
    ("SENSACOES", "FRANGO GRELHADO"): ["SENSACOES FRANGO GRELHADO", "SENSACOES FRANGO"],
    ("SENSACOES II", "FRANGO GRELHADO"): ["SENSACOES FRANGO GRELHADO", "SENSACOES FRANGO"],
    ("STAX ACANALADO", "SAL"): ["RUFFLES TUBO SAL"],
    ("STAX ACANALADO", "CEBOLLA"): ["RUFFLES TUBO SOUR CREAM"],
    # Tortilla chips (TC)
    ("DORITOS", "QUESO"): ["DORITOS NACHO"],
    ("DORITOS", "PIZZA"): ["DORITOS PIZZA"],
    ("DORITOS", "KETCHUP"): ["DORITOS KETCHUP"],
    ("DORITOS", "SPICY"): ["DORITOS SPICY NACHO"],
    ("DORITOS", "SWEET CHILI"): ["DORITOS SWEET CHILI"],
    ("DORITOS DINAMITA", "FLAMIN HOT"): ["DINAMITAS FLAMIN HOT"],
    ("DORITOS DINAMITA", "PIMIENTA/PIMENTA"): ["DINAMITAS PIMENTA MEXICANA"],
    ("TOSTITOS REDONDO", "SAL"): ["TOSTITOS SAL"],
    ("TOSTITOS REDONDO", "PIMIENTA/PIMENTA"): ["TOSTITOS PIMENTA MEXICANA"],
    # Extruded
    ("CEBOLITOS", "CEBOLLA"): ["CEBOLITOS"],
    ("CHEETOS BOLA", "QUESO"): ["CHEETOS BOLA"],
    ("CHEETOS LUA", "QUESO PARMESANO"): ["CHEETOS LUA"],
    ("CHEETOS ONDA", "REQUEIJAO"): ["CHEETOS ONDA"],
    ("CHEETOS PATINHA", "QUESO"): ["INSUMOS PATINHA"],
    ("CHEETOS TRIANGULO", "QUESO"): ["INSUMOS TRIANGULO"],
    ("FANDANGOS", "PRESUNTO"): ["FANDANGOS PRESUNTO"],
    ("FANDANGOS", "QUESO"): ["FANDANGOS QUEIJO"],
    ("FANDANGOS", "CHURRASCO"): ["FANDANGOS CHURRASCO"],
    ("FANDANGOS", "BACON/TOCINO"): ["FANDANGOS BACON"],
    ("SKYNI", "NATURAL"): ["SKINY NATURAL"],
    ("SKYNI", "PRESUNTO"): ["SKINY PRESUNTO"],
    ("SKYNI", "QUESO"): ["SKINY QUEIJO"],
    ("ANEL", "CEBOLLA"): ["FOFURA CEBOLA"],
    ("ANEL", "CHURRASCO"): ["FOFURA CHURRASCO"],
    ("ANEL", "PRESUNTO"): ["FOFURA PRESUNTO"],
    ("ANEL", "QUESO"): ["FOFURA QUEIJO"],
    ("ANEL", "REQUEIJAO"): ["FOFURA REQUEIJAO"],
    # Other
    ("STIKSY", "SAL"): ["STIKSY"],
    ("PINGO REDONDO", "BACON/TOCINO"): ["PINGO BACON"],
    ("PINGO QUADRADO", "PICANHA"): ["PINGO PICANHA"],
    ("BACONZITOS", "BACON/TOCINO"): ["BACONZITOS"],
    ("POPCORN", "QUESO"): ["POPCORNERS CHEDDAR"],
    ("POPCORN", "KETTLE CORN"): ["POPCORNERS KETTLE"],
    ("POPCORN", "SAL"): ["POPCORNERS SAL"],
    ("CHEETOS KRUNCHY", "QUESO"): ["CRUNCHY WHITE CHEESE"],
    ("CHEETOS KRUNCHY", "PIMIENTA/PIMENTA"): ["CRUNCHY PIMENTA MEXICANA"],
    ("TORCIDA TOR", "AJO"): ["PAO DE ALHO"],
    ("TORCIDA TOR", "BACON/TOCINO"): ["BACON"],
    ("TORCIDA TOR", "CAMARON PIMIENTA"): ["TORCIDA CAMARAO"],
    ("TORCIDA TOR", "CEBOLLA"): ["CEBOLA"],
    ("TORCIDA TOR", "CHURRASCO"): ["CHURRASCO"],
    ("TORCIDA TOR", "COSTILLAR"): ["COSTELINHA"],
    ("TORCIDA TOR", "PIMIENTA/PIMENTA"): ["PIMENTA MEXICANA"],
    ("TORCIDA TOR", "QUESO"): ["QUEIJO"],
    ("TORCIDA TOR", "VINAGRETE"): ["VINAGRETE"],
    ("CARTONADO", "3 PACK CHOCOLATE"): ["TODDYNHO 9X3X200"],
    ("CARTONADO", "3 PACK LIGHT"): ["TODDYNHO LEVINHO 200ML"],
    # Toddy powder (sachet lines; pote lines override below)
    ("CHOCOLATE 700", "CHOCOLATE"): ["TODDY SACHET 700G"],
    ("CHOCOLATE 300", "CHOCOLATE"): ["TODDY SACHET 300G", "TODDY ORIGINAL SACHET 300GRX24"],
    ("CHOCOLATE 1020", "CHOCOLATE"): ["TODDY ORIGINAL SACHET 1,02KGX12"],
    ("CHOCOLATE 560", "CHOCOLATE"): ["TODDY ORIGINAL SACHET 560GRX16", "TODDY ORIGINAL SACHET 560GRX24"],
    ("CHOCOLATE 480", "CHOCOLATE"): ["TODDY ORIGINAL SACHET 480GRX20"],
    ("CHOCOLATE 620", "CHOCOLATE"): ["TODDY ORIGINAL SACHET 620GRX12"],
    ("CHOCOLATE 175", "CHOCOLATE"): ["TODDY ORIGINAL SACHET 175GRX42"],
    ("CHOCOLATE 280", "CHOCOLATE LIGHT"): ["TODDY LIGHT 280G"],
    ("CHOCOLATE 2000", "CHOCOLATE"): ["TODDY ORIGINAL 6X1800"],
    ("TODDY LIGHT", "CHOCOLATE 380"): ["TODDY LIGHT 24X380"],
    ("CHOCOLATE 400", "CHOCOLATE"): ["TODDY ORIG POTE 370GX24"],
    ("CHOCOLATE 800", "CHOCOLATE"): ["TODDY ORIGINAL 750GRX12"],
}

# Line-specific overrides (where the same Corte/Sabor means a different template row)
LINE_ALIASES = {
    # Sorocaba OPTIMA 1/2 (pote lines): the goldens book everything on the EXP rows
    "TODDY_1": {("CHOCOLATE 400", "CHOCOLATE"): ["TODDY ORIGINAL EXP 370"],
                ("CHOCOLATE 800", "CHOCOLATE"): ["TODDY ORIGINAL EXP 750"],
                ("TODDY 200", "CHOCOLATE"): ["TODDY ORIGINAL EXP 200"]},
    "TODDY_2": {("CHOCOLATE 400", "CHOCOLATE"): ["TODDY ORIGINAL EXP 370"],
                ("CHOCOLATE 800", "CHOCOLATE"): ["TODDY ORIGINAL EXP 750"],
                ("TODDY 200", "CHOCOLATE"): ["TODDY ORIGINAL EXP 200"]},
}

SHEET_LAYOUT = {"PC": (6, 54, 30, 126), "TC": (4, 45, 20, 97)}  # header row, first product row, rows, TE target row
DEFAULT_LAYOUT = (4, 44, 12, 80)


def py(v):
    if isinstance(v, float):
        return repr(round(v, 6))
    return repr(v)


def main(check_only: bool) -> int:
    goldens = os.listdir(GOLDEN)
    sys.path.insert(0, os.path.join(ROOT, "modular_architecture"))
    import engine
    db = engine.load_brazil_db(DBDIR)

    out = [HEADER]
    problems = 0
    for key, kw, blocks in PLANTS:
        fn = next(f for f in goldens if kw in f.lower())
        wb = openpyxl.load_workbook(os.path.join(GOLDEN, fn), data_only=True)
        blank = openpyxl.load_workbook(os.path.join(ROOT, "templates", "TE_blank_templates",
                                                    fn.replace("_complete", "")))
        out.append(f"\n# {'-' * 75}\n# {key}  ({fn.replace('_complete', '')})\n# {'-' * 75}")
        out.append(f"{key} = PlantSpec(\n    key={key!r},\n    jobs=[")
        for sheet, cons, source, line in blocks:
            ws = wb[sheet]
            hrow, prow, nrows, trow = SHEET_LAYOUT.get(sheet, DEFAULT_LAYOUT)
            o = 8 * cons
            header = ws.cell(hrow, 2 + o).value
            ttype = ws.cell(7, 3 + o).value if sheet not in ("PC", "TC") else None
            if str(blank[sheet].cell(7, 3 + o).value or "").startswith("="):
                ttype = None  # line type is a formula in the blank (e.g. Stax) — leave it
            target = ws.cell(trow, 7 + o).value
            fryer = tuple(ws.cell(r, 3 + o).value for r in (10, 11, 12)) if sheet == "PC" else None
            rows = []
            for i in range(nrows):
                name = ws.cell(prow + i, 2 + o).value
                if name in (None, ""):
                    continue
                # skip columns that are formulas in the blank (their golden value is calculated)
                vals = [None if str(blank[sheet].cell(prow + i, c + o).value or "").startswith("=")
                        else ws.cell(prow + i, c + o).value for c in (3, 4, 5, 6, 7)]
                vals = [None if (v == "" or (isinstance(v, str) and v.startswith("="))) else v for v in vals]
                rows.append([str(name), vals, []])
            names = {norm(r[0]): r for r in rows}
            aliases = {**ALIASES, **LINE_ALIASES.get(line, {})}
            for corte, sabor, kg in db.production.get(norm(line), []):
                k = (norm(corte), norm(sabor))
                cands = aliases.get(k, [])
                target_row = next((names[norm(c)] for c in cands if norm(c) in names), None)
                if target_row is None and len(rows) == 1 and corte is None:
                    target_row = rows[0]  # single-SKU line without Corte/Sabor
                    continue
                if target_row is None:
                    problems += 1
                    print(f"UNMATCHED {key} {sheet}{cons} {line}: {corte} / {sabor} ({kg:,.0f} kg)")
                    continue
                if (corte, sabor) not in target_row[2]:
                    target_row[2].append((corte, sabor))
            # add the remaining aliases that point to rows on this line (future months)
            assigned = {(norm(a), norm(b)) for r in rows for a, b in r[2]}
            for k, cands in aliases.items():
                if k in assigned:
                    continue
                r = next((names[norm(c)] for c in cands if norm(c) in names), None)
                if r is not None:
                    r[2].append(k)
                    assigned.add(k)
            out.append(f"        # {line} ({source})")
            args = [py(sheet), str(cons), py(header), py(source), py(line)]
            if ttype is not None:
                args.append(f"type_value={py(ttype)}")
            if isinstance(target, (int, float)):
                args.append(f"te_target={py(target)}")
            if fryer and any(v is not None for v in fryer):
                args.append(f"fryer=({', '.join(py(v) for v in fryer)})")
            out.append(f"        Job({', '.join(args)}, skus=(")
            for name, vals, keys in rows:
                kv = ", ".join(f"{n}={py(v)}" for n, v in zip(("ptype", "seasoning", "oil", "col_f", "rated"), vals) if v is not None)
                kk = ", ".join(f"({py(a)}, {py(b)})" for a, b in keys)
                out.append(f"            SKU({py(name)}{', ' + kv if kv else ''}, db_keys=({kk}{',' if len(keys) == 1 else ''})),")
            out.append("        )),")
        out.append("    ],\n)")
    out.append(FOOTER.format(registry="\n".join(f"    {k}.key: {k}," for k, _, _ in PLANTS),
                             keywords="\n".join(f"    ({kw!r}, {k!r})," for k, kw, _ in PLANTS)))
    if problems:
        print(f"{problems} production rows have no matching SKU.")
    if not check_only:
        with open(OUT, "w", encoding="utf-8") as f:
            f.write("\n".join(out) + "\n")
        print(f"Wrote {OUT}")
    return 1 if problems else 0


HEADER = '''# plants.py — Brazil plant registry
#
# One PlantSpec per template file; one Job per line block in that template.
#   Job(sheet, cons, header_text, source, db_line, ...)
#     sheet       template sheet ("PC", "TC", "Extruded", "Other", "Stax", ...)
#     cons        block index on that sheet (0 = first block, columns B..I)
#     header_text line name exactly as in the template dropdown
#     source      downtime export: "APP", "ITA" or "MES"
#     db_line     'Linha' as it appears in the exports
#     type_value  line type (row 7) for Default-layout sheets
#     te_target   TE % target cell
#     fryer       PC only: (wash water temp, moisture aim, surface water)
#     skus        product table, top to bottom. Each SKU lists the
#                 (Corte, Sabor) pairs from the production export that feed it.
#
# Bootstrapped from the OCT 2026 goldens by analysis/build_plants_from_goldens.py.
# From here on, edit this file by hand: add a SKU row when a new product
# appears (the app warns "SKU ... is not in the SKU master").
from __future__ import annotations

import os as _os
import unicodedata as _ud
from typing import Dict

from engine import Job, PlantSpec, SKU
'''

FOOTER = '''

# -------------------------
# Filename matching
# -------------------------
# Keyword (lower-case, accents stripped) -> REGISTRY key. Brazilian templates
# are named like "Snk_TE_<Plant>_P11 2026 OCT(P10 UK).xlsm".
_FNAME_KEYWORDS = [
{keywords}
]


def _norm_fname(s: str) -> str:
    """Strip diacritics and lowercase a filename string for matching."""
    s = "".join(ch for ch in _ud.normalize("NFKD", s) if not _ud.combining(ch))
    return s.lower()


def match_filename(fname: str) -> "str | None":
    """Return the REGISTRY key for a template filename, or None if unrecognised."""
    n = _norm_fname(_os.path.basename(fname))
    for kw, key in _FNAME_KEYWORDS:
        if kw in n:
            return key
    return None


def find_templates(folder: str) -> "list[tuple[str, str]]":
    """Scan *folder* for .xlsm templates whose name matches a known plant.
    Returns (plant_key, absolute_filepath) pairs sorted by filename.
    Completed/filled copies ("_complete", "_Filled") are skipped."""
    results = []
    try:
        for fname in sorted(_os.listdir(folder)):
            low = fname.lower()
            if not low.endswith(".xlsm") or low.startswith("~$"):
                continue
            if "_complete" in low or "_filled" in low:
                continue
            key = match_filename(fname)
            if key:
                results.append((key, _os.path.join(folder, fname)))
    except OSError:
        pass
    return results


# -------------------------
# Registry
# -------------------------
REGISTRY: Dict[str, PlantSpec] = {{
{registry}
}}
'''

if __name__ == "__main__":
    sys.exit(main("--check" in sys.argv))
