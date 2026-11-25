# TE-Automation-Brazil

Fills PepsiCo Brazil's True Efficiency (TE) reporting templates from the monthly True Efficiency exports.
Forked from the Mexico implementation ([TE-Automation](https://github.com/p3droJosue/TE-Automation)).
See `CLAUDE.md` for architecture and maintenance, and `analysis/BRAZIL_ANALYSIS.md` for how the data maps.

- `modular_architecture/` — the app (`app.py` UI, `engine.py` fill engine, `plants.py` plant registry + SKU master)
- `databases/` — the four monthly exports (APP, ITA, MES downtime + production)
- `templates/TE_blank_templates/` — blank templates (input)
- `templates/TE_completed_templates/` — completed "golden" templates to compare against
- `analysis/` — DB vs golden checks (`check_against_golden.py` runs without Excel)

## Quick start (Windows + Excel)

```powershell
git clone https://github.com/p3droJosue/TE-Automation-Brazil.git
cd TE-Automation-Brazil
git checkout claude/eager-galileo-fbakkz
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install customtkinter xlwings openpyxl pillow pyinstaller
cd modular_architecture
python app.py
```

In the **Run All** tab pick `..\databases` as the databases folder and `..\templates\TE_blank_templates` as the templates folder.
Each plant is saved as `<template>_Filled.xlsm` next to its blank (blanks are never overwritten); warnings go to `<template>_Filled_warnings.txt`.
