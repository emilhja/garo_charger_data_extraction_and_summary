# GARO Charger Billing Workflow

This repository contains four small Python scripts that together handle the yearly charger billing workflow: fetch GARO usage, build the Excel summary workbook, import invoice kWh totals, and flag anomalies for review.
It was built in part because the standard GARO interface for these older chargers only allowed data export as PDF, which made repeatable extraction and billing work unnecessarily manual. Additionally it extracts kWh consumed from electricity invoices.

The repository is intentionally script-based rather than packaged as an application. Each script has one clear responsibility, and the normal workflow is to run them in sequence.

## Workflow At A Glance

The billing flow is:

1. `fetch_garo_energy.py`
   Pull daily GARO charger readings from the device API into `data/energy_2026.json` and `data/energy_2026.csv`.
2. `build_energy_summary_workbook.py`
   Combine the GARO export, garage metadata, and the pricing workbook into `energy_2026_summary.xlsx`.
3. `extract_invoice_kwh.py`
   Read Elhandel invoice PDFs and write monthly `Nedre` totals into the `Grunddata` sheet.
4. `check_anomalies.py`
   Review missing months, zero-usage months, and spikes before finalising billing.

If you are new to the repo, start by reading `garage_config.example.json`, then the four scripts in the order above.

## Scripts

- `fetch_garo_energy.py`
  Downloads daily charger usage from the GARO REST API and writes `data/energy_2026.json` and `data/energy_2026.csv`.
- `build_energy_summary_workbook.py`
  Builds `energy_2026_summary.xlsx` from `data/energy_2026.json`, `garage_config.json`, and the pricing workbook in `2026/Förbrukning/`. Runs anomaly checks at the end (see `check_anomalies.py`).
- `extract_invoice_kwh.py`
  Scans invoice PDFs under `2026/Fakturor/` and writes extracted monthly `Nedre` kWh values into the `Grunddata` sheet in `energy_2026_summary.xlsx`. `Övre` invoice totals are printed for cross-checking only because the workbook derives `Övre` from garage-sheet formulas.
- `check_anomalies.py`
  Scans `data/energy_2026.json` for missing months, zero readings, and usage spikes. Can be run standalone or is called automatically by `build_energy_summary_workbook.py`.

## Project Layout

- `fetch_garo_energy.py`
  API fetch script. Safe to rerun; it replaces the generated JSON and CSV exports.
- `build_energy_summary_workbook.py`
  Workbook generator. Safe to rerun; it rebuilds `energy_2026_summary.xlsx` from source inputs.
- `extract_invoice_kwh.py`
  Invoice importer. Safe to rerun; it updates the `Grunddata` sheet in the workbook.
- `check_anomalies.py`
  Validation helper used both standalone and from the workbook builder.
- `2026/Avläsningsrapporter/`
  Year-specific source material.
- `2026/Fakturor/`
  Invoice PDFs used for manual/imported monthly kWh totals.
- `2026/Förbrukning/`
  Pricing workbook used as the basis for `Grunddata`.
- `garage_config.json`
  Garage metadata, apartment numbers, ownership, and garage grouping.
- `garage_config.example.json`
  Sanitized example config for public sharing. Copy it to `garage_config.json` and replace values locally.
- `.env`
  Local environment variables such as the GARO API endpoint.

## Setup

Activate the virtual environment before running the scripts:

```bash
source venv/bin/activate
```

Copy the example environment file if needed:

```bash
cp .env.example .env
```

Copy the example garage config if needed:

```bash
cp garage_config.example.json garage_config.json
```

Install dependencies if needed:

```bash
pip install -r requirements.txt
```

## Workflow

1. Fetch charger usage from the GARO device:

```bash
python fetch_garo_energy.py
```

This writes:

- `data/energy_2026.json`
- `data/energy_2026.csv`

2. Build the Excel summary workbook:

```bash
python build_energy_summary_workbook.py
```

This writes:

- `energy_2026_summary.xlsx`

3. Import invoice kWh totals into `Grunddata`:

```bash
python extract_invoice_kwh.py
```

This updates `Grunddata` for `Nedre` only. `Övre` invoice values are reported in the console for validation but are not written back into the workbook.

## Data Sources

- GARO API: configured in `.env` as `GARO_API_BASE_URL`
- Pricing workbook: `2026/Förbrukning/Priser och Förbrukningsuppgifter laddstationer 2026.xlsx`
- Garage metadata: `garage_config.json`
- Invoice folders:
  - `2026/Fakturor/Övre/<MM - Mon>/`
  - `2026/Fakturor/Nedre/<MM - Mon>/`

## Important Assumptions

- `Övre` garage usage comes from the GARO API and is summed from the per-garage sheets.
- `Nedre` usage comes from invoice PDFs and is written into `Grunddata` column C.
- `extract_invoice_kwh.py` only trusts Elhandel invoices. Elnät invoices are ignored.
- The pricing workbook must contain a sheet named `Grunddata`.
- `garage_config.json` is the authoritative mapping for garage ownership, apartment numbers, and grouping.

## Validation

After running the scripts, check:

- `data/energy_2026.json` and `data/energy_2026.csv` after `python fetch_garo_energy.py`
- formulas, monthly totals, and `Grunddata` values in `energy_2026_summary.xlsx`
- imported invoice values after `python extract_invoice_kwh.py`
- `Övre` invoice totals printed by `extract_invoice_kwh.py` against the formula-driven `Grunddata` values

## Safety

- Generated files are overwritten when the scripts are rerun, so review workbook and export changes before sharing them.
- `.env`, `garage_config.json`, `data/`, and `2026/` may contain private addresses, ownership data, and billing data. Keep them out of public repositories.
- `fetch_garo_energy.py` talks directly to the GARO device on the local network. Do not point `GARO_API_BASE_URL` at another host unless you intend to query that device.
- `extract_invoice_kwh.py` raises an error if the same month/group contains conflicting invoice totals. That is a deliberate safety check and should be investigated, not bypassed.

## Development

Written together with [Claude Code](https://claude.ai/code).

## License

This project is licensed under the MIT License. See `LICENSE`.

## Notes

- `fetch_garo_energy.py` requires access to the local GARO device on the same network.
- The GARO endpoint is read from `.env`, so update `GARO_API_BASE_URL` there if the device address changes.
- Keep `.env`, `garage_config.json`, `data/`, and `2026/` out of public repositories. This repo includes `garage_config.example.json` for safe sharing.
- `build_energy_summary_workbook.py` keeps all configured garages in the workbook even if some have no usage yet.
- `extract_invoice_kwh.py` only uses Elhandel invoices and raises an error if conflicting duplicate invoices are found for the same month/group.
- `extract_invoice_kwh.py` expects invoice PDFs in month subfolders such as `2026/Fakturor/Övre/01 - Jan/`.
