# GARO Charger Billing Workflow

This repository contains three standalone Python scripts for collecting GARO charger usage, building the yearly billing workbook, and importing invoice kWh totals.
It was built in part because the standard GARO interface for these older chargers only allowed data export as PDF, which made repeatable extraction and billing work unnecessarily manual. Additionally it extracts kWh consumed from electricity invoices.

## Scripts

- `fetch_garo_energy.py`
  Downloads daily charger usage from the GARO REST API and writes `data/energy_2026.json` and `data/energy_2026.csv`.
- `build_energy_summary_workbook.py`
  Builds `energy_2026_summary.xlsx` from `data/energy_2026.json`, `garage_config.json`, and the pricing workbook in `2026/Förbrukning/`. Runs anomaly checks at the end (see `check_anomalies.py`).
- `extract_invoice_kwh.py`
  Scans invoice PDFs under `2026/Fakturor/` and writes extracted monthly kWh values into the `Grunddata` sheet in `energy_2026_summary.xlsx`.
- `check_anomalies.py`
  Scans `data/energy_2026.json` for missing months, zero readings, and usage spikes. Can be run standalone or is called automatically by `build_energy_summary_workbook.py`.

## Project Layout

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

## Data Sources

- GARO API: configured in `.env` as `GARO_API_BASE_URL`
- Pricing workbook: `2026/Förbrukning/Priser och Förbrukningsuppgifter laddstationer 2026.xlsx`
- Invoice folders:
  - `2026/Fakturor/Övre/`
  - `2026/Fakturor/Nedre/`

## Validation

After running the scripts, check:

- `data/energy_2026.json` and `data/energy_2026.csv` after `python fetch_garo_energy.py`
- formulas, monthly totals, and `Grunddata` values in `energy_2026_summary.xlsx`
- imported invoice values after `python extract_invoice_kwh.py`

## Development

Written together with [Claude Code](https://claude.ai/code).

## Notes

- `fetch_garo_energy.py` requires access to the local GARO device on the same network.
- The GARO endpoint is read from `.env`, so update `GARO_API_BASE_URL` there if the device address changes.
- Keep `.env`, `garage_config.json`, `data/`, and `2026/` out of public repositories. This repo includes `garage_config.example.json` for safe sharing.
- `build_energy_summary_workbook.py` keeps all configured garages in the workbook even if some have no usage yet.
- `extract_invoice_kwh.py` only uses Elhandel invoices and raises an error if conflicting duplicate invoices are found for the same month/group.
