"""Download daily GARO charger energy data and export yearly JSON/CSV files."""

import csv
import json
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import requests

YEAR = 2026
DATA_DIR = Path("data")
ENV_FILE = Path(".env")
GARO_API_BASE_URL_ENV = "GARO_API_BASE_URL"
CSV_FIELDNAMES = [
    "serial",
    "name",
    "meter_serial",
    "date",
    "energy_kwh",
    "meter_start_wh",
    "meter_stop_wh",
]


def load_dotenv_file(env_file: Path = ENV_FILE) -> None:
    """Load simple `KEY=VALUE` pairs from a dotenv file into the environment."""
    if not env_file.exists():
        return

    for raw_line in env_file.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip("'\""))


def get_garo_api_base_url() -> str:
    """Return the configured GARO API base URL or raise a clear configuration error."""
    load_dotenv_file()
    base_url = os.getenv(GARO_API_BASE_URL_ENV)
    if not base_url:
        raise KeyError(
            f"Missing {GARO_API_BASE_URL_ENV} in the environment or {ENV_FILE}."
        )
    return base_url


def fetch_chargebox_config(base_url: str) -> dict[str, Any]:
    """Fetch the GARO chargebox configuration used to map serials to garage names."""
    response = requests.get(f"{base_url}/config", timeout=10)
    response.raise_for_status()
    return response.json()


def fetch_monthly_energy_data(
    base_url: str,
    chargebox_serial: int | str,
    year: int,
    month: int,
) -> dict[str, Any] | None:
    """Fetch daily energy readings for one charger and month.

    Returns `None` when the GARO API signals that the requested month has no data.
    """
    request_payload = {
        "chargeboxSerial": int(chargebox_serial),
        "year": year,
        "month": month,
        "meterSerial": "DEFAULT",
        "resolution": "DAY",
    }
    response = requests.post(f"{base_url}/energy", json=request_payload, timeout=10)
    if response.status_code in {204, 404}:
        return None
    response.raise_for_status()
    return response.json()


def get_last_month_to_fetch(target_year: int) -> int:
    """Return the latest month that should be fetched for the configured export year."""
    current_date = datetime.now(timezone.utc)
    if target_year < current_date.year:
        return 12
    if target_year == current_date.year:
        return current_date.month
    raise ValueError(
        f"Configured YEAR={target_year} is in the future relative to the current year {current_date.year}."
    )


def build_daily_energy_rows() -> list[dict[str, Any]]:
    """Collect one export row per charger-day for the configured year."""
    base_url = get_garo_api_base_url()
    chargebox_config = fetch_chargebox_config(base_url)
    chargebox_serials = chargebox_config.get("energySerials", [])
    if not chargebox_serials:
        raise KeyError("No `energySerials` found in the GARO `/config` response.")

    serial_to_reference = {
        str(slave["serialNumber"]): slave.get("reference", str(slave["serialNumber"]))
        for slave in chargebox_config.get("slaveList", [])
        if "serialNumber" in slave
    }

    energy_rows: list[dict[str, Any]] = []
    last_month_to_fetch = get_last_month_to_fetch(YEAR)
    for chargebox_serial in chargebox_serials:
        garage_name = serial_to_reference.get(str(chargebox_serial), str(chargebox_serial))
        print(f"Fetching {garage_name} (serial {chargebox_serial})...")
        for month in range(1, last_month_to_fetch + 1):
            month_data = fetch_monthly_energy_data(base_url, chargebox_serial, YEAR, month)
            if month_data is None:
                print(f"  {YEAR}-{month:02d}: no data")
                continue

            timestamps = month_data.get("timestamps", [])
            energy_values = month_data.get("values", [])
            meter_start_values = month_data.get("start", [])
            meter_stop_values = month_data.get("stop", [])
            meter_serial = month_data.get("meterSerial", "")

            if len(timestamps) != len(energy_values):
                raise ValueError(
                    f"Mismatched timestamp/value lengths for serial {chargebox_serial}, "
                    f"{YEAR}-{month:02d}: {len(timestamps)} timestamps, {len(energy_values)} values."
                )

            for index, timestamp_ms in enumerate(timestamps):
                reading_date = datetime.fromtimestamp(
                    timestamp_ms / 1000,
                    tz=timezone.utc,
                ).strftime("%Y-%m-%d")
                energy_rows.append(
                    {
                        "serial": chargebox_serial,
                        "name": garage_name,
                        "meter_serial": meter_serial,
                        "date": reading_date,
                        "energy_kwh": round(energy_values[index], 4) if index < len(energy_values) else None,
                        "meter_start_wh": meter_start_values[index] if index < len(meter_start_values) else None,
                        "meter_stop_wh": meter_stop_values[index] if index < len(meter_stop_values) else None,
                    }
                )

            month_total_kwh = sum(value for value in energy_values if value is not None)
            print(f"  {YEAR}-{month:02d}: {month_total_kwh:.2f} kWh ({len(timestamps)} days)")

    return energy_rows


def write_output_files(energy_rows: list[dict[str, Any]]) -> None:
    """Write the yearly JSON and CSV exports used by the workbook generator."""
    DATA_DIR.mkdir(exist_ok=True)
    json_output_path = DATA_DIR / f"energy_{YEAR}.json"
    csv_output_path = DATA_DIR / f"energy_{YEAR}.csv"

    with json_output_path.open("w", encoding="utf-8") as json_file:
        json.dump(energy_rows, json_file, indent=2)
    print(f"\nSaved {len(energy_rows)} records to {json_output_path}")

    with csv_output_path.open("w", newline="", encoding="utf-8") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=CSV_FIELDNAMES)
        writer.writeheader()
        writer.writerows(energy_rows)
    print(f"Saved {len(energy_rows)} records to {csv_output_path}")


def main() -> None:
    """Run the yearly GARO export workflow from API fetch to file output."""
    energy_rows = build_daily_energy_rows()
    write_output_files(energy_rows)


if __name__ == "__main__":
    main()
