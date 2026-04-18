"""Download daily GARO charger energy data and export yearly JSON/CSV files."""

import csv
import json
import os
import unicodedata
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import urlparse
from zoneinfo import ZoneInfo

import requests

YEAR = 2026
DATA_DIR = Path("data")
ENV_FILE = Path(".env")
GARO_API_BASE_URL_ENV = "GARO_API_BASE_URL"
PROJECT_TIMEZONE = ZoneInfo("Europe/Stockholm")
CSV_FIELDNAMES = [
    "serial",
    "name",
    "meter_serial",
    "date",
    "energy_kwh",
    "meter_start_wh",
    "meter_stop_wh",
]


def report_config_diagnostics(chargebox_config: dict[str, Any]) -> None:
    """Print non-fatal warnings for suspicious GARO config entries."""
    energy_serials = [str(serial) for serial in chargebox_config.get("energySerials", [])]
    slave_list = chargebox_config.get("slaveList", [])

    serial_to_slave = {
        str(slave["serialNumber"]): slave
        for slave in slave_list
        if "serialNumber" in slave
    }
    orphaned_energy_serials = [
        serial for serial in energy_serials if serial not in serial_to_slave
    ]
    unnamed_energy_serials = [
        serial
        for serial in energy_serials
        if serial in serial_to_slave
        and not str(serial_to_slave[serial].get("reference", "")).strip()
    ]

    if not orphaned_energy_serials and not unnamed_energy_serials:
        return

    print("\nConfig diagnostics:")
    if orphaned_energy_serials:
        print(
            "- Missing from slaveList: "
            f"{', '.join(orphaned_energy_serials)}"
        )
    for serial in unnamed_energy_serials:
        slave = serial_to_slave[serial]
        reference = str(slave.get("reference", "")).strip() or "<missing>"
        origin = str(slave.get("origin", "")).strip() or "<missing>"
        print(
            f"- Found in slaveList: serialNumber={serial}, reference={reference}, origin={origin}"
        )
        print("  This serial is present in energySerials but has no usable garage reference.")
    print("  Fetch will continue, but these chargers are likely misconfigured in GARO.")


def summarize_http_error(error: requests.HTTPError) -> str:
    """Return a short, user-facing summary of an HTTP error response."""
    response = error.response
    if response is None:
        return str(error)

    response_body = response.text.strip().replace("\n", " ")
    if len(response_body) > 200:
        response_body = f"{response_body[:197]}..."

    details = f"HTTP {response.status_code}"
    if response.reason:
        details = f"{details} {response.reason}"
    if response_body:
        details = f"{details} - {response_body}"
    return details


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
    raw_base_url = os.getenv(GARO_API_BASE_URL_ENV)
    base_url = sanitize_garo_api_base_url(raw_base_url)
    if not base_url:
        raise KeyError(
            f"Missing {GARO_API_BASE_URL_ENV} in the environment or {ENV_FILE}."
        )
    return base_url


def sanitize_garo_api_base_url(base_url: str | None) -> str | None:
    """Trim and normalize the configured GARO base URL.

    Some editors or copy/paste operations can introduce hidden combining marks at the
    end of the URL, which requests then percent-encodes into the path and causes 404s.
    """
    if base_url is None:
        return None

    cleaned_url = unicodedata.normalize("NFKC", base_url).strip()
    cleaned_url = "".join(
        character
        for character in cleaned_url
        if unicodedata.category(character)[0] != "C"
        and not unicodedata.combining(character)
    ).rstrip("/")

    parsed_url = urlparse(cleaned_url)
    if not parsed_url.scheme or not parsed_url.netloc:
        raise ValueError(
            f"Invalid {GARO_API_BASE_URL_ENV}: {base_url!r}. Expected a full http(s) URL."
        )

    return cleaned_url


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
    current_date = datetime.now(PROJECT_TIMEZONE)
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
    report_config_diagnostics(chargebox_config)
    chargebox_serials = chargebox_config.get("energySerials", [])
    if not chargebox_serials:
        raise KeyError("No `energySerials` found in the GARO `/config` response.")

    serial_to_reference = {
        str(slave["serialNumber"]): slave.get("reference", str(slave["serialNumber"]))
        for slave in chargebox_config.get("slaveList", [])
        if "serialNumber" in slave
    }

    energy_rows: list[dict[str, Any]] = []
    failed_requests: dict[str, dict[str, Any]] = defaultdict(
        lambda: {"garage_name": "", "months": [], "error": ""}
    )
    last_month_to_fetch = get_last_month_to_fetch(YEAR)
    for chargebox_serial in chargebox_serials:
        garage_name = serial_to_reference.get(
            str(chargebox_serial), str(chargebox_serial)
        )
        print(f"Fetching {garage_name} (serial {chargebox_serial})...")
        for month in range(1, last_month_to_fetch + 1):
            try:
                month_data = fetch_monthly_energy_data(
                    base_url, chargebox_serial, YEAR, month
                )
            except requests.HTTPError as error:
                error_summary = summarize_http_error(error)
                failure = failed_requests[str(chargebox_serial)]
                failure["garage_name"] = garage_name
                failure["months"].append(f"{YEAR}-{month:02d}")
                failure["error"] = error_summary
                print(f"  {YEAR}-{month:02d}: API error, skipping")
                continue
            if month_data is None:
                print(f"  {YEAR}-{month:02d}: no data")
                continue

            timestamps = month_data.get("timestamps", [])
            energy_values = month_data.get("values", [])
            meter_start_values = month_data.get("start", [])
            meter_stop_values = month_data.get("stop", [])
            meter_serial = month_data.get("meterSerial", "")

            n = len(timestamps)
            if len(energy_values) != n:
                raise ValueError(
                    f"Mismatched timestamp/value lengths for serial {chargebox_serial}, "
                    f"{YEAR}-{month:02d}: {n} timestamps, {len(energy_values)} values."
                )
            if meter_start_values and len(meter_start_values) != n:
                raise ValueError(
                    f"Mismatched meter_start length for serial {chargebox_serial}, "
                    f"{YEAR}-{month:02d}: {n} timestamps, {len(meter_start_values)} meter_start."
                )
            if meter_stop_values and len(meter_stop_values) != n:
                raise ValueError(
                    f"Mismatched meter_stop length for serial {chargebox_serial}, "
                    f"{YEAR}-{month:02d}: {n} timestamps, {len(meter_stop_values)} meter_stop."
                )

            for index, timestamp_ms in enumerate(timestamps):
                reading_date = (
                    datetime.fromtimestamp(timestamp_ms / 1000, tz=timezone.utc)
                    .astimezone(PROJECT_TIMEZONE)
                    .strftime("%Y-%m-%d")
                )
                energy_rows.append(
                    {
                        "serial": chargebox_serial,
                        "name": garage_name,
                        "meter_serial": meter_serial,
                        "date": reading_date,
                        "energy_kwh": round(energy_values[index], 4),
                        "meter_start_wh": (
                            meter_start_values[index] if meter_start_values else None
                        ),
                        "meter_stop_wh": (
                            meter_stop_values[index] if meter_stop_values else None
                        ),
                    }
                )

            month_total_kwh = sum(value for value in energy_values if value is not None)
            print(
                f"  {YEAR}-{month:02d}: {month_total_kwh:.2f} kWh ({len(timestamps)} days)"
            )

    if failed_requests:
        print("\nCompleted with API errors:")
        for serial, failure in sorted(failed_requests.items()):
            month_list = ", ".join(failure["months"])
            print(
                f"- {failure['garage_name']} (serial {serial}): "
                f"{month_list} -> {failure['error']}"
            )

    return energy_rows


def write_output_files(energy_rows: list[dict[str, Any]]) -> None:
    """Write the yearly JSON and CSV exports used by the workbook generator."""
    DATA_DIR.mkdir(exist_ok=True)
    json_output_path = DATA_DIR / f"energy_{YEAR}.json"
    csv_output_path = DATA_DIR / f"energy_{YEAR}.csv"
    json_tmp_path = json_output_path.with_suffix(".json.tmp")
    csv_tmp_path = csv_output_path.with_suffix(".csv.tmp")

    with json_tmp_path.open("w", encoding="utf-8") as json_file:
        json.dump(energy_rows, json_file, indent=2)
    json_tmp_path.replace(json_output_path)
    print(f"\nSaved {len(energy_rows)} records to {json_output_path}")

    with csv_tmp_path.open("w", newline="", encoding="utf-8") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=CSV_FIELDNAMES)
        writer.writeheader()
        writer.writerows(energy_rows)
    csv_tmp_path.replace(csv_output_path)
    print(f"Saved {len(energy_rows)} records to {csv_output_path}")


def main() -> None:
    """Run the yearly GARO export workflow from API fetch to file output."""
    energy_rows = build_daily_energy_rows()
    write_output_files(energy_rows)


if __name__ == "__main__":
    main()
