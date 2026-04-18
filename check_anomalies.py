"""Scan energy JSON for anomalies and print a structured report."""

import csv
import json
import statistics
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path

YEAR = 2026
CONFIG_FILE = Path("garage_config.json")
DATA_DIR = Path("data")
ENERGY_FILE = DATA_DIR / f"energy_{YEAR}.json"
ENERGY_CSV_FILE = DATA_DIR / f"energy_{YEAR}.csv"
MONTH_NAMES = [
    "Jan",
    "Feb",
    "Mars",
    "April",
    "Maj",
    "Juni",
    "Juli",
    "Aug",
    "Sep",
    "Okt",
    "Nov",
    "Dec",
]

# Garage+months where 0 kWh is expected — suppress false positives.
EXPECTED_ZERO_MONTHS: dict[str, list[int]] = {
    "Garage08": list(range(1, 13)),  # no sessions ever recorded
    "Garage12": [1, 2, 3],  # started recording April 2026
}

# Always-on notes — known structural issues independent of data.
STANDING_ISSUES = [
    "Garage03: only serial 931377 configured. Second charger likely exists — "
    "kWh may be ~50% of actual. Find second serial in GARO admin and add to energySerials.",
]

SPIKE_FACTOR = 3.0  # flag if month > SPIKE_FACTOR × median of other non-zero months


def load_config() -> dict:
    """Load the garage configuration used to classify expected charger data."""
    with CONFIG_FILE.open(encoding="utf-8") as f:
        return json.load(f)


def current_completed_months() -> list[int]:
    """Return month numbers that should already have data for the configured year."""
    now = datetime.now(timezone.utc)
    if now.year > YEAR:
        return list(range(1, 13))
    return list(range(1, now.month + 1))


def load_monthly_totals() -> tuple[dict[str, dict[int, float]], dict[str, set[int]]]:
    """Return {garage_name: {month_num: total_kwh}} from JSON records."""
    rows = load_energy_rows()
    totals: dict[str, dict[int, float]] = defaultdict(lambda: defaultdict(float))
    seen: dict[str, set[int]] = defaultdict(set)
    for row in rows:
        month = int(row["date"][5:7])
        totals[row["name"]][month] += row.get("energy_kwh") or 0.0
        seen[row["name"]].add(month)
    # Convert defaultdict → plain dict so missing keys are detectable.
    return {garage: dict(months) for garage, months in totals.items()}, seen


def load_energy_rows() -> list[dict]:
    """Load energy rows from JSON, or fall back to CSV if JSON is empty/invalid."""
    if ENERGY_FILE.exists():
        try:
            with ENERGY_FILE.open(encoding="utf-8") as f:
                rows = json.load(f)
        except json.JSONDecodeError:
            rows = None
        else:
            if not isinstance(rows, list):
                raise ValueError(
                    f"Expected a JSON list in {ENERGY_FILE}, got {type(rows).__name__}."
                )
            return rows

    if ENERGY_CSV_FILE.exists():
        with ENERGY_CSV_FILE.open(newline="", encoding="utf-8") as csv_file:
            rows = []
            for row in csv.DictReader(csv_file):
                rows.append(
                    {
                        "name": row["name"],
                        "date": row["date"],
                        "energy_kwh": float(row["energy_kwh"] or 0),
                    }
                )
        print(
            f"Using CSV fallback because {ENERGY_FILE} is empty or invalid: {ENERGY_CSV_FILE}"
        )
        return rows

    raise ValueError(
        f"Energy data could not be loaded. {ENERGY_FILE} is empty or invalid and "
        f"no CSV fallback was found at {ENERGY_CSV_FILE}. Run fetch_garo_energy.py again."
    )


def detect_anomalies(
    config: dict,
    monthly_totals: dict[str, dict[int, float]],
    seen_months: dict[str, set[int]],
    completed_months: list[int],
) -> list[dict[str, str | None]]:
    """Return structured anomaly records for missing data, zeros, and spikes."""
    anomalies = []

    ovre_garages = set(config.get("övre", {}).keys())
    nedre_garages = set(config.get("nedre", {}).keys())
    all_json_garages = set(monthly_totals.keys())

    # Nedre garages should have no JSON data.
    surprise_nedre = all_json_garages & nedre_garages
    for garage in sorted(surprise_nedre):
        anomalies.append(
            {
                "severity": "WARN",
                "garage": garage,
                "month": None,
                "kind": "UNEXPECTED_DATA",
                "detail": "Nedre garage appears in JSON — data source should be invoices only.",
            }
        )

    # Garages in JSON but not in config at all.
    unknown_garages = all_json_garages - ovre_garages - nedre_garages
    for garage in sorted(unknown_garages):
        anomalies.append(
            {
                "severity": "WARN",
                "garage": garage,
                "month": None,
                "kind": "UNKNOWN_GARAGE",
                "detail": f"Garage '{garage}' in JSON but not in garage_config.json.",
            }
        )

    for garage in sorted(ovre_garages):
        garage_totals = monthly_totals.get(garage, {})
        garage_seen = seen_months.get(garage, set())
        expected_zeros = EXPECTED_ZERO_MONTHS.get(garage, [])

        non_zero_values = [kwh for m, kwh in garage_totals.items() if kwh > 0]
        median_kwh = (
            statistics.median(non_zero_values) if len(non_zero_values) >= 2 else None
        )

        for month in completed_months:
            month_label = MONTH_NAMES[month - 1]
            kwh = garage_totals.get(month)

            if month not in garage_seen:
                # No records at all — API returned 204/404 for this month.
                if month not in expected_zeros:
                    anomalies.append(
                        {
                            "severity": "WARN",
                            "garage": garage,
                            "month": month_label,
                            "kind": "NO_API_DATA",
                            "detail": "Month absent from JSON (API returned no data). Charger offline or fetch failed?",
                        }
                    )
            elif kwh == 0.0:
                # Records exist but sum to zero.
                if month not in expected_zeros:
                    anomalies.append(
                        {
                            "severity": "WARN",
                            "garage": garage,
                            "month": month_label,
                            "kind": "ZERO_KWH",
                            "detail": "API returned records but total is 0 kWh. Charger fault or no sessions?",
                        }
                    )
            elif median_kwh is not None and kwh > SPIKE_FACTOR * median_kwh:
                anomalies.append(
                    {
                        "severity": "INFO",
                        "garage": garage,
                        "month": month_label,
                        "kind": "SPIKE",
                        "detail": f"{kwh:.1f} kWh is >{SPIKE_FACTOR:.0f}× median ({median_kwh:.1f} kWh). Double-check meter reading.",
                    }
                )

    return anomalies


def print_report(anomalies: list[dict[str, str | None]]) -> None:
    """Print a human-readable anomaly report grouped by severity."""
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    print(f"\n{'='*60}")
    print(f"  ANOMALY REPORT — energy_{YEAR}.json  [{now_str}]")
    print(f"{'='*60}")

    if STANDING_ISSUES:
        print("\n[KNOWN ISSUES — always review]")
        for issue in STANDING_ISSUES:
            print(f"  * {issue}")

    warns = [a for a in anomalies if a["severity"] == "WARN"]
    infos = [a for a in anomalies if a["severity"] == "INFO"]

    def _print_group(items, label):
        """Print one severity group heading followed by its anomaly lines."""
        if not items:
            return
        print(f"\n[{label}]")
        for a in items:
            month_str = f"  {a['month']:>5}:" if a["month"] else "      "
            print(f"  {a['garage']:<12}{month_str}  [{a['kind']}] {a['detail']}")

    _print_group(warns, "WARN")
    _print_group(infos, "INFO")

    total = len(anomalies)
    if total == 0:
        print("\n  No anomalies detected (excluding known issues).")
    else:
        print(
            f"\n  {len(warns)} warning(s), {len(infos)} info(s) — review before finalising workbook."
        )
    print(f"{'='*60}\n")


def run_checks() -> list[dict]:
    """Load data, run all checks, return anomaly list."""
    config = load_config()
    monthly_totals, seen_months = load_monthly_totals()
    completed_months = current_completed_months()
    return detect_anomalies(config, monthly_totals, seen_months, completed_months)


def main() -> None:
    """Run the anomaly checker as a standalone script."""
    if not ENERGY_FILE.exists():
        print(f"ERROR: {ENERGY_FILE} not found. Run fetch_garo_energy.py first.")
        return
    anomalies = run_checks()
    print_report(anomalies)


if __name__ == "__main__":
    main()
