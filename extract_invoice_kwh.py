"""Extract monthly kWh values from invoice PDFs and write them to the summary workbook.

Expected invoice structure:
  2026/Fakturor/Övre/<MM - Mon>/<invoice>.pdf
  2026/Fakturor/Nedre/<MM - Mon>/<invoice>.pdf

Each month typically contains two PDFs: an Elhandel invoice and an Elnät invoice.
The Elhandel invoice contains the kWh usage on the "Spotpris" row, so that is the
document this script looks for.
"""

import re
from pathlib import Path
import pdfplumber
import openpyxl

FAKTUROR_DIR = Path("2026/Fakturor")
SUMMARY_FILE = Path("energy_2026_summary.xlsx")

MONTH_ROW = {
    1: 3, 2: 4, 3: 5, 4: 6, 5: 7, 6: 8,
    7: 9, 8: 10, 9: 11, 10: 12, 11: 13, 12: 14,
}

# Regex: match "Spotpris  2026-01-01 - 2026-01-31  2 772 kWh ..."
SPOTPRIS_RE = re.compile(
    r"Spotpris\s+(\d{4})-(\d{2})-\d{2}\s+-\s+\d{4}-\d{2}-\d{2}\s+([\d\s]+)\s*kWh",
    re.IGNORECASE,
)


def extract_kwh_from_pdf(pdf_path: Path) -> tuple[int, int] | None:
    """Return `(month, kwh)` for an Elhandel invoice, otherwise `None`."""
    with pdfplumber.open(pdf_path) as pdf:
        if len(pdf.pages) < 2:
            print(f"  Warning: {pdf_path.name} has only one page — skipping (expected 2+)")
            return None
        text = pdf.pages[1].extract_text() or ""

    if "ELHANDEL" not in text.upper():
        return None

    m = SPOTPRIS_RE.search(text)
    if not m:
        return None

    month = int(m.group(2))
    kwh = int(m.group(3).replace(" ", ""))
    return month, kwh


def scan_fakturor() -> dict[str, dict[int, int]]:
    """Scan invoice folders and return a nested mapping of group -> month -> kWh."""
    usage_by_group_and_month: dict[str, dict[int, int]] = {"Övre": {}, "Nedre": {}}

    for group in ("Övre", "Nedre"):
        group_dir = FAKTUROR_DIR / group
        if not group_dir.exists():
            print(f"Warning: {group_dir} not found")
            continue

        for pdf_path in sorted(group_dir.rglob("*.pdf")):
            result = extract_kwh_from_pdf(pdf_path)
            if result:
                month, kwh = result
                existing_kwh = usage_by_group_and_month[group].get(month)
                if existing_kwh is not None:
                    if existing_kwh != kwh:
                        raise ValueError(
                            f"Conflicting invoice values for {group} month {month}: "
                            f"{existing_kwh} kWh and {kwh} kWh ({pdf_path})"
                        )
                    print(f"  Warning: duplicate invoice for {group} month={month} ignored ({pdf_path.name})")
                    continue
                usage_by_group_and_month[group][month] = kwh
                print(f"  {group} month={month}: {kwh} kWh  ({pdf_path.name})")

    return usage_by_group_and_month


def has_invoice_data(usage_by_group_and_month: dict[str, dict[int, int]]) -> bool:
    """Return `True` when at least one invoice total was extracted."""
    return any(usage_by_group_and_month[group] for group in usage_by_group_and_month)


def update_summary_workbook(usage_by_group_and_month: dict[str, dict[int, int]]) -> None:
    """Write extracted Nedre kWh totals into the `Grunddata` sheet.

    Övre kWh (col B) is derived from SUM formulas over individual garage sheets
    and must not be overwritten here. Övre invoice values are printed for
    cross-checking only.
    """
    if not SUMMARY_FILE.exists():
        raise FileNotFoundError(f"Summary workbook not found: {SUMMARY_FILE}. Run build_energy_summary_workbook.py first.")

    workbook = openpyxl.load_workbook(SUMMARY_FILE)
    if "Grunddata" not in workbook.sheetnames:
        raise KeyError(f"`Grunddata` sheet not found in summary workbook: {SUMMARY_FILE}")
    worksheet = workbook["Grunddata"]

    if usage_by_group_and_month["Övre"]:
        print("\nÖvre invoice totals (informational — not written; Grunddata col B uses SUM formulas from garage sheets):")
        for month in sorted(usage_by_group_and_month["Övre"]):
            print(f"  Övre month {month}: {usage_by_group_and_month['Övre'][month]} kWh")

    for month, kwh in usage_by_group_and_month["Nedre"].items():
        row = MONTH_ROW[month]
        worksheet.cell(row=row, column=3).value = kwh

    workbook.save(SUMMARY_FILE)
    print(f"\nSaved {SUMMARY_FILE}")


def main() -> None:
    """Run the invoice extraction flow and update the workbook if data is found."""
    print("Scanning invoices...\n")
    usage_by_group_and_month = scan_fakturor()

    if not has_invoice_data(usage_by_group_and_month):
        print("No Elhandel invoices found.")
        return

    update_summary_workbook(usage_by_group_and_month)
    print("\nGrunddata Nedre kWh values written:")
    for month in sorted(usage_by_group_and_month["Nedre"]):
        print(f"  Nedre month {month}: {usage_by_group_and_month['Nedre'][month]} kWh")


if __name__ == "__main__":
    main()
