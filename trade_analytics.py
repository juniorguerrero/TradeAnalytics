"""Summarize sales and gross margin from an Excel workbook or CSV."""
import argparse
import csv
import json
from collections import defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

FIELDS = ("date", "sku", "units", "unit_price", "unit_cost")


def rows(path):
    if path.suffix.lower() == ".csv":
        with path.open(encoding="utf-8-sig", newline="") as file:
            yield from csv.reader(file, strict=True)
    elif path.suffix.lower() == ".xlsx":
        from openpyxl import load_workbook

        book = load_workbook(path, read_only=True, data_only=True)
        try:
            sheet = book["Sales"] if "Sales" in book.sheetnames else book.worksheets[0]
            yield from sheet.iter_rows(values_only=True)
        finally:
            book.close()
    else:
        raise ValueError("Use a .csv or .xlsx file")


def number(value):
    try:
        result = Decimal(str(value))
    except InvalidOperation as error:
        raise ValueError("expected a number") from error
    if not result.is_finite() or result < 0:
        raise ValueError("expected a finite nonnegative number")
    return result


def summarize(path):
    source = rows(path)
    try:
        header = next(source, ())
        columns = [str(value or "").strip().lower() for value in header]
        if any(columns.count(field) != 1 for field in FIELDS):
            raise ValueError("Include exactly one column for each of: " + ", ".join(FIELDS))
        positions = [columns.index(field) for field in FIELDS]
        totals = defaultdict(lambda: [0, Decimal(0), Decimal(0)])
        for line, row in enumerate(source, 2):
            if all(value is None or value == "" for value in row):
                continue
            try:
                day, sku, units, price, cost = (row[index] for index in positions)
                day = day.date() if isinstance(day, datetime) else date.fromisoformat(str(day))
                sku = str(sku or "").strip()
                if not sku:
                    raise ValueError("SKU is required")
                units = number(units)
                if units != units.to_integral_value():
                    raise ValueError("units must be a whole number")
                price, cost = number(price), number(cost)
            except (IndexError, ValueError) as error:
                raise ValueError(f"row {line}: {error}") from error
            record = totals[sku, day.strftime("%Y-%m")]
            record[0] += int(units)
            record[1] += units * price
            record[2] += units * cost
    finally:
        source.close()
    return [
        {"sku": sku, "month": month, "units": units, "revenue": str(revenue),
         "gross_margin_amount": str(revenue - cost),
         "gross_margin_percent": str(round((revenue - cost) * 100 / revenue, 2)) if revenue else None}
        for (sku, month), (units, revenue, cost) in sorted(totals.items())
    ]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("sales_file", type=Path)
    args = parser.parse_args()
    try:
        print(json.dumps(summarize(args.sales_file), indent=2))
    except (ValueError, OSError, csv.Error, ImportError) as error:
        parser.exit(1, f"{error}\n")


if __name__ == "__main__":
    main()
