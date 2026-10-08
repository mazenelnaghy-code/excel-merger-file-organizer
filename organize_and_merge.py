"""Tidy a messy folder and merge its monthly Excel reports into one workbook.

Usage:
    python organize_and_merge.py demo_inbox            # does it
    python organize_and_merge.py demo_inbox --dry-run  # only prints what it would do

What it does:
1. Reads every monthly branch report, whatever its name ("sales_cairo_jan_2024.xlsx", "Cairo Jan 2024 report.xlsx" ...),
   works out the branch and month, and moves it to Reports/<branch>/2024-01 <branch>.xlsx
2. Moves PDFs to Invoices/, images to Images/, everything else to Other/
3. Merges all reports into "All branches 2024.xlsx": every sale in one sheet, plus a Summary sheet
   (revenue per branch per month) and a Totals sheet
4. Writes organize_log.csv: one line per file, from where to where
"""
import argparse
import re
import shutil
from pathlib import Path

import pandas as pd

BRANCHES = {"cairo": "Cairo", "alexandria": "Alexandria", "portsaid": "Port Said", "port said": "Port Said"}
MONTHS = {m: i for i, m in enumerate(["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}
FOLDERS = {".pdf": "Invoices", ".jpg": "Images", ".jpeg": "Images", ".png": "Images"}


def identify(name):
    """Return (branch, month number) from any of the naming styles, or None."""
    low = name.lower()
    branch = next((v for k, v in BRANCHES.items() if k in low), None)
    month = next((n for m, n in MONTHS.items() if re.search(rf"(^|[^a-z]){m}", low)), None)
    return (branch, month) if branch and month else None


def main():
    ap = argparse.ArgumentParser(description="Tidy a folder and merge its monthly Excel reports.")
    ap.add_argument("folder", type=Path)
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()
    root, log, sales = args.folder, [], []

    for f in sorted(p for p in root.iterdir() if p.is_file()):
        if f.suffix.lower() == ".xlsx" and (found := identify(f.stem)):
            branch, month = found
            df = pd.read_excel(f)
            sales.append(df.assign(branch=branch, revenue=df["quantity"] * df["unit_price"]))
            target = root / "Reports" / branch / f"2024-{month:02d} {branch}.xlsx"
        else:
            target = root / FOLDERS.get(f.suffix.lower(), "Other") / f.name
        log.append({"from": f.name, "to": str(target.relative_to(root))})
        if not args.dry_run:
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(f), str(target))

    merged = pd.concat(sales, ignore_index=True)
    merged["month"] = merged["date"].astype(str).str[:7]
    summary = merged.pivot_table(index="month", columns="branch", values="revenue", aggfunc="sum").fillna(0)
    summary["Total"] = summary.sum(axis=1)
    totals = merged.groupby("branch").agg(orders=("product", "count"), units=("quantity", "sum"), revenue=("revenue", "sum"))

    if not args.dry_run:
        with pd.ExcelWriter(root / "All branches 2024.xlsx", engine="openpyxl") as xl:
            merged[["date", "branch", "product", "quantity", "unit_price", "revenue"]].sort_values(["date", "branch"]) \
                .to_excel(xl, sheet_name="All sales", index=False)
            summary.to_excel(xl, sheet_name="Summary")
            totals.to_excel(xl, sheet_name="Totals")
            for ws in xl.sheets.values():
                ws.freeze_panes = "B2"
                for col in ws.columns:
                    ws.column_dimensions[col[0].column_letter].width = 16
        pd.DataFrame(log).to_csv(root / "organize_log.csv", index=False)

    moved = pd.DataFrame(log)["to"].str.split("/|\\\\", regex=True).str[0].value_counts()
    print(f"{len(log)} files {'would be ' if args.dry_run else ''}sorted:")
    for folder, n in moved.items():
        print(f"  {n:>3}  -> {folder}/")
    print(f"{len(sales)} monthly reports merged: {len(merged)} sales rows, revenue {merged['revenue'].sum():,.0f} EGP")
    print("-> All branches 2024.xlsx (All sales, Summary, Totals) + organize_log.csv")


if __name__ == "__main__":
    main()
