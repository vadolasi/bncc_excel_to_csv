import csv
from pathlib import Path

import xlrd

book = xlrd.open_workbook("./BNCC.xls")

output_dir = Path("./output")
output_dir.mkdir(exist_ok=True)

for sheet_name in book.sheet_names()[1:-1]:
    if " - " not in sheet_name and ". " not in sheet_name:
        sheet = book.sheet_by_name(sheet_name)

        with open(output_dir / f"{sheet_name}.csv", "w", newline="") as output_file:
            for i in range(2, sheet.nrows):
                writer = csv.writer(output_file, delimiter="|")
                writer.writerow([row.value.strip().replace("\n", ";") for row in sheet.row(i)])
