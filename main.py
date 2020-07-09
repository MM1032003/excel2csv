import csv, os, openpyxl

os.makedirs("csv_files", exist_ok=True)


def get_value(cell):
    return cell.value
def is_excel(filename):
    if filename.endswith(".xlsx"):
        return True

    return False

print("Gathering Files...")
excel_files = list(filter(is_excel, os.listdir(".")))

for excel_file in excel_files:
    print(f"Working on {excel_file}...")
    wb              = openpyxl.load_workbook(excel_file)
    excel_file_name =  os.path.splitext(excel_file)[0]


    for sheet_name in wb.sheetnames:
        outputfile      = os.path.join("csv_files", excel_file_name+"_"+ sheet_name +".csv")
        print(f"Creating {outputfile}...")
        csv_file        = open(outputfile, "w")
        csv_writer      = csv.writer(csv_file)
        ws      = wb[sheet_name]
        rows    = ws.rows
        for row in rows:
            row_list    = map(lambda cell:cell.value, row)
            csv_writer.writerow(row_list)

csv_file    = None
print("Done!")