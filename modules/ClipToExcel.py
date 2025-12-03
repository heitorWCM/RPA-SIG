import pyperclip
import xlsxwriter
import os

def parse_clipboard_table(text):
    rows = text.strip().split("\n")
    return [row.split("\t") for row in rows]

def write_to_excel(data, output_folder, filename="copied_table.xlsx"):
    os.makedirs(output_folder, exist_ok=True)
    file_path = os.path.join(output_folder, filename)

    workbook = xlsxwriter.Workbook(file_path)
    worksheet = workbook.add_worksheet()

    for r, row in enumerate(data):
        for c, value in enumerate(row):
            worksheet.write(r, c, value)

    workbook.close()
    return file_path

# STEP 1: Get copied table
clipboard_text = pyperclip.paste()

# STEP 2: Parse into rows/columns
parsed = parse_clipboard_table(clipboard_text)

# STEP 3: Save to Excel
target_folder = r"C:\temp\excel_exports"
file_saved = write_to_excel(parsed, target_folder)

print("Saved Excel file:", file_saved)
