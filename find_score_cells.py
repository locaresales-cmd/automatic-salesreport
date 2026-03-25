
import openpyxl

template_path = "/Users/fujimotogakuto/Library/CloudStorage/GoogleDrive-manamana072554@gmail.com/マイドライブ/営業レポート作成AI/251125エキスパートハブ様_営業レポート.xlsx"
wb = openpyxl.load_workbook(template_path)
sheet = wb.active

target_headers = ["商談姿勢", "営業人間力", "商談対応力", "総合数値"]
found_cells = {}

for row in sheet.iter_rows(min_row=1, max_row=60, min_col=1, max_col=10):
    for cell in row:
        if cell.value:
            val_str = str(cell.value)
            for header in target_headers:
                if header in val_str:
                    found_cells[header] = cell.coordinate
                    print(f"Found '{header}' at {cell.coordinate}: {cell.value}")
                    # Also check nearby cells for values
                    # Usually score is next to header (Row+1 or Col+1)

print("Search complete.")
