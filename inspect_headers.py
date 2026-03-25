
import openpyxl

template_path = "/Users/fujimotogakuto/Library/CloudStorage/GoogleDrive-manamana072554@gmail.com/マイドライブ/営業レポート作成AI/251125エキスパートハブ様_営業レポート.xlsx"
wb = openpyxl.load_workbook(template_path)
sheet = wb.active

for row in range(10, 16):
    cell_a = sheet[f"A{row}"]
    print(f"A{row}: {cell_a.value}")
