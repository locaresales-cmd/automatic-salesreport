import openpyxl
import pandas as pd

TEMPLATE_PATH = "/Users/fujimotogakuto/Library/CloudStorage/GoogleDrive-manamana072554@gmail.com/マイドライブ/営業レポート作成AI/251125エキスパートハブ様_営業レポート.xlsx"

def analyze_excel(path):
    print(f"Loading {path}...")
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        sheet = wb.active
        print(f"Active sheet: {sheet.title}")
        
        print("\n--- Non-empty cells (First 50 rows) ---")
        for row in range(1, 51):
            row_data = []
            for col in range(1, 15): # A to N
                cell = sheet.cell(row=row, column=col)
                val = cell.value
                if val:
                    row_data.append(f"{cell.coordinate}: {str(val)[:20]}")
            if row_data:
                print(f"Row {row}: " + ", ".join(row_data))

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    analyze_excel(TEMPLATE_PATH)
