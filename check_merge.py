import openpyxl
from openpyxl.cell.cell import MergedCell

TEMPLATE_PATH = "/Users/fujimotogakuto/Library/CloudStorage/GoogleDrive-manamana072554@gmail.com/マイドライブ/営業レポート作成AI/251125エキスパートハブ様_営業レポート.xlsx"

def check_merged_cells():
    print(f"Loading {TEMPLATE_PATH}...")
    try:
        wb = openpyxl.load_workbook(TEMPLATE_PATH)
        sheet = wb.active
        
        targets = [
            'B1', 'B4', 'B5', 'B6', 'B7', 'B8', 'B9', 'B10', 'B11', 
            'B15', 'B49', 
            'J3', 'J5', 'J6', 'J7', 'J8', 'J9', 'J10',
            'A50'
        ]
        
        print(f"Checking {len(targets)} target cells...")
        
        for coord in targets:
            cell = sheet[coord]
            if isinstance(cell, MergedCell):
                print(f"[!] {coord} is a MergedCell. Writing here will fail.")
                # Find which range it belongs to
                for merged_range in sheet.merged_cells.ranges:
                    if coord in merged_range:
                        print(f"    -> Belongs to range: {merged_range}")
                        print(f"    -> Top-left cell to write to: {merged_range.start_cell}") # This doesn't exist as attr, convert
                        break
            else:
                print(f"[OK] {coord} is a normal Cell (or top-left of merge).")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    check_merged_cells()
