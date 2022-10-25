from pathlib import Path
from query import query
from workbook import Workbook

if __name__ == "__main__":
    wb = Workbook()
    wb.open_file()
    wb.choose_worklist()
    ws = wb.get_worklist()
    query(ws)
    wb.get_wb().save(Path('.') / 'work' / 'result.xlsx')
