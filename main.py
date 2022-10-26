"""
A module that runs the program
"""

from pathlib import Path
from query import user_query
from workbook import Workbook

if __name__ == "__main__":
    wb = Workbook()
    wb.open_file()
    wb.choose_worklist()
    ws = wb.get_worklist()
    user_query(ws)
    wb.save()
