"""
A module that runs the program
"""

from query import Query
from workbook import Workbook

if __name__ == "__main__":
    workbook = Workbook()
    workbook.open_file()
    user_query = Query(workbook)
    user_query.run()
