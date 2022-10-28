"""
Workbook class implementation
"""

from pathlib import Path
import openpyxl


class Workbook:
    """
    A class capable of opening and saving excel files.
    Also stores info of active files and sheets
    """
    def __init__(self):
        self.work_book = None
        self.wb_name = None
        self.sheets = None
        self.work_sheet = None

    def open_file(self):
        """
        Opens an excel file from a working directory
        """
        self.wb_name = input('Input file name')
        if not self.get_file_name():
            self.wb_name = 'demo'
        if Path(fr'.\work\{self.get_file_name()}.xlsx').exists() is False:
            raise FileNotFoundError("The file you're looking for does not exist")
        self.work_book = openpyxl.load_workbook(Path('.') / 'work' / f'{self.get_file_name()}.xlsx')
        self.sheets = self.work_book.sheetnames

    def choose_worklist(self):
        """
        Allows user to choose a sheet from a current workbook
        """
        while True:
            choose_ws = input(f'The file "{self.get_file_name()}" '
                              f'contains the next spreadsheets: {self.get_sheets_names()}\n'
                              'Please, choose one of the available sheets '
                              'by entering its name or ordinal number: ')
            if choose_ws == 'stop':
                print('Program stopped.')
                return None
            try:
                self.work_sheet = self.get_wb()[self.get_sheets_names()[int(choose_ws) - 1]]
                break
            except IndexError:
                print('List index out of range.\n '
                      'Please, try again or enter "stop" to stop searching.')
            except ValueError:
                if choose_ws not in self.get_sheets_names():
                    print('The sheet you are looking for was not found.\n'
                          'Please, try again or enter "stop" to stop searching.')
                else:
                    self.work_sheet = self.get_wb()[choose_ws]
                    break
        print(f'Current worksheet is {self.get_worklist().title}')
        return 0

    def save(self):
        """
        Saves changes in a new file
        """
        self.get_wb().save(Path('.') / 'work' / f'{self.get_file_name()}_result.xlsx')

    def get_wb(self):
        """
        Returns an active workbook
        """
        return self.work_book

    def get_file_name(self):
        """
        Returns the name of current workbook
        """
        return self.wb_name

    def get_sheets_names(self):
        """
        Returns a list of available sheets in a current workbook
        """
        return self.sheets

    def get_worklist(self):
        """
        Returns an active worksheet
        """
        return self.work_sheet
