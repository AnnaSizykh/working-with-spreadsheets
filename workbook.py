import openpyxl
from pathlib import Path


class Workbook:
    def __init__(self):
        self.wb = None
        self.wb_name = None
        self.sheets = None
        self.ws = None

    def open_file(self):
        self.wb_name = input()
        if not self.get_file_name():
            self.wb_name = 'demo'
        if Path(f'.\work\{self.get_file_name()}.xlsx').exists() is False:
            raise FileNotFoundError ("The file you're looking for does not exist")
        self.wb = openpyxl.load_workbook(Path('.') / 'work' / f'{self.get_file_name()}.xlsx')
        self.sheets = self.wb.sheetnames

    def get_file_name(self):
        return self.wb_name

    def get_sheets_names(self):
        return self.sheets

    def choose_worklist(self):
        while self.ws is None:
            choose_ws = input(f'The file "{self.get_file_name()}" contains the next spreadsheets: {self.get_sheets_names()}\n'
                              'Please, choose one of the available sheets by entering its name or ordinal number: ')
            if choose_ws == 'stop':
                break
            try:
                self.ws = self.get_sheets_names()[int(choose_ws) - 1]
            except IndexError:
                print('List index out of range. Please, try again or enter "stop" to stop searching.')
            except ValueError:
                if choose_ws not in self.get_sheets_names():
                    print('The sheet you are looking for was not found. Please, try again or enter "stop" to stop searching.')
                else:
                    self.ws = choose_ws
