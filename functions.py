"""
Functions implementation
"""

import math
import re


class Functions:
    """
    Abstract functions class implementation.
    Capable of searching and storing info about functions.
    Runs functions logic
    """
    def __init__(self, function_name, sheet):
        self.called_type = None
        self.func_name = function_name
        self.function = None
        self.function_types = {'counting functions':
                               ['add', 'sub', 'multi', 'div', 'mean'],
                               'difficult functions': ['rounded', 'expo', 'log'],
                               'moving functions': ['move', 'copy'],
                               'delete function': ['delete'],
                               'compare function': ['compare'],
                               'searching function': ['find']}
        self.variables = None
        self.sheet = sheet
        self.result_cell = None

    def run(self):
        """
        Executes function
        """
        if self.define_function_type() == -1:
            return 1
        if self.function_call() == -1:
            return 1
        if self.get_function().run() == -1:
            return -1
        return 0

    def define_function_type(self):
        """
        Defines a type of function relying on user's input
        """
        for function_type in self.function_types:
            if self.get_function_name() in self.function_types.get(function_type):
                self.called_type = function_type
                return 0
        return -1

    def function_call(self):
        """
        Defines a functions' class that might be executed
        """
        if self.get_called_type() is None:
            return -1
        if self.get_called_type() == 'counting functions':
            self.function = CountingFunctions(self.get_function_name(), self.get_worksheet())
        elif self.get_called_type() == 'difficult functions':
            self.function = DifficultFunctions(self.get_function_name(), self.get_worksheet())
        elif self.get_called_type() == 'moving functions':
            self.function = MovingFunctions(self.get_function_name(), self.get_worksheet())
        elif self.get_called_type() == 'delete function':
            self.function = DeleteFunction(self.get_function_name(), self.get_worksheet())
        elif self.get_called_type() == 'compare function':
            self.function = CompareFunction(self.get_function_name(), self.get_worksheet())
        elif self.get_called_type() == 'searching function':
            self.function = SearchingFunction(self.get_function_name(), self.get_worksheet())
        else:
            return -1
        return 0

    def set_variables(self):
        """
        Used for getting data from user
        """
        raise NotImplementedError

    def get_attributes(self):
        """
        Converts user's input to attributes
        """
        raise NotImplementedError

    @staticmethod
    def cell_pattern_check(*args):
        """
        Checks the correctness of user's cells coordinates input
        """
        cell_pattern = re.compile(r'^[A-Z]+\d+$')
        for cell in args:
            if not isinstance(cell, str):
                return -1
            if re.match(cell_pattern, cell) is None:
                return -1
        return 0

    def get_function_name(self):
        """
        Returns user's input
        """
        return self.func_name

    def get_called_type(self):
        """
        Returns a chosen functions' class name
        """
        return self.called_type

    def get_function(self):
        """
        Returns a chosen class
        """
        return self.function

    def get_variables(self):
        """
        Returns user's variables
        """
        return self.variables

    def get_worksheet(self):
        """
        Returns current worksheet
        """
        return self.sheet


class CountingFunctions(Functions):
    """
    Counting functions class implementation
    """
    def __init__(self, function_name, sheet):
        super().__init__(function_name, sheet)
        self.left_cell = None
        self.right_cell = None

    def run(self):
        """
        Executes checks and function
        """
        self.set_variables()
        if self.get_attributes() == -1:
            return -1
        if self.cell_pattern_check(self.left_cell, self.right_cell, self.result_cell) == -1:
            return -1
        if self.function_call() == -1:
            return -1
        return 0

    def function_call(self):
        """
        Defines a function that might be executed and executes it
        """
        if self.get_function_name() == 'add':
            self.addition(self.sheet, self.left_cell, self.right_cell, self.result_cell)
            return 0
        if self.get_function_name() == 'sub':
            self.subtraction(self.sheet, self.left_cell, self.right_cell, self.result_cell)
            return 0
        if self.get_function_name() == 'multi':
            self.multiplication(self.sheet, self.left_cell, self.right_cell, self.result_cell)
            return 0
        if self.get_function_name() == 'div':
            self.division(self.sheet, self.left_cell, self.right_cell, self.result_cell)
            return 0
        if self.get_function_name() == 'mean':
            self.mean(self.sheet, self.left_cell, self.right_cell, self.result_cell)
            return 0
        return -1

    def set_variables(self):
        """
        Gets variables from user
        """
        self.variables = input('Counting functions variables description').split()

    def get_attributes(self):
        """
        Converts variables to attributes
        """
        if self.variables is None or len(self.variables) != 3:
            return -1
        self.left_cell = self.get_variables()[0]
        self.right_cell = self.get_variables()[1]
        self.result_cell = self.get_variables()[2]
        return 0

    @staticmethod
    def addition(sheet, left_corner, right_corner, result_cell):
        """
        Sums every numeral in selected area.
        """
        summa = 0
        for cell_column in sheet[left_corner: right_corner]:
            for cell in cell_column:
                addend = cell.value
                if isinstance(addend, (float, int)):
                    summa += addend
        sheet[result_cell] = summa
        return summa

    @staticmethod
    def subtraction(sheet, reduced_cell, subtrahend_cell, result_cell):
        """
        Subtracts value of a second given cell from a first one.
        """
        reduced = sheet[reduced_cell].value
        subtrahend = sheet[subtrahend_cell].value
        if not (isinstance(reduced, (float, int)) and isinstance(subtrahend, (float, int))):
            sheet[result_cell] = 'N/A'
            return 'N/A'
        difference = reduced - subtrahend
        sheet[result_cell] = difference
        return difference

    @staticmethod
    def multiplication(sheet, left_corner, right_corner, result_cell):
        """
        Multiplies every numeral in selected area.
        """
        composition = 1
        for cell_column in sheet[left_corner: right_corner]:
            for cell in cell_column:
                factor = cell.value
                if isinstance(factor, (float, int)):
                    composition *= factor
        sheet[result_cell] = composition
        return composition

    @staticmethod
    def division(sheet, divisible_cell, divisor_cell, result_cell):
        """
        Divides value of a first given cell on a first one.
        """
        divisible = sheet[divisible_cell].value
        divisor = sheet[divisor_cell].value
        if not (isinstance(divisible, (float, int))
                and isinstance(divisor, (float, int))) or divisor == 0:
            sheet[result_cell] = 'N/A'
            return 'N/A'
        quotient = divisible / divisor
        sheet[result_cell] = quotient
        return quotient

    @staticmethod
    def mean(sheet, left_corner, right_corner, result_cell):
        """
        Finds the mean between all numerals in a given area.
        """
        counter = 0
        summa = 0
        for cell_column in sheet[left_corner: right_corner]:
            for cell in cell_column:
                addend = cell.value
                if isinstance(addend, (float, int)):
                    summa += addend
                    counter += 1
        if counter == 0:
            sheet[result_cell] = 'N/A'
            return 'N/A'
        result = summa / counter
        sheet[result_cell] = result
        return result


class DifficultFunctions(Functions):
    """
    Difficult functions class implementation
    """
    def __init__(self, function_name, sheet):
        super().__init__(function_name, sheet)
        self.cell = None
        self.number = None

    def run(self):
        """
        Executes checks and function
        """
        self.set_variables()
        if self.get_attributes() == -1:
            return -1
        if self.cell_pattern_check(self.cell, self.result_cell) == -1:
            return -1
        if self.function_call() == -1:
            return -1
        return 0

    def function_call(self):
        """
        Defines a function that might be executed and executes it
        """
        if self.get_function_name() == 'rounded':
            self.round_number(self.sheet, self.cell, self.number, self.result_cell)
            return 0
        if self.get_function_name() == 'expo':
            self.exponentiation(self.sheet, self.cell, self.number, self.result_cell)
            return 0
        if self.get_function_name() == 'log':
            self.logarithm(self.sheet, self.cell, self.number, self.result_cell)
            return 0
        return -1

    def set_variables(self):
        """
        Gets variables from user
        """
        self.variables = input('Difficult functions variables description').split()

    def get_attributes(self):
        """
        Converts variables to attributes
        """
        if self.variables is None or len(self.variables) != 3:
            return -1
        self.cell = self.get_variables()[0]
        self.number = int(self.get_variables()[1])
        self.result_cell = self.get_variables()[2]
        return 0

    @staticmethod
    def round_number(sheet, number_cell, decimal, result_cell):
        """
        Rounds a numeral in given cell.
        """
        number = sheet[number_cell].value
        if isinstance(number, (float, int)) and isinstance(decimal, int):
            rounded = round(number, decimal)
            sheet[result_cell] = rounded
            return rounded
        return number

    @staticmethod
    def exponentiation(sheet, number_cell, extent, result_cell):
        """
        Raises a numeral in cell to a given pover.
        """
        number = sheet[number_cell].value
        if not (isinstance(number, (float, int)) and isinstance(extent, (float, int))):
            sheet[result_cell] = 'N/A'
            return 'N/A'
        result = number ** extent
        sheet[result_cell] = result
        return result

    @staticmethod
    def logarithm(sheet, number_cell, base, result_cell):
        """
        Calculates a logarithm using a number in the cell and a given base.
        """
        number = sheet[number_cell].value
        if (not (isinstance(number, (float, int)) and isinstance(result_cell, (float, int)))) \
                or number < 0 or base < 0 or base == 1:
            sheet[result_cell] = 'N/A'
            return 'N/A'
        result = math.log(number, base)
        sheet[result_cell] = result
        return result


class MovingFunctions(Functions):
    """
    Moving functions class implementation
    """
    def __init__(self, function_name, sheet):
        super().__init__(function_name, sheet)
        self.left_corner = None
        self.right_corner = None

    def run(self):
        """
        Executes checks and function
        """
        self.set_variables()
        if self.get_attributes() == -1:
            return -1
        if self.cell_pattern_check(self.left_corner, self.right_corner, self.result_cell) == -1:
            return -1
        if self.function_call() == -1:
            return -1
        return 0

    def function_call(self):
        """
        Defines a function that might be executed and executes it
        """
        if self.get_function_name() == 'move':
            self.move(self.sheet, self.left_corner, self.right_corner, self.result_cell)
            return 0
        if self.get_function_name() == 'copy':
            self.copy(self.sheet, self.left_corner, self.right_corner, self.result_cell)
            return 0
        return -1

    def set_variables(self):
        """
        Gets variables from user
        """
        self.variables = input('Moving functions variables description').split()

    def get_attributes(self):
        """
        Converts variables to attributes
        """
        if self.variables is None or len(self.variables) != 3:
            return -1
        self.left_corner = self.get_variables()[0]
        self.right_corner = self.get_variables()[1]
        self.result_cell = self.get_variables()[2]
        return 0

    @staticmethod
    def move(sheet, moved_left_corner, moved_right_corner, added_left_corner):
        """
        Moves values from one area to another.
        """
        moving_column = sheet[added_left_corner].column
        moving_row = sheet[added_left_corner].row
        for cell_column in sheet[moved_left_corner: moved_right_corner]:
            for cell in cell_column:
                moved_value = cell.value
                sheet.cell(row=moving_row, column=moving_column, value=moved_value)
                cell.value = ''
                moving_row += 1
            moving_column += 1
        return 'successful'

    @staticmethod
    def copy(sheet, copied_left_corner, copied_right_corner, added_left_corner):
        """
        Copies values from one area to another.
        """
        copy_column = sheet[added_left_corner].column
        copy_row = sheet[added_left_corner].row
        for cell_column in sheet[copied_left_corner: copied_right_corner]:
            for cell in cell_column:
                copied_value = cell.value
                sheet.cell(row=copy_row, column=copy_column, value=copied_value)
                copy_row += 1
            copy_column += 1
        return 'successful'


class DeleteFunction(Functions):
    """
    Deletion function class implementation
    """
    def __init__(self, function_name, sheet):
        super().__init__(function_name, sheet)
        self.left_corner = None
        self.right_corner = None

    def run(self):
        """
        Executes checks and function
        """
        self.set_variables()
        if self.get_attributes() == -1:
            return -1
        if self.cell_pattern_check(self.left_corner, self.right_corner) == -1:
            return -1
        if self.function_call() == -1:
            return -1
        return 0

    def function_call(self):
        """
        Executes a function
        """
        if self.get_function_name() == 'delete':
            self.delete(self.sheet, self.left_corner, self.right_corner)
            return 0
        return -1

    def set_variables(self):
        """
        Gets variables from user
        """
        self.variables = input('Delete function variables description').split()

    def get_attributes(self):
        """
        Converts variables to attributes
        """
        if self.variables is None or len(self.variables) != 2:
            return -1
        self.left_corner = self.get_variables()[0]
        self.right_corner = self.get_variables()[1]
        return 0

    @staticmethod
    def delete(sheet, left_corner: str, right_corner: str):
        """
        Clears all cells in a given area.
        """
        for cell_column in sheet[left_corner: right_corner]:
            for cell in cell_column:
                cell.value = ''
        return 'successful'


class CompareFunction(Functions):
    """
    Comparison function class implementation
    """
    def __init__(self, function_name, sheet):
        super().__init__(function_name, sheet)
        self.cell_1 = None
        self.cell_2 = None

    def run(self):
        """
        Executes checks and function
        """
        self.set_variables()
        if self.get_attributes() == -1:
            return -1
        if self.cell_pattern_check(self.cell_1, self.cell_2, self.result_cell) == -1:
            return -1
        if self.function_call() == -1:
            return -1
        return 0

    def function_call(self):
        """
        Executes a function
        """
        if self.get_function_name() == 'compare':
            self.compare(self.sheet, self.cell_1, self.cell_2, self.result_cell)
            return 0
        return -1

    def set_variables(self):
        """
        Gets variables from user
        """
        self.variables = input('Compare function variables description').split()

    def get_attributes(self):
        """
        Converts variables to attributes
        """
        if self.variables is None or len(self.variables) != 3:
            return -1
        self.cell_1 = self.get_variables()[0]
        self.cell_2 = self.get_variables()[1]
        self.result_cell = self.get_variables()[2]
        return 0

    @staticmethod
    def compare(sheet, cell_1: str, cell_2: str, result_cell: str):
        """
        Compares values from two cells
        """
        value_1 = sheet[cell_1].value
        value_2 = sheet[cell_2].value
        if value_1 == value_2:
            sheet[result_cell] = True
            return True
        sheet[result_cell] = False
        return False


class SearchingFunction(Functions):
    """
    Searching function class implementation
    """
    def __init__(self, function_name, sheet):
        super().__init__(function_name, sheet)
        self.element = None

    def run(self):
        """
        Executes checks and function
        """
        self.set_variables()
        if self.get_attributes() == -1:
            return -1
        if self.function_call() == -1:
            return -1
        return 0

    def function_call(self):
        """
        Executes a function
        """
        if self.get_function_name() == 'find':
            self.find(self.sheet, self.element)
            return 0
        return -1

    def set_variables(self):
        """
        Gets variables from user
        """
        self.variables = input('Find function variables description').split()

    def get_attributes(self):
        """
        Converts variables to attributes
        """
        if self.variables is None or len(self.variables) != 1:
            return -1
        self.element = self.get_variables()[0]
        return 0

    @staticmethod
    def find(sheet, element: str):
        """
        Finds an element in cells.
        """
        similar_results = []
        for cell_object in sheet:
            for cell in cell_object:
                if element == str(cell.value) or element in str(cell.value):
                    similar_results.append(cell.coordinate)
        if len(similar_results) != 0:
            return similar_results
        return 'no matches'
