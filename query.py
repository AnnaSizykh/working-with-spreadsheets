"""
A module for getting user's queries
"""

from functions import Functions


class Query:
    """
    Query class implementation
    """
    def __init__(self, workbook):
        self.query = None
        self.help_message = ''' 
If you want to complete the session type - stop.

        
These are available functions: 
        
To add up all numeric values in the selected area type: add
Tu subtract numeric value type: sub
To multiply all numeric values in the selected type: multi 
To divide numeric value type: div
To calculate the arithmetic mean type: mean
To raise a number in a cell to a degree type: expo 
To round to the number type: rounded
To calculate the logarithm of the numeric value type: log
To move values from one area to another type: move
To copy values from one area to another type: copy
To delete data from cell type: del
To compare of cell values type: compare
To search for a specific element in the table type: find

Type coordinate of a cell as they are, without additional symbols. 
Good example: A21
Bad example: AK-47
'''
        self.empty_query_message = 'Empty query. Type "help" for instructions.'
        self.called_function = None
        self.workbook = workbook
        self.called_function_result = None

    def run(self):
        """
        Executing logic
        """
        if self.workbook.choose_worklist() is None:
            return 0
        while True:
            self.input_query()
            if self.get_query() is None:
                return -1
            if self.get_query() == 'stop':
                self.workbook.save()
                return 0
            if not self.get_query():
                print(self.empty_query())
            if self.get_query() == 'help':
                print(self.get_help())
            if self.get_query() == 'change':
                if self.workbook.choose_worklist() is None:
                    break
            else:
                self.called_function = Functions(self.get_query(), self.workbook.get_worklist())
                self.called_function_result = self.called_function.run()
                if self.get_called_func_result() == 1:
                    print('No such function. Type "help" to get a list of available functions')
                elif self.get_called_func_result() == -1:
                    print('Oops, something went wrong.')
                elif self.get_called_func_result() == 0:
                    print('Action completed')
        return 0

    def input_query(self):
        """
        Gets a function name that user wants to execute
        """
        self.query = input('Please, input function name: ')

    def get_query(self):
        """
        Returns a function's name
        """
        return self.query

    def get_help(self):
        """
        Returns help message
        """
        return self.help_message

    def empty_query(self):
        """
        Returns message for empty query
        """
        return self.empty_query_message

    def get_called_func_result(self):
        """
        Returns result of called function work.
        Used for check of function work correctness
        """
        return self.called_function_result
