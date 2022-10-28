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
        self.help_message = 'Help message.'
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
                    return -1
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
