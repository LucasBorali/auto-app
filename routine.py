

class Routine:
    def __init__(self, table_sheet=None):
        self.functions = []
        self.table_sheet = table_sheet
       
    
    def add_function(self, function):
        self.functions.append(function)
        
    def exclude_function(self):
        self.functions.pop()

