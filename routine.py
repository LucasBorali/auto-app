

class Routine:
    def __init__(self, table_sheet=None):
        self.functions = []
        self.table_sheet = table_sheet
       
    
    def add_function(self, function, index=None):
        if index:
            self.functions.insert(index, function)
        elif index == 0:
             self.functions.insert(index, function)
        else:
            self.functions.append(function)
        
    def exclude_function(self, index=-1):
        self.functions.pop(index)
        

