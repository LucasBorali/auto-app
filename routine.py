

class Routine:
    def __init__(self):
        self.functions = []
       
    
    def add_function(self, function):
        self.functions.append(function)
        
    def exclude_function(self):
        self.functions.pop()

