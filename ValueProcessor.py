class ValuesHolder:
    def __init__(self, index, value):
        self.index = index
        self.value = value

class ValuesHelper:
    current_index = -1
    def __init__(self, values):
        self.values = values

    def index(self):
        return self.current_index

    def next(self):
        self.current_index = self.current_index + 1 if self.current_index < len(self.values) - 1 else  0
        return self.values[self.current_index]
