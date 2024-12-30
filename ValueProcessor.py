import logging

logger = logging.getLogger(__name__)

class ValuesHolder:
    def __init__(self, fraction_index, index, value):
        self.fraction_index = fraction_index
        self.index = index
        self.value = value

class ValuesHelper:
    def __init__(self, values):
        self.values = values
        self.current_fraction = -1
        self.min_index = 0
        self.max_index = len(self.values) - 1

    def index(self):
        for item in self.values:
            if item.fraction_index == self.current_fraction:
                return item.index

    def get_item_by_fraction_index(self, idx):
        for item in self.values:
            if item.fraction_index == idx:
                self.current_fraction = item.fraction_index
                return item

    def set_value_by_index(self, idx, value):
        for item in self.values:
            if item.index == idx:
                item.value = value

    def next(self):
        if self.current_fraction == len(self.values) - 1:
            self.current_fraction = 0
        else:
            self.current_fraction += 1
        item = self.get_item_by_fraction_index(self.current_fraction)
        return item.value

    def next_min(self):
        item = self.get_item_by_fraction_index(self.min_index).value
        self.min_index += 1
        if self.min_index > len(self.values) - 1:
            logging.error("MIN index out of bound.")
            raise IndexError
        return item

    def next_max(self):
        item = self.get_item_by_fraction_index(self.max_index).value
        self.max_index -= 1
        if self.min_index < 0:
            logging.error("MAX index out of bound.")
            raise IndexError
        return item
