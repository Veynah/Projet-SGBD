class FreeLevelHandler:
    def __init__(self, new_row):
        self.new_row = new_row

    def add_additional_fields(self):
        self.new_row['YEAR'] = self.calculate_year()


    def calculate_year(self):
        pass
