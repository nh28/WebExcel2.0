from Website import Website  # Assuming Website is defined in a separate module
from ExcelHandler import ExcelHandler  # Assuming ExcelHandler is defined in a separate module

class DataWebsite(Website):
    def __init__(self, excel, sheet):
        self.handler = ExcelHandler(excel, sheet)
        self.table_num = 0
        super().__init__()
        

    def set_elements(self, doc):
        return doc.select("details#normals-data table")

    def handle_only_year_and_code(self, table, row_index):
        if self.table_num == 1:
            values = table.select("td")
            index = 200
            for val in values:
                self.handler.write_excel(index, self.COLAK, val.text.strip())
                index += 1
            self.table_num+=1
        else:
            self.iterate_table(row_index, table, self.COLAK, self.handler)
            self.table_num += 1

    def set_col_index(self):
        return self.COLY
    