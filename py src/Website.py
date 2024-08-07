import requests
from bs4 import BeautifulSoup

class Website:
    def __init__(self):
        self.COLY = 24
        self.COLR = 17
        self.COLAK = 36
    
    def set_elements(self, doc):
        raise NotImplementedError("Subclasses must implement set_elements method")
    
    def handle_only_year_and_code(self, table, row_index):
        raise NotImplementedError("Subclasses must implement handle_only_year_and_code method")
    
    def set_col_index(self):
        raise NotImplementedError("Subclasses must implement set_col_index method")
    
    def parse_website(self, website, excel_handler):
        try:
            response = requests.get(website)
            if response.status_code == 200:
                soup = BeautifulSoup(response.text, 'html.parser')
                all_tables = self.set_elements(soup)
                
                row_index = 0
                for single_table in all_tables:
                    table_name = single_table.select_one('a').text.strip()
                    row_index = excel_handler.find_index(table_name, 230)
                    
                    print(table_name)
                    if table_name.lower() in ["frost-free", "sans gel", "période de neige", "snow-period"]: 
                        self.handle_only_year_and_code(single_table, row_index)
                    else:
                        self.iterate_table(row_index, single_table, self.set_col_index(), excel_handler)
            else:
                print("Failed to retrieve the webpage:", response.status_code)
        except Exception as e:
            print("Error occurred while parsing website:", e)
    
    def iterate_table(self, row_index, single_table, column, excel_handler):
        if row_index != -1:
            row_index += 2
            rows = single_table.select('tbody tr')
            
            for row in rows:
                row_name = row.select_one('th').text.strip()
                col_index = column
                if row.select_one('td') is not None: #ensures that analytics can fill out ff accurately 
                    excel_val = excel_handler.find_value(row_index)
                    print(excel_val)
                    print(row_name)
                    while (row_name.lower() not in excel_val.lower() and
                           excel_val.lower() not in row_name.lower() and
                           row_index < 230):
                        row_index += 1
                        excel_val = excel_handler.find_value(row_index)
                    if row_index < 230:
                        row_index = self.iterate_row(row, row_index, col_index, excel_handler)
                    else:
                        print(row_name)
                        print("Row not found")
    
    def iterate_row(self, row, row_index, col_index, excel_handler):
        values = row.select('td')
        
        for value in values:
            if value is not None and value.text.strip() != "":
                excel_handler.write_excel(row_index, col_index, value.text.strip())
            if value.select_one('b') is not None:
                excel_handler.write_excel(row_index, self.COLAK, value.text.strip())
            col_index += 1
        
        return row_index + 1
