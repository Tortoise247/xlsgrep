# search_excel_usecase.py
import os

class SearchExcelUsecase:
    def __init__(self, reader):
        self.reader = reader

    def search(self, directory, keyword):
        for root, _, files in os.walk(directory):
            for file in files:
                if file.endswith('.xlsx'):
                    filepath = os.path.join(root, file)
                    try:
                        for sheet_title, row in self.reader.read_rows(filepath):
                            for cell in row:
                                if isinstance(cell, str) and keyword in cell:
                                    yield (filepath, sheet_title, cell)
                    except Exception as e:
                        yield (filepath, None, f"Error: {e}")
