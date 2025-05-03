# search_excel_usecase.py
import os
from typing import List, Tuple
from excel_reader import IExcelReader

class SearchExcelUsecase:
    def __init__(self, reader: IExcelReader):
        self.reader = reader

    def search(self, directory: str, keyword: str) -> List[Tuple[str, str, str]]:
        results = []
        for root, _, files in os.walk(directory):
            for file in files:
                if file.endswith('.xlsx'):
                    filepath = os.path.join(root, file)
                    try:
                        rows = self.reader.read_rows(filepath)
                        for sheet_name, row in rows:
                            for cell in row:
                                if isinstance(cell, str) and keyword in cell:
                                    results.append((filepath, sheet_name, cell))
                    except Exception as e:
                        results.append((filepath, None, f"Error: {e}"))
        return results
