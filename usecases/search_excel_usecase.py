import os
from interfaces.excel_reader import IExcelReader
from domain.search_result import SearchResult

class SearchExcelUsecase:
    def __init__(self, reader: IExcelReader):
        self.reader = reader

    def search(self, directory: str, keyword: str):
        """
        指定されたディレクトリ内のExcelファイルを検索し、キーワードを含むセルを返す。
        """
        for root, _, files in os.walk(directory):
            for file in files:
                if file.endswith('.xlsx'):
                    filepath = os.path.join(root, file)
                    try:
                        # ファイル内のシートと行を処理
                        for sheet_title, row in self.reader.read_rows(filepath):
                            for cell in row:
                                if isinstance(cell, str) and keyword in cell:
                                    yield SearchResult(filepath, sheet_title, cell)
                    except Exception as e:
                        # エラーが発生した場合もジェネレーターを維持
                        yield SearchResult(filepath, None, f"Error: {e}")