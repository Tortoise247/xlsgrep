import os
import re  # 正規表現モジュールをインポート
from interfaces.excel_reader import IExcelReader
from domain.search_result import SearchResult

class SearchExcelUsecase:
    def __init__(self, reader: IExcelReader):
        self.reader = reader

    def search(self, directory: str, pattern: str):
        """
        指定されたディレクトリ内のExcelファイルを検索し、正規表現パターンに一致するセルを返す。
        """
        regex = re.compile(pattern)  # 正規表現パターンをコンパイル

        for root, _, files in os.walk(directory):
            for file in files:
                if file.endswith('.xlsx'):
                    filepath = os.path.join(root, file)
                    try:
                        # ファイル内のシートと行を処理
                        for sheet_title, row in self.reader.read_rows(filepath):
                            for cell in row:
                                if isinstance(cell, str) and regex.search(cell):  # 正規表現で検索
                                    yield SearchResult(filepath, sheet_title, cell)
                    except Exception as e:
                        # エラーが発生した場合もジェネレーターを維持
                        yield SearchResult(filepath, None, f"Error: {e}")