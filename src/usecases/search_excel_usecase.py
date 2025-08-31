
import os
import re  # 正規表現モジュールをインポート
from concurrent.futures import ThreadPoolExecutor, as_completed
from interfaces.excel_reader import IExcelReader
from domain.search_result import SearchResult

class SearchExcelUsecase:
    def __init__(self, reader: IExcelReader):
        self.reader = reader

    def search(self, directory: str, pattern: str, max_workers: int = 8):
        """
        指定されたディレクトリ内のExcelファイルを検索し、正規表現パターンに一致するセルを返す。
        マルチスレッドでファイルごとに並列検索する。
        """
        regex = re.compile(pattern)

        # 検索対象ファイル一覧を作成
        filepaths = []
        for root, _, files in os.walk(directory):
            for file in files:
                if file.endswith('.xlsx'):
                    filepaths.append(os.path.join(root, file))

        def search_file(filepath):
            results = []
            try:
                for sheet_title, row in self.reader.read_rows(filepath):
                    for cell in row:
                        if isinstance(cell, str) and regex.search(cell):
                            results.append(SearchResult(filepath, sheet_title, cell))
            except Exception as e:
                results.append(SearchResult(filepath, None, f"Error: {e}"))
            return results

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_file = {executor.submit(search_file, fp): fp for fp in filepaths}
            for future in as_completed(future_to_file):
                for result in future.result():
                    yield result