# excel_reader.py
from typing import List, Tuple

class IExcelReader:
    def read_rows(self, filepath: str) -> List[Tuple]:
        """指定されたExcelファイルから全ての行データを取得"""
        raise NotImplementedError()
