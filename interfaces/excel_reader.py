from abc import ABC, abstractmethod
from typing import Generator, Tuple, List

class IExcelReader(ABC):
    @abstractmethod
    def read_rows(self, filepath: str) -> Generator[Tuple[str, List], None, None]:
        """Excelファイルの行を読み取るジェネレーターを返す"""
        pass