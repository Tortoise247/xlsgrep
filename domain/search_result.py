from dataclasses import dataclass

@dataclass
class SearchResult:
    filepath: str
    sheet_title: str
    content: str