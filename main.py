# main.py
from infrastructure.openpyxl_excel_reader import OpenPyxlExcelReader
from usecases.search_excel_usecase import SearchExcelUsecase

def main():
    dir_path = input("検索するフォルダ: ")
    keyword = input("検索するキーワード: ")

    # 依存性注入
    reader = OpenPyxlExcelReader()
    usecase = SearchExcelUsecase(reader)

    # ユースケースの実行
    results = usecase.search(dir_path, keyword)

    # 結果の表示
    for result in results:  # SearchResult オブジェクトを直接受け取る
        if result.sheet_title:
            print(f"{result.filepath} | {result.sheet_title}: {result.content}")
        else:
            print(f"{result.filepath}: {result.content}")

if __name__ == "__main__":
    main()
