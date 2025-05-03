# main.py
from openpyxl_excel_reader import OpenpyxlExcelReader
from search_excel_usecase import SearchExcelUsecase

def main():
    dir_path = input("検索するフォルダ: ")
    keyword = input("検索するキーワード: ")

    reader = OpenpyxlExcelReader()
    usecase = SearchExcelUsecase(reader)
    results = usecase.search(dir_path, keyword)

    for filepath, sheet, content in results:
        if sheet:
            print(f"{filepath} | {sheet}: {content}")
        else:
            print(f"{filepath}: {content}")

if __name__ == "__main__":
    main()
