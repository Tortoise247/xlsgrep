# main.py
from dependency_injector import containers, providers
from infrastructure.openpyxl_excel_reader import OpenPyxlExcelReader
from usecases.search_excel_usecase import SearchExcelUsecase

class AppContainer(containers.DeclarativeContainer):
    excel_reader = providers.Factory(OpenPyxlExcelReader)
    search_usecase = providers.Factory(SearchExcelUsecase, reader=excel_reader)

def main():
    container = AppContainer()
    usecase = container.search_usecase()

    dir_path = input("検索するフォルダ: ")
    keyword = input("検索するキーワード: ")

    for result in usecase.search(dir_path, keyword):
        if result.sheet_title:
            print(f"{result.filepath} | {result.sheet_title}: {result.content}")
        else:
            print(f"{result.filepath}: {result.content}")

if __name__ == "__main__":
    main()
