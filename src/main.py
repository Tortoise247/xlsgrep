# main.py
import time  # 処理時間を計測するために time モジュールをインポート
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
    pattern = input("検索する正規表現パターン: ")

    start_time = time.time()  # 処理開始時刻を記録

    for result in usecase.search(dir_path, pattern):
        if result.sheet_title:
            print(f"{result.filepath} | {result.sheet_title}: {result.content}")
        else:
            print(f"{result.filepath}: {result.content}")

    end_time = time.time()  # 処理終了時刻を記録
    elapsed_time = end_time - start_time  # 処理時間を計算

    print(f"\n処理時間: {elapsed_time:.2f} 秒")

if __name__ == "__main__":
    main()
