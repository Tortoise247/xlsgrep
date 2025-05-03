from search_excel import search_excel

def main():
    dir_path = input("検索するフォルダ: ")
    keyword = input("検索するキーワード: ")
    search_excel(dir_path, keyword)

if __name__ == "__main__":
    main()
