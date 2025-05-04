# xlsgrep

`xlsgrep` は、指定されたディレクトリ内の Excel ファイルを検索し、特定のキーワードを含むセルを抽出するツールです。このプロジェクトは、クリーンアーキテクチャの原則に基づいて設計されています。

---

## プロジェクト構成

```
SRC
│  main.py
│  main.py.bak
│  requirement.txt
│  
├─domain
│  └─  search_result.py
├─infrastructure
│  └─ openpyxl_excel_reader.py
├─interfaces
│  └─ excel_reader.py
└─usecases
   └─ search_excel_usecase.py
```

---

## 各ディレクトリとファイルの説明

### 1. **`main.py`**
- **役割**: アプリケーションのエントリーポイント。
- **責務**:
  - ユーザーから検索ディレクトリとキーワードを入力として受け取る。
  - 依存性注入（DI）コンテナを使用して、ユースケース層（`SearchExcelUsecase`）を初期化。
  - ユースケース層を呼び出して検索結果を取得し、結果を表示する。

---

### 2. **`domain/search_result.py`**
- **役割**: ドメイン層のエンティティを定義。
- **責務**:
  - 検索結果を表現するデータ構造を提供。
  - 検索結果には、ファイルパス、シート名、セルの内容が含まれる。

---

### 3. **`interfaces/excel_reader.py`**
- **役割**: インターフェース層。
- **責務**:
  - Excel ファイルを読み取るためのインターフェースを定義。
  - 具体的な実装（例: `OpenPyxlExcelReader`）に依存しない抽象化を提供。

---

### 4. **`infrastructure/openpyxl_excel_reader.py`**
- **役割**: インフラ層。
- **責務**:
  - `openpyxl` ライブラリを使用して Excel ファイルを読み取る具体的な実装を提供。
  - `IExcelReader` インターフェースを実装。

---

### 5. **`usecases/search_excel_usecase.py`**
- **役割**: ユースケース層。
- **責務**:
  - 指定されたディレクトリ内の Excel ファイルを検索し、キーワードを含むセルを抽出。
  - ファイル検索と行データの処理を行う。

---

### 6. **`requirement.txt`**
- **役割**: プロジェクトで使用する Python ライブラリを定義。
- **内容**:
  - `openpyxl`: Excel ファイルを操作するためのライブラリ。
  - `dependency-injector`: 依存性注入を実現するためのライブラリ。

---

## クリーンアーキテクチャの観点での設計

このプロジェクトは、以下のクリーンアーキテクチャの原則に従っています：

1. **依存性の逆転**:
   - ユースケース層（`SearchExcelUsecase`）は、具体的な実装（`OpenPyxlExcelReader`）ではなく、インターフェース（`IExcelReader`）に依存しています。
   - これにより、Excel リーダーの実装を簡単に差し替えることが可能です。

2. **責務の分離**:
   - ドメイン層（`SearchResult`）、ユースケース層（`SearchExcelUsecase`）、インフラ層（`OpenPyxlExcelReader`）が明確に分離されています。
   - 各層はそれぞれの責務に集中しており、コードの保守性が向上しています。

3. **依存性注入（DI）の活用**:
   - `AppContainer` を使用して依存関係を管理しています。
   - これにより、テスト時にモックを注入するなど、柔軟な設計が可能です。

---

## 実行方法

1. 必要なライブラリをインストールします：
   ```bash
   pip install -r requirement.txt
2. アプリケーションを実行します：
   ```bash
   python main.py
   ```
3. プロンプトに従って、検索ディレクトリとキーワードを入力します。
4. 検索結果が表示されます。