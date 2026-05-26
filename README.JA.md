# xelixir: Excel 操作ツール

このリポジトリは、Java で実装した Excel 操作ツール群を Python からラップし、
Microsoft Copilot / Claude / LibreChat などの MCP クライアントから利用できる MCP サーバーを提供します。

Excel 操作のコア部分は Java + Apache POI で実装しており、
その上に Python で MCP サーバーを構築することで、次のような利点があります。

- Excel 操作は Apache POI が非常に成熟しており、細かい機能や互換性の面で有利
- Python 製の MCP サーバーは実装が容易で、`fastmcp` などのライブラリを活用しやすい

Python から Java を呼び出すラッパーを用意することで、

- 「Excel 操作は Java/Apache POI に任せつつ、MCP サーバー自体は Python で軽量に構築する」

という構成を実現しています。

## プロジェクト概要とアーキテクチャ

- **目的**
  - Excel ブック/シートの読み書き、書式設定、グラフ作成、ピボットテーブル用メタデータ作成、
    シート管理などの操作を MCP ツールとして提供すること
  - これらの操作を Microsoft Copilot / Claude / LibreChat などの MCP クライアントから安全かつ再利用可能な形で呼び出せるようにすること
  - Excel 操作の実装を Java + Apache POI に集約し、MCP レイヤーを Python に分離することで、保守性と拡張性を高めること

- **Java 側 (Excel コア処理)**
  - Java 11 以上 + Apache POI を利用
  - 実装は [`java/src/xelixir`](java/src/xelixir/ExcelUtils.java:1) 以下に配置
  - [`compile.sh`](java/tools/compile.sh:1) で `src/xelixir/*.java` をコンパイルし、`dist` ディレクトリにクラスファイルを出力
  - 利用する各種ライブラリ JAR は [`java/jars`](java/jars/poi-5.4.1.jar:1) に配置

- **Python 側 (MCP サーバー)**
  - Python 3.11 以上
  - メインパッケージは [`pyproject.toml`](mcp/pyproject.toml:1) で定義された `xelixir`
  - MCP エントリーポイントは [`xelixir.py`](mcp/xelixir.py:1)
    - `FastMCP` を利用して MCP ツールを登録
    - `stdio` / `http` / `sse` での公開に対応
    - SSE モード時には `/files/...` で静的ファイルを配信
  - Java ツールを呼び出す薄いラッパーは [`mcp/src/excel/wrapper.py`](mcp/src/excel/wrapper.py:1) にあり、
    [`mcp/src/excel/__init__.py`](mcp/src/excel/__init__.py:1) から再エクスポートされています。

- **サポートするトランスポート**
  - `stdio`: Claude Desktop などのデスクトップ MCP ランタイム向け
  - `http`: streamable HTTP
  - `sse`: Server-Sent Events。LibreChat などの Web UI からの利用を想定

- **ファイル共有のパターン**
  - コンテナ内の共有ディレクトリ: デフォルトは `/mnt/data`。`WORKSPACE_DIR` 環境変数で変更可能
  - 共有ディレクトリ配下のファイルに対して、`<EXCEL_PUBLIC_BASE_URL>/files/<相対パス>` 形式の `download_url` を生成

## 目次

1. [システム要件](#システム要件)
2. [インストール](#インストール)
3. [スタンドアロン CLI の使い方](#スタンドアロン-cli-の使い方)
4. [MCP クライアントとの連携](#mcp-クライアントとの連携)
5. [Docker を使う場合](#docker-を使う場合)
6. [ファイル共有と download_url](#ファイル共有と-download_url)
7. [主なツール一覧](#主なツール一覧)
8. [開発・貢献方法](#開発・貢献方法)
9. [ライセンス](#ライセンス)
10. [注意事項](#注意事項)

## システム要件

- Java 11 以上
- Python 3.11 以上
- Excel ファイル形式: .xlsx
- 対応 OS: Windows 10/11, macOS 12+, Linux (Ubuntu 20.04+)

## インストール

### 1. リポジトリのクローン

```bash
git clone git@github.com:isoittech/xelixir.git
cd xelixir
```

### 2. Java ツールのビルド

Java 側の Excel ツールは、[`java/tools/compile.sh`](java/tools/compile.sh:1) でコンパイルできます:

```bash
cd java
./tools/compile.sh
cd ..
```

`dist` ディレクトリにコンパイル結果が出力され、Python 側のラッパーから利用されます。

### 3. Python MCP サーバーのセットアップ

Python 版 MCP サーバーは [`mcp/`](mcp) ディレクトリにあります。
`uv` を使ってローカル環境から起動することを想定しています。

```bash
cd mcp
uv run xelixir --help
```

初回実行時に、[`pyproject.toml`](mcp/pyproject.toml:1) に基づいて必要な依存関係 (`fastmcp>=0.3.0` など) がインストールされます。
MCP 側のパッケージ構成やローカル開発に特化した情報は、[`mcp/README.md`](mcp/README.md:1) も参照してください。

## スタンドアロン CLI の使い方

このリポジトリには MCP サーバーに加えて、[`cli/`](cli) 配下に独立した CLI パッケージも含まれています。
CLI は MCP サーバーと同じ Excel ラッパー層を再利用しているため、MCP クライアントを使わずにターミナルから直接 Excel 操作を実行できます。

CLI パッケージ定義は [`cli/pyproject.toml`](cli/pyproject.toml:1)、コマンド実装本体は [`cli/src/xelixir_cli/tools.py`](cli/src/xelixir_cli/tools.py:1) にあります。

リポジトリのルートから次のように実行できます。

```bash
uv run --project cli xelixir-tool --help
uv run --project cli xelixir-tool tool --help
```

基本的な使用例:

```bash
# ワークブックを新規作成
uv run --project cli xelixir-tool tool create-excel \
  --path ./example.xlsx \
  --sheet-name Data

# シート一覧を取得
uv run --project cli xelixir-tool tool list-sheets \
  --path ./example.xlsx

# インライン JSON で 2 次元配列を書き込み
uv run --project cli xelixir-tool tool write-excel \
  --path ./example.xlsx \
  --sheet-name Data \
  --data-json '[["A1","B1"],["A2","B2"]]'
```

CLI の出力は MCP ツールの戻り値にできるだけ合わせた JSON 形式です。
また、配列系の入力は `--data-json` や `--rows-json` などに対して、インライン JSON 文字列または JSON ファイルパスを渡せます。

## MCP クライアントとの連携

### Claude Desktop などから利用する場合

1. `xelixir/mcp` を MCP 設定ディレクトリとして指定します。
2. [`mcp/mcp-config.json`](mcp/mcp-config.json:1) に、MCP サーバーの定義が含まれています。

`mcp-config.json` の例:

```json
{
  "mcpServers": {
    "xelixir": {
      "command": "uv",
      "args": [
        "run",
        "xelixir"
      ],
      "env": {
        "PYTHONPATH": "./src"
      }
    }
  }
}
```

MCP クライアントを再起動すると、`xelixir` というサーバー名で
Excel 操作用の MCP ツール群が利用できるようになります。

## Docker を使う場合

[`mcp/`](mcp) ディレクトリには [`Dockerfile`](mcp/Dockerfile:1) / [`docker-compose.yml`](mcp/docker-compose.yml:1) も用意しています。
SSE モードで HTTP 経由の接続を行いたい場合に利用できます。

```bash
cd mcp
docker build -t xelixir .
docker run --rm -p 8585:8585 xelixir xelixir -t sse -p 8585
```

`docker-compose.yml` を使う場合:

```bash
cd mcp
docker compose up -d
```

MCP クライアント側からは、`http://host.docker.internal:8585/sse` などの SSE エンドポイントを指定します
（LibreChat など、クライアントから見た到達先を指定）。

## ファイル共有と download_url

LibreChat などの MCP クライアントは、ツール実行環境のファイルパスとして `/mnt/data/...` を使うことがあります
（例: 添付ファイルやコード実行の成果物）。
xelixir からも同じパスで読み書きできるように、ホスト側の「共有ディレクトリ」を xelixir コンテナの `/mnt/data` にマウントしてください。

- コンテナ内の共有ディレクトリ: デフォルト `/mnt/data`
- サーバー側の環境変数 `WORKSPACE_DIR` で上書き可能

`docker-compose.yml` では、次のようにホスト -> コンテナのマウントを行います:

```yaml
services:
  xelixir:
    volumes:
      - ${MCP_SHARED_DIR:-./workspace}:/mnt/data
```

- ホスト側で `MCP_SHARED_DIR` を設定すると、実際の共有ディレクトリの場所を制御できます。
- コンテナ内では、[`xelixir.py`](mcp/xelixir.py:76) 内の `get_shared_directory()` が
  `WORKSPACE_DIR`（未設定なら `/mnt/data`）を参照して、download 対象となるパスを検証します。

### download_url と EXCEL_PUBLIC_BASE_URL

エンドユーザーはコンテナ内パス（例: `/mnt/data/file.xlsx`）へ直接アクセスできないため、
xelixir は SSE モード時に `/files/...` を公開し、ツール結果に `download_url`（クリック可能な URL）を含めます。

`download_url` を生成するために、デプロイ環境で `EXCEL_PUBLIC_BASE_URL` の設定が必須です。

- 例: `https://your-domain.example` や `http://your-host:8585`
- ツールが共有ディレクトリ配下のパス `/mnt/data/xxx.xlsx` に書き込んだ場合、
  `build_download_url_for_path()` は `<EXCEL_PUBLIC_BASE_URL>/files/xxx.xlsx` のような URL を生成します。
- `EXCEL_PUBLIC_BASE_URL` が未設定の場合、`tool_create_excel` や `tool_write_excel` など、
  download URL を必要とするツールはエラーを返します（コンテナ内パスのみを返さないようにするため）。

### 同一ファイルの継続編集（上書き保存）

LibreChat にアップロードされた Excel が `/mnt/data/xxx.xlsx` のようなパスで参照できる場合、
xelixir の各ツールは **そのパスを直接読み書き** できます。

- 既存ファイルを編集したい場合は、以降のツール呼び出しで **同じ `path`（例: `/mnt/data/xxx.xlsx`）を使い続けてください**。
- [`write_excel()`](mcp/xelixir.py:242) やシート操作系ツールは、基本的に **同一パスへ更新して保存（上書き）** します。
- 新規作成の [`create_excel()`](mcp/xelixir.py:170) は「既存ファイルを上書きしない」仕様のため、
  アップロード済みファイルの編集用途には使いません。

この方式により「保存のたびに別ファイルを生成する」必要がなく、
スレッド/セッション中は同一ファイルを継続編集できます（ディスク消費を抑えられます）。

## 主なツール一覧

以下は、MCP クライアントから利用できる主なツールの例です。

### EXCEL ファイルの読み込み

```json
{
  "server_name": "xelixir",
  "tool_name": "read_excel",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "range": "A1:C10"
  }
}
```

### EXCEL ファイルへの書き込み

```json
{
  "server_name": "xelixir",
  "tool_name": "write_excel",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "data": [
      ["A1", "B1", "C1"],
      ["A2", "B2", "C2"]
    ]
  }
}
```

### セル範囲への書き込み（開始セル指定）

`write_excel` は A1 起点固定ですが、開始セルを指定してピンポイントに書き込む場合は [`tool_write_range()`](mcp/xelixir.py:282) を使います。

- `start_cell` を左上として `data`（2次元配列）を書き込みます（上書き）。
- 文字列が `"="` で始まる場合は **数式として設定** します。
- `null` は **空白セル** になります。

```json
{
  "server_name": "xelixir",
  "tool_name": "write_range",
  "arguments": {
    "path": "/path/to/file.xlsx",
    "sheet_name": "Sheet1",
    "start_cell": "D5",
    "data": [
      [1, 2, "=SUM(A1:B1)"],
      [3, null, "text"]
    ]
  }
}
```

### 行の追記（append）

同じファイルに対して「毎回 A1 から書き直す」のを避けたい場合は [`tool_append_rows()`](mcp/xelixir.py:317) を使います。

- `anchor_column`（例: `"A"`）を **上から走査** し、最初に空（未定義/BLANK/空文字）になっている行に追記します。
- `rows` は **2次元配列（行の配列）** です。
- 文字列が `"="` で始まる場合は **数式として設定** します。
- `null` は **空白セル** になります。

```json
{
  "server_name": "xelixir",
  "tool_name": "append_rows",
  "arguments": {
    "path": "/path/to/file.xlsx",
    "sheet_name": "Sheet1",
    "anchor_column": "A",
    "rows": [
      ["2025-12-22", "Alice", 100],
      ["2025-12-23", "Bob", 200]
    ]
  }
}
```

### 新しいシートの作成

```json
{
  "server_name": "xelixir",
  "tool_name": "create_sheet",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "NewSheet"
  }
}
```

### 新しい EXCEL ファイルの作成

```json
{
  "server_name": "xelixir",
  "tool_name": "create_excel",
  "arguments": {
    "filePath": "/path/to/new_file.xlsx",
    "sheetName": "Sheet1"  // 省略可、デフォルトは "Sheet1"
  }
}
```

### ワークブックのメタデータ取得

```json
{
  "server_name": "xelixir",
  "tool_name": "get_workbook_metadata",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "includeRanges": false  // 省略可、範囲情報を含めるかどうか
  }
}
```

### シート名の変更

```json
{
  "server_name": "xelixir",
  "tool_name": "rename_worksheet",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "oldName": "Sheet1",
    "newName": "NewName"
  }
}
```

### シートの削除

```json
{
  "server_name": "xelixir",
  "tool_name": "delete_worksheet",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1"
  }
}
```

### シートのコピー

```json
{
  "server_name": "xelixir",
  "tool_name": "copy_worksheet",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sourceSheet": "Sheet1",
    "targetSheet": "Sheet1Copy"
  }
}
```

### セルへの数式適用

```json
{
  "server_name": "xelixir",
  "tool_name": "apply_formula",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "cell": "C1",
    "formula": "=SUM(A1:B1)"
  }
}
```

### 数式構文の検証

```json
{
  "server_name": "xelixir",
  "tool_name": "validate_formula_syntax",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "cell": "C1",
    "formula": "=SUM(A1:B1)"
  }
}
```

### セル範囲の書式設定

```json
{
  "server_name": "xelixir",
  "tool_name": "format_range",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "startCell": "A1",
    "endCell": "C3",
    "bold": true,
    "italic": false,
    "fontSize": 12,
    "fontColor": "#FF0000",
    "bgColor": "#FFFF00"
  }
}
```

### セルの結合

```json
{
  "server_name": "xelixir",
  "tool_name": "merge_cells",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "startCell": "A1",
    "endCell": "C1"
  }
}
```

### セルの結合解除

```json
{
  "server_name": "xelixir",
  "tool_name": "unmerge_cells",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "startCell": "A1",
    "endCell": "C1"
  }
}
```

### セル範囲のコピー

```json
{
  "server_name": "xelixir",
  "tool_name": "copy_range",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "sourceStart": "A1",
    "sourceEnd": "C3",
    "targetStart": "D1",
    "targetSheet": "Sheet2"  // 省略可、省略時は同じシート
  }
}
```

### セル範囲の削除

```json
{
  "server_name": "xelixir",
  "tool_name": "delete_range",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "startCell": "A1",
    "endCell": "C3",
    "shiftDirection": "up"  // "up" または "left"。デフォルトは "up"
  }
}
```

### Excel 範囲の検証

```json
{
  "server_name": "xelixir",
  "tool_name": "validate_excel_range",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "startCell": "A1",
    "endCell": "C3"  // 省略可
  }
}
```

### グラフの作成

```json
{
  "server_name": "xelixir",
  "tool_name": "create_chart",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "dataRange": "A1:C10",
    "chartType": "column",  // "column", "line", "bar", "area", "scatter", "pie"
    "targetCell": "E1",
    "title": "サンプルグラフ",  // 省略可
    "xAxis": "X軸ラベル",      // 省略可
    "yAxis": "Y軸ラベル"       // 省略可
  }
}
```

### ピボットテーブルのメタデータ作成

```json
{
  "server_name": "xelixir",
  "tool_name": "create_pivot_table",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "dataRange": "A1:D100",
    "rows": ["Category"],
    "values": ["Sales"],
    "columns": ["Region"],  // 省略可
    "aggFunc": "sum"  // "sum", "count", "average", "max", "min" など
  }
}
```

### シート名一覧の取得

```json
{
  "server_name": "xelixir",
  "tool_name": "list_sheets",
  "arguments": {
    "filePath": "/path/to/book.xlsx"
  }
}
```

`list_sheets` ツールは次のような結果を返します:

```json
{
  "path": "/path/to/book.xlsx",
  "sheets": ["Sheet1", "Data", "Summary"]
}
```

## 開発・貢献方法

現時点では、専用の CONTRIBUTING ドキュメントは用意していませんが、
このリポジトリに変更を加える場合は次の点を意識してください。

- Java ツールが常に [`java/tools/compile.sh`](java/tools/compile.sh:1) でビルドできる状態を維持する
- 新しい Java ツールを追加した場合は、対応する Python ラッパーを [`mcp/src/excel`](mcp/src/excel/__init__.py:1) に追加し、
  MCP ツール登録（[`mcp/xelixir.py`](mcp/xelixir.py:1)）も忘れずに行う
- ユーザー向けの挙動を変更・追加した場合は、[`README.md`](README.md:1) / [`README.JA.md`](README.JA.md:1) の両方を更新する
- コミットメッセージや PR 説明には、英語または日本語で変更内容と背景を簡潔に記載する

Issue や Pull Request は歓迎します。

## ライセンス

このリポジトリには、Apache License 2.0 の全文を含む [`LICENSE`](LICENSE:1) ファイルが含まれています。
特に明記がない限り、プロジェクト全体はこのライセンスの条件に従って利用されることを意図しています。

> 補足: [`mcp/pyproject.toml`](mcp/pyproject.toml:7) の `license` フィールドは現時点で "MIT" を示していますが、
> リポジトリ直下の [`LICENSE`](LICENSE:1) を正とみなし、今後のバージョンで整合を取る予定です。

## 注意事項

- ファイルパスは絶対パスで指定すること
- シート名を省略できるツールでは、省略した場合に最初のシートが対象になる
- 範囲指定は "A1:C10" のような A1 形式で記述する
- `create_excel` で既存ファイルパスを指定するとエラーになる（既存ファイルの上書きは行わない）
- ピボットテーブル機能は、現時点ではメタデータ構築のみで、実際の Excel ピボットテーブルオブジェクトを作成しない実装になっている場合があります
