# 概要 - エクセル(xlsx)処理API機能設計

## 実装言語と主なライブラリ

- Python
    - FastAPI
    - openpyxl
    - pytest

## 関数

### エクセル変換（http受け口）

イメージしやすいように受け口部分の内容です。設計書が必要なわけではありません。

httpから実行されたPOST処理を受け付け、変換後のエクセルを戻す。

- 関数名: transform_excel

処理

1. httpのPOST処理を受け付ける。
2. 引数に指定されたエクセルを、処理内容に従い変換する。
3. 変換されたエクセルを戻す。

インプット（Body）

```json
{
    "file": "<BASE64 Excel>",
    "mimetype": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "operations": [
        {
            "sheet_name": "<シート名>",
            "processing": [
                {
                    "processing_type": "",
                    "target": {
                        "cells": {
                            "start_cell": {
                                "col_letter": "A",  # 行指定の場合不要
                                "row": 1  # 列指定の場合不要
                            },
                            "end_cell": {
                                "col_letter": "B",
                                "row": 3
                            },  # 範囲の場合はend_cellも指定
                        },
                        "values": [
                            [1, 11],
                            ["2", 22],
                            [3, "B3"],
                        ],  # 行の配列 / 行数又は列数が一致しない場合エラー
                        "styles": {}  # 罫線（個別又は範囲全体の外周指定可）、背景色、文字色、フォント、フォントサイズ、太字、イタリック、アンダーライン、行の高さ、列の幅
                    },
                    "paste_target": {  # コピーパターンのみ
                        "sheet_name": "",
                        "cells": {
                            "starting_point": {
                                "col_letter": "C",
                                "row": 1
                            },  # 起点のみ指定
                        },
                        "is_insert": True,  # 挿入処理かコピー処理か
                    },
                }
            ]
        }
    ]
}
```

operationsの順番で、各シートにprocessingの内容を適用する。

- processing_type
    - copy: cellsから行指定、列指定、セル指定を判定する。
    - copy_sheet: 指定した名前でシートのコピー
    - copy_style: cellsから行指定、列指定、セル指定を判定する。

    - insert_sheet: targetは空。対象のシートが存在せず、初回がinsert_sheetでない場合はエラー。
    - delete_sheet: targetは空。
    - insert: 行又は列のみ。
    - delete: 行又は列のみ。
    - hidden: 行又は列のみ。

    - set_cells: values及びstylesに指定された値を設定する。valuesとstylesの両方がない場合はエラー。
    - join_cells: セルを結合する。

アウトプット（Response Body）

```json
{
    "output": "<BASE64 Excel>",
    "mimetype": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "status": "Success",
    "error_code": 200,
    "status_code": 200,
}
```

エラー発生時（Response Body）

```json
{
    "output": "Bad Request: Syntax error in the request body.",
    "status": "Error",
    "error_code": 400,
    "status_code": 400,
}
```

### オペレーション処理

オペレーションひとまとまり毎（operationsの1つ）に処理を行う。

- 関数名: xlsx_operation

処理

1. 指定されたシート名に対してオペレーション内容を適用する。
2. 変換されたエクセルを戻す。

インプット

| 引数名 | 型 | 概要 |
| :--- | :--- | :--- |
| excel_book | io.BytesIO | 処理対象エクセル |
| sheet_name | str | 処理シート名 |
| processing | list | 処理内容 |

アウトプット

| 型 | 概要 |
| :--- | :--- |
| io.BytesIO | 処理反映後のエクセル |

### オペレーションごとの関数

実装後記載想定
