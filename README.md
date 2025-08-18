# このリポジトリについて

このリポジトリはExcel, PowerPoint, PDF, textのような多種多様な文書をマークダウンに変換する、[markitdown](https://github.com/microsoft/markitdown) の実験用リポジトリです

## 構成

```bash
document-to-markdown/
├── README.md
├── pyproject.toml
├── src
│   ├── media
│   │   └── sample-multilingual-text.pdf   # サンプルPDF
│   ├── convert-all.py                      # Excel, PDF, PowerPoint, Word, Text変換用
│   ├── excel-to-markdown.py                # Excel変換用
│   ├── pdf-to-markdown.py                  # PDF変換用
│   ├── powerpoint-to-markdown.py           # PowerPoint変換用
│   └── word-to-markdown.py                 # Word変換用
└── uv.lock
```

## 利用ライブラリ

markitdown[all]==0.1.2

## セットアップ

```bash
# ディレクトリ移動
cd document-to-markdown

# 仮想環境作成
uv venv --python 3.12.3

# 仮想環境有効化
source .venv/bin/activate

# ライブラリインストール
uv sync
```

## 使い方

### Excelファイルの変換

```bash
# 標準出力に表示
uv run src/excel-to-markdown.py 'EXCEL-FILE-PATH'

# ファイルに出力
uv run src/excel-to-markdown.py 'CONVERTED-MARKDOWN-FILE-PATH'
```

### PDFファイルの変換

```bash
# 標準出力に表示
uv run src/pdf-to-markdown.py 'PDF-FILE-PATH'

# ファイルに出力
uv run src/pdf-to-markdown.py 'CONVERTED-MARKDOWN-FILE-PATH'
```

### PowerPointファイルの変換

```bash
# 標準出力に表示
uv run src/powerpoint-to-markdown.py 'POWERPOINT-FILE-PATH'

# ファイルに出力
uv run src/powerpoint-to-markdown.py 'CONVERTED-MARKDOWN-FILE-PATH'
```

### Wordファイルの変換

```bash
# 標準出力に表示
uv run src/word-to-markdown.py 'WORD-FILE-PATH'

# ファイルに出力
uv run src/word-to-markdown.py 'CONVERTED-MARKDOWN-FILE-PATH'
```

### Excel, PDF, PowerPoint, Word, Textファイルの変換

```bash
# ディレクトリ内の全ファイルを変換
uv run src/convert-all.py --directorypath src/media

# 単一ファイルを変換
uv run src/convert-all.py --filepath src/media/document.pdf

# カスタム出力ディレクトリ指定
uv run src/convert-all.py --directorypath src/media --output custom-output
```

### 基本的な使用方法

各スクリプトは以下の形式で実行します：

```bash
uv run src/[スクリプト名] [入力ファイルパス] [出力ファイルパス（オプション）]
```

- 出力ファイルパスを指定しない場合は標準出力に表示されます
- ファイル名にスペースや特殊文字が含まれる場合はダブルクォートで囲んでください

## 参考資料

[markitdown](https://github.com/microsoft/markitdown)