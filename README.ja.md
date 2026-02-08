# pptx-templatex

PowerPointテンプレートエンジン - スライドのコピーとプレースホルダー置換機能を提供するPythonライブラリ

## 機能

- PowerPointファイル（.pptx）からスライドをコピー
- `{{ }}` で囲まれたプレースホルダーを置換
- ネストされたオブジェクトへのアクセス（例: `{{ user.name }}`）
- 配列要素へのアクセス（例: `{{ items[0].name }}`）
- JSON設定ファイルによる一括処理

## インストール

```bash
pip install -e .
```

開発用の依存関係も含める場合:

```bash
pip install -e ".[dev]"
```

## クイックスタート

1. プレースホルダー付きのPowerPointテンプレートを作成：
   - PowerPointを開いてプレゼンテーションを作成
   - `{{ name }}`、`{{ user.email }}`、`{{ items[0].title }}` のようなプレースホルダーをテキストボックスに追加
   - `template.pptx` として保存

2. JSON設定ファイル (`config.json`) を作成：
```json
{
  "slides": [
    {
      "src_page": 1,
      "replace_texts": {
        "name": "山田太郎",
        "user": {
          "email": "taro@example.com"
        },
        "items": [
          {"title": "最初のアイテム"}
        ]
      }
    }
  ]
}
```

3. コマンドを実行：
```bash
pptx-templatex template.pptx config.json output.pptx
```

## 使い方

### コマンドラインインターフェース

インストール後、`pptx-templatex` コマンドが使えます：

```bash
# 基本的な使い方
pptx-templatex template.pptx config.json output.pptx

# ヘルプを表示
pptx-templatex --help

# バージョンを表示
pptx-templatex --version
```

### Python API

```python
from pptx_templatex import TemplateEngine

# テンプレートファイルを読み込む
engine = TemplateEngine("template.pptx")

# 設定を定義
config = {
    "slides": [
        {
            "src_page": 1,
            "replace_texts": {
                "name": "太郎",
                "title": "エンジニア"
            }
        }
    ]
}

# 処理して出力
engine.process(config, "output.pptx")
```

### ネストされたオブジェクトへのアクセス

```python
config = {
    "slides": [
        {
            "src_page": 1,
            "replace_texts": {
                "user": {
                    "name": "山田太郎",
                    "email": "taro@example.com"
                }
            }
        }
    ]
}
```

テンプレート内では:
```
名前: {{ user.name }}
メール: {{ user.email }}
```

### 配列要素へのアクセス

```python
config = {
    "slides": [
        {
            "src_page": 1,
            "replace_texts": {
                "items": [
                    {"name": "商品A", "price": "1000"},
                    {"name": "商品B", "price": "2000"}
                ]
            }
        }
    ]
}
```

テンプレート内では:
```
最初の商品: {{ items[0].name }} - {{ items[0].price }}円
二番目の商品: {{ items[1].name }} - {{ items[1].price }}円
```

### 複数のスライドを作成

```python
config = {
    "slides": [
        {"src_page": 1, "replace_texts": {"title": "イントロダクション"}},
        {"src_page": 2, "replace_texts": {"content": "メインコンテンツ"}},
        {"src_page": 1, "replace_texts": {"title": "まとめ"}},
    ]
}
```

### JSON設定ファイルを使用

`config.json` ファイルを作成：
```json
{
  "slides": [
    {
      "src_page": 1,
      "replace_texts": {
        "name": "太郎",
        "items": [{"key": "value"}]
      }
    }
  ]
}
```

CLIで使用：
```bash
pptx-templatex template.pptx config.json output.pptx
```

またはPython APIで使用：
```python
engine = TemplateEngine("template.pptx")
engine.process("config.json", "output.pptx")
```

### 実践的な例

**シナリオ**: 複数のユーザー向けにパーソナライズされたプレゼンテーションを生成

1. `template.pptx` を作成：
   - スライド1: タイトルスライドに `{{ name }}` と `{{ title }}`
   - スライド2: コンテンツスライドに `{{ company.name }}` と `{{ company.address }}`
   - スライド3: リストスライドに `{{ items[0].description }}`

2. `config.json` を作成：
```json
{
  "slides": [
    {
      "src_page": 1,
      "replace_texts": {
        "name": "山田花子",
        "title": "シニアデベロッパー"
      }
    },
    {
      "src_page": 2,
      "replace_texts": {
        "company": {
          "name": "テック株式会社",
          "address": "東京都渋谷区1-2-3"
        }
      }
    },
    {
      "src_page": 3,
      "replace_texts": {
        "items": [
          {"description": "パフォーマンスを50%向上"},
          {"description": "バグを30%削減"}
        ]
      }
    }
  ]
}
```

3. 生成：
```bash
pptx-templatex template.pptx config.json hanako_presentation.pptx
```

これで、すべてのプレースホルダーが置き換えられた3枚のスライドを持つプレゼンテーションが作成されます。

## 設定フォーマット

### 設定オブジェクトの構造

```json
{
  "slides": [
    {
      "src_page": 1,
      "replace_texts": {
        "key": "value"
      }
    }
  ]
}
```

- `slides` (必須): スライド設定の配列
  - `src_page` (必須): コピー元のスライド番号（1から始まる）
  - `replace_texts` (オプション): 置換するテキストのマッピング

### プレースホルダーの記法

- 単純な置換: `{{ key }}`
- ネストされたキー: `{{ user.name }}`, `{{ company.department.name }}`
- 配列アクセス: `{{ items[0] }}`, `{{ users[0].name }}`
- 複雑なパス: `{{ company.departments[0].teams[1].name }}`

## テスト

```bash
pytest
```

カバレッジレポート付き:

```bash
pytest --cov=pptx_templatex --cov-report=html
```

## プロジェクト構造

```
pptx-templatex/
├── pptx_templatex/
│   ├── __init__.py
│   ├── template_engine.py      # メインのテンプレートエンジン
│   ├── placeholder_replacer.py  # プレースホルダー置換ロジック
│   └── exceptions.py            # カスタム例外
├── tests/
│   ├── __init__.py
│   ├── test_template_engine.py
│   └── test_placeholder_replacer.py
├── examples/
│   ├── example_usage.py
│   └── config.json
├── pyproject.toml
└── README.md
```

## ライセンス

MIT

## 作者

SpringMT
