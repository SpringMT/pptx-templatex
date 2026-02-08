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

## トラブルシューティング・開発ノート

開発中に遭遇した問題と対応方法をまとめています。

### 1. スライドコピーの問題

**問題**: 通常のshapeコピーではレイアウトや書式が正しく保持されない

**対応**: XMLレベルで`deepcopy`を使用してshape要素を完全に複製
```python
from copy import deepcopy
new_element = deepcopy(shape.element)
dest_slide.shapes._spTree.insert_element_before(new_element, "p:extLst")
```

### 2. 画像コピーの問題

**問題**: python-pptx 1.0.2で`get_or_add_image_part()`のAPIが変更され、戻り値が`(image_part, rId)`のタプルになった

**対応**: 旧APIと新APIの両方に対応する実装
```python
result = dest_part.get_or_add_image_part(image_stream)
if isinstance(result, tuple):
    # 新API (python-pptx 1.0.2+)
    new_image_part, new_rId = result
else:
    # 旧API
    new_image_part = result
```

画像のrId（関係ID）をマッピングしてXML内の参照を更新：
```python
# rIdマッピングを作成
rId_mapping[old_rId] = new_rId

# XML内のblip要素を更新
blip.set(f'{{{r_ns}}}embed', rId_mapping[old_rId])
```

### 3. フォントと色の保持の問題

**問題1**: `Font name: None`のテキストランがコピー後にデフォルトフォント（MS Gothic）になる

**対応**: スライド内の他のshapeから定義済みフォントを収集し、`None`のrunに適用
```python
# スライド全体からフォントを収集
slide_fonts = set()
for shape in slide.shapes:
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            if run.font.name is not None:
                slide_fonts.add(run.font.name)

# Noneのrunに適用
reference_font_name = slide_fonts.pop() if slide_fonts else "Meiryo UI"
```

**問題2**: プレースホルダー置換時にフォント・色情報が失われる

**対応**: テキスト置換前にフォント情報を保存し、置換後に復元
```python
# 保存
ref_font_name = reference_run.font.name
ref_color_rgb = reference_run.font.color.rgb

# テキスト置換
paragraph.clear()
new_run = paragraph.add_run()
new_run.text = new_text

# 復元
new_run.font.name = ref_font_name
new_run.font.color.rgb = ref_color_rgb
```

### 4. プレースホルダーが複数runに分割される問題

**問題**: PowerPointは`{{`や`[`などの特殊文字で自動的にrunを分割するため、`{{ name }}`が3つのrunに分かれる
```
Run 1: '{{'
Run 2: 'name'
Run 3: '}}'
```

**対応**: 段落全体のテキストを取得して一括置換し、新しい単一のrunとして設定
```python
full_text = paragraph.text  # 全runを結合したテキスト
new_text = PlaceholderReplacer.replace_text(full_text, replacements)

# 全runを削除して新しいrunを作成
paragraph.clear()
new_run = paragraph.add_run()
new_run.text = new_text
```

### 5. 制御文字（`_x000B_`）の問題

**問題**: 段落内の改行（vertical tab: `\x0B`）が`_x000B_`という文字列として表示される

**対応**: 垂直タブを改行に変換し、その他の不要な制御文字のみを削除
```python
# 垂直タブを改行に変換
new_text = new_text.replace('\x0B', '\n')
# その他の制御文字を削除（\nと\rは保持）
new_text = re.sub(r'[\x00-\x08\x0C\x0E-\x1F]', '', new_text)
```

### 6. テーマフォントの問題

**問題**: テンプレートで`+mj-lt`（テーマのメジャーラテンフォント）などのテーマフォントを使用している場合、新規プレゼンテーションにコピーするとテーマ情報が失われる

**対応**: テンプレートファイルをベースに出力プレゼンテーションを作成
```python
# 空のプレゼンテーションではなく、テンプレートから作成
output_prs = Presentation(str(template_path))

# テンプレートの全スライドを削除
while len(output_prs.slides) > 0:
    rId = output_prs.slides._sldIdLst[0].rId
    output_prs.part.drop_rel(rId)
    del output_prs.slides._sldIdLst[0]
```

これにより、テーマ（フォント、色、エフェクト）、スライドマスター、レイアウトが全て保持されます。

### 7. カスタムレイアウトの問題

**問題**: 「ユーザー設定レイアウト_1_1」のようなカスタムレイアウトが`prs.slide_layouts`のリストに存在せず、レイアウトが見つからない

**対応**: スライドマスター経由でレイアウトを検索
```python
# スライドのレイアウトからマスターを取得
source_master = source_layout.slide_master

# ターゲットプレゼンテーションで同じ名前のマスターを検索
for master in target_prs.slide_masters:
    if master.name == source_master.name:
        target_master = master
        break

# マスター内でレイアウトを検索
for layout in target_master.slide_layouts:
    if layout.name == source_layout.name:
        slide_layout = layout
        break
```

### デバッグツール

問題の調査に使用したデバッグツールを`tools/`ディレクトリに保存しています：

- **`analyze_text_format.py`**: フォント名、サイズ、色などのテキスト書式情報を分析
  ```bash
  python3 tools/analyze_text_format.py template.pptx
  ```

- **`debug_layout.py`**: スライドマスター、レイアウト、各スライドが使用しているレイアウトを表示
  ```bash
  python3 tools/debug_layout.py template.pptx
  ```

## ライセンス

MIT

## 作者

SpringMT
