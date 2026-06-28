# a4h2p_spec.md — liber8_ppter 技術仕様書
# Absolute HTML to PPTX — How（記述方法・変換ルール）
# AIへの指示例：「a4h2p_spec.mdの約束事を遵守してスライドHTMLを生成してください」

更新日: 2026-06-16

---

## ドキュメント責務の分離（SSOT原則）

| ドキュメント | 責務 |
|---|---|
| **a4h2p_spec.md**（本ファイル） | PPTX変換の共通ルール（サイズ換算・フォント対応・図形タイプ・座標計算ロジック） |
| **design_catalogue.dc.html** | 各 Molecule / Organism の個別数値定義（位置・幅・スペーシング・バリアント仕様） |

> **原則**: Molecule・Organism の正確な座標値・スペーシング値・バリアント仕様は  
> 常に **design_catalogue.dc.html を唯一の正（Single Source of Truth）** として参照すること。  
> 本 spec.md に個別数値を転記しない。転記値が存在する場合は Catalogue 側を優先する。

---

## コード注釈ルール（Sample ファイル共通）

HTMLサンプルファイル（`a4h2p_sample.html` / `A4H2P Sample v2.dc.html`）では以下の注釈ルールを統一する。

### ① セクションバナー（スライド単位）
```html
<!-- ============================================================ {SlideID} / {LayoutCode} / {使用Organism・Molecule} ============================================================ -->
```
- **SlideID**: `B07`、`B07b` など（data-label と一致）
- **LayoutCode**: `L-1`（1カラム）、`L-2H`（横2分割）など
- **使用M/O**: `M-B0 + M-B3`、`O-B5` など（複数はスペース区切り `+`）

### ② インラインコメント（コンポーネント単位）
```html
<!-- M-Bx 説明 -->
<!-- O-Bx 説明 -->
```
- 各コンポーネントの **最初の要素の直前** に1行だけ置く
- 同一コンポーネントが繰り返す場合は **最初の1回のみ** コメントし、繰り返し要素には `<!-- (繰り返し) -->` を省略可

### 対応表（主要コンポーネント）
| コード | 名称 |
|---|---|
| M-B0 | ヘッドラインバンド |
| M-B1 | スライドヘッダー（タイトル＋ページ番号＋下線） |
| M-B2 | アクセントバー（縦 or 横） |
| M-B3 | 箇条書き（L1 / L2） |
| M-B4 | KPIカード |
| M-B6 | アジェンダ行（時刻＋縦バー＋テキスト） |
| M-B10 | プロフィールカード |
| M-B11 | ゾーン見出し（縦バー＋テキスト） |
| O-B1 | 比較テーブル |
| O-B2 | タスク進捗テーブル |
| O-B3 | 評価比較テーブル |
| O-B5 | 横型フロー＋グルーピング |
| O-B7 | 乗算式フロー（バリアントA/B） |

## 0. サイズ換算リファレンス

### 0-1. キャンバスとPPTXの関係

| 項目 | HTML | PPTX |
|---|---|---|
| スライド幅 | 1920 px | 25.4 cm |
| スライド高さ | 1080 px | 14.29 cm |
| スケール比 | ÷2 | = PPTX px |
| 1cm | 75.6 px | 37.8 px |

```
HTML px ÷ 2 = PPTX px
PPTX px ÷ 37.8 = cm
HTML px ÷ 75.6 = cm
```

---

### 0-2. フォントサイズ換算（実測ベース）

```
PPTX pt × 8/3（≈2.667） = HTML px
HTML px × 0.375 = PPTX pt
```

**実測確認済みの対応表：**

| PPTX pt | HTML px | トークン名 | 用途 |
|---|---|---|---|
| 10pt | 27px | type-label | ラベル・著作権 |
| 12pt | 32px | type-caption | キャプション |
| 16pt | 43px | type-body-sm | 箇条書き第2階層 |
| 18pt | 48px | type-body | 本文（✅POTX実測）|
| 20pt | 53px | type-heading | 見出し |
| 22pt | 59px | type-title | スライドタイトル（✅POTX実測）|
| 38pt | 101px | type-display | 大見出し（✅POTX実測）|

---

### 0-3. テキストボックス高さの実測ロジック

**POTX実測値（Noto Sans JP・1行・上下余白なし・自動調整）：**

| PPTX pt | テキストボックス高さ(cm) | HTML px換算 |
|---|---|---|
| 22pt（type-title） | 0.94 cm | 71 px |
| 38pt（type-display） | 1.63 cm | 123 px |

**比率の導出：**
```
0.94 ÷ 22 = 0.04273 cm/pt
1.63 ÷ 38 = 0.04289 cm/pt
→ 平均: pt × 0.0427 cm = テキストボックス高さ
```

**HTML px への換算：**
```
テキストボックス高さ(px) = PPTX pt × 0.0427 × 75.6
                        ≈ PPTX pt × 3.23
                        = HTML px × 1.6 ÷ 行数 で逆算も可
```

**設計時の高さ計算式（推奨）：**
```
height = font-size(px) × 行数

例: type-title 1行 → 59 × 1 = 59px

※ 旧式（× 1.6）は廃止。PPTXの行間「1行」= line-height:font-size と等価のため。
```

**テキストボックスのTop位置計算ルール：**
```
text.top = slot.top + (slot.height − font-size) / 2

例: スロットrect top:490 h:77, font-size:43
  → text top = 490 + (77 − 43) / 2 = 507px

例: バナーrect top:730 h:100, font-size:48
  → text top = 730 + (100 − 48) / 2 = 756px

※ スライドタイトル（font-size:59）は上端8px固定（top:8）。上記式の適用外。

---

### 0-4. 自由演技サイズの選び方

トークン外の大きな数値（KPI数値・セクション番号など）を選ぶ際の基準：

| 視覚的な役割 | 推奨 HTML px | PPTX pt相当 | 根拠 |
|---|---|---|---|
| KPI数値（メイン） | 101px | 38pt | type-displayと同等 |
| KPI数値（サブ） | 75px | 28pt | type-titleの1.3倍 |
| セクション番号 | 160〜240px | 60〜90pt | A系コンテンツ高さの1/4〜1/6 |
| 引用符・装飾文字 | 160px | 60pt | — |

**テキストボックス高さ逆算：**
```
自由演技サイズ N px のテキストボックス高さ = N × 行数
例: 101px（KPI数値、1行） → 高さ = 101px
例: 48px（本文、2行）    → 高さ = 96px
```

---

## 1. スライドサイズ・座標系

| 項目 | 値 |
|---|---|
| HTMLキャンバス | 1920 × 1080 px（16:9） |
| PPTX実寸 | 25.4 × 14.29 cm（960 × 540 px） |
| python-pptx設定 | `prs.slide_width = Pt(1920)` / `prs.slide_height = Pt(1080)` |
| 座標指定 | 全要素 `position:absolute` + `left/top/width/height` をpx明示 |
| 座標変換 | px値を `Pt(n)` にそのまま渡す |
| `right` / `bottom` | 使用禁止（左上起点のみ） |

---

## 1-B. ボディゾーン レイアウトテンプレート

スライドのBody領域を分割する際の座標テンプレート。  
AIはここから該当レイアウトを選び、各ゾーンの `left/top/width/height` をそのまま使用する。

### 基準値

| 名称 | 値 | 説明 |
|---|---|---|
| SAFE_L | 38 | 左セーフマージン |
| SAFE_R | 1882 | 右端（1920−38） |
| SAFE_B | 1042 | 下端（1080−38） |
| BODY_TOP | 136 | Body開始Y（ディバイダー後98px + 上パディング38px） |
| BODY_W | 1844 | Body幅（1882−38） |
| BODY_H | 906 | Body高さ（1042−136） |
| GAP | 40 | ゾーン間の余白 |

> **ヘッダーなしスライド（表紙等）** の場合は BODY_TOP=38・BODY_H=1004 で計算する。
> 構造上の最小値：ディバイダー直下 = 98px。BODY_TOP=136 はこれに上パディング38pxを加えた実用値。
> **M-B0ヘッドラインあり** の場合は BODY_TOP=272（ヘッドライン下端262px + 余白10px）・BODY_H=774（1046-272）で計算する。

---

### レイアウト一覧

#### L-1: Full（1ゾーン）
| ゾーン | left | top | width | height |
|---|---|---|---|---|
| A | 38 | 136 | 1844 | 906 |

---

#### L-2H: 左右2分割 50/50
| ゾーン | left | top | width | height |
|---|---|---|---|---|
| Left  | 38  | 136 | 902 | 906 |
| Right | 980 | 136 | 902 | 906 |

#### L-2H-4060: 左右2分割 40/60
| ゾーン | left | top | width | height |
|---|---|---|---|---|
| Left  | 38  | 136 | 697  | 906 |
| Right | 775 | 136 | 1107 | 906 |

#### L-2H-6040: 左右2分割 60/40
| ゾーン | left | top | width | height |
|---|---|---|---|---|
| Left  | 38   | 136 | 1107 | 906 |
| Right | 1185 | 136 | 697  | 906 |

---

#### L-2V: 上下2分割 50/50
| ゾーン | left | top | width | height |
|---|---|---|---|---|
| Top    | 38 | 136 | 1844 | 433 |
| Bottom | 38 | 609 | 1844 | 433 |

---

#### L-3H: 左中右3分割 33/33/33
| ゾーン | left | top | width | height |
|---|---|---|---|---|
| Left   | 38   | 136 | 588 | 906 |
| Center | 666  | 136 | 588 | 906 |
| Right  | 1294 | 136 | 588 | 906 |

---

#### L-4: 上下左右4分割
| ゾーン | left | top | width | height |
|---|---|---|---|---|
| Top-Left     | 38  | 136 | 902 | 433 |
| Top-Right    | 980 | 136 | 902 | 433 |
| Bottom-Left  | 38  | 609 | 902 | 433 |
| Bottom-Right | 980 | 609 | 902 | 433 |

---

#### L-2H-Rs: 左右2分割・右だけ上下分割
| ゾーン | left | top | width | height |
|---|---|---|---|---|
| Left         | 38  | 136 | 902 | 906 |
| Right-Top    | 980 | 136 | 902 | 433 |
| Right-Bottom | 980 | 609 | 902 | 433 |

#### L-2H-Ls: 左右2分割・左だけ上下分割
| ゾーン | left | top | width | height |
|---|---|---|---|---|
| Left-Top    | 38  | 136 | 902 | 433 |
| Left-Bottom | 38  | 609 | 902 | 433 |
| Right       | 980 | 136 | 902 | 906 |

---

### 使い方

1. スライドの内容から最適なレイアウト名を選ぶ（例: `L-2H-Rs`）
2. 各ゾーンの座標をそのまま `position:absolute; left:…px; top:…px; width:…px; height:…px;` に転記
3. 各ゾーン内のコンテンツ要素は、そのゾーンの `left/top` を起点に配置する

```html
<!-- 例: L-2H-Rs（左 | 右上/右下）-->
<!-- Left zone -->
<div style="position:absolute; left:38px; top:136px; width:902px; height:906px; …">…</div>
<!-- Right-Top zone -->
<div style="position:absolute; left:980px; top:136px; width:902px; height:433px; …">…</div>
<!-- Right-Bottom zone -->
<div style="position:absolute; left:980px; top:609px; width:902px; height:433px; …">…</div>
```

---

## 2. スライド検出・レイアウト指定

```html
<section
  data-pptx-layout="04_body"
  data-label="03 LT発表一覧"
  style="position:relative; width:1920px; height:1080px; overflow:hidden;
         font-family:'Noto Sans JP',sans-serif; background:#FFFFFF;">
```

```python
for section in soup.select('section[data-pptx-layout]'):
    layout_name = section['data-pptx-layout']
    label = section.get('data-label', '')

    # レイアウト名で検索・フォールバックはblank
    layout = next(
        (l for l in prs.slide_layouts if l.name == layout_name),
        prs.slide_layouts[6]
    )
    slide = prs.slides.add_slide(layout)
```

**注意:**
- `data-pptx-background` は廃止（使用しない）
- スライド背景色はPPTレイアウト側で管理
- CSS `background` はHTML描画用のみ（白以外の場合に記述）

---

## 3. data-pptx-type 一覧

### 3-1. text

```html
<div data-pptx-type="text"
     style="position:absolute; left:38px; top:0px; width:560px; height:95px;
            font-family:'Noto Sans JP',sans-serif; font-size:59px; font-weight:700;
            color:#0088C8; line-height:95px; text-align:left; letter-spacing:0;">
  LT 発表一覧
</div>
```

**変換先:** `slide.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h))`

| CSS属性 | python-pptx | 備考 |
|---|---|---|
| `left/top/width/height` | 座標・サイズ | px → Pt() |
| `font-family` | `r.font.name` | CSS先頭フォント名。汎用ファミリー名（sans-serif等）は無視 |
| `font-size` | `r.font.size = Pt(css_px × 0.375)` | **px→pt変換: ×0.375** |
| `font-weight:700` | `r.font.bold = True` | それ以外は False |
| `color` | `r.font.color.rgb` | #RRGGBBのみ |
| `text-align:left` | `PP_ALIGN.LEFT` | デフォルト（原則） |
| `text-align:center` | `PP_ALIGN.CENTER` | **例外許容：KPIカードのKPI名ラベル・乗算式フロー（Molecule）の円内ラベル・ポジショニングマップの軸ラベルおよびプロット要素ラベル・比較テーブル（O-B1/O-B3）の列ヘッダー名および評価記号セル（◎△○）** |
| `text-align:right` | `PP_ALIGN.RIGHT` | |
| `line-height` | **無視** | heightで位置制御するため |
| `letter-spacing` | **無視** | python-pptxで未サポート |
| テキスト内容 | `el.decode_contents()` を `<br>` で分割 | 各行を段落として追加 |

**⚠ フォントサイズ換算（liber8_ppter実測値）:**
```
HTML px × 0.375 = PPTX pt
例: 59px × 0.375 = 22.125pt ≈ 22pt（type-title）
例: 48px × 0.375 = 18pt（type-body）
例: 101px × 0.375 = 37.875pt ≈ 38pt（type-display）
```
※ 旧spec（×0.75）はHTMLキャンバス960px時の値。liber8_ppterは1920pxキャンバスのため×0.375を使用。

---

### 3-2. rect

```html
<div data-pptx-type="rect"
     style="position:absolute; left:38px; top:136px; width:520px; height:400px;
            background:#D6EEF8;"></div>
```

**変換先:** `slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Pt(x), Pt(y), Pt(w), Pt(h))`

| CSS属性 | python-pptx | 備考 |
|---|---|---|
| `background` / `background-color` | `shape.fill.solid()` + `fore_color.rgb` | |
| 枠線 | `shape.line.fill.background()` | 常に非表示 |

---

### 3-3. line

```html
<div data-pptx-type="line"
     style="position:absolute; left:38px; top:94px; width:1844px; height:2px;
            background:#F0F0F0;"></div>
```

**変換先:** `add_shape(RECTANGLE)` — rect と同じ処理
**識別方法:** `height ≤ 4px` で line と判断（typeの明示が優先）

---

### 3-4. oval

```html
<div data-pptx-type="oval"
     style="position:absolute; left:38px; top:313px; width:20px; height:20px;
            background:#0088C8;"></div>
```

**変換先:** `slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, Pt(x), Pt(y), Pt(w), Pt(h))`

| CSS属性 | python-pptx |
|---|---|
| `background` | `shape.fill.solid()` + `fore_color.rgb` |
| `border-radius` | **実装なし・完全無視**（Converterはこの属性を読まない） |

---

### 3-5. triangle（等辺三角形）

```html
<div data-pptx-type="triangle"
     data-pptx-rotation="90"
     style="position:absolute; left:860px; top:570px; width:64px; height:64px;
            background:#0088C8;"></div>
```

**変換先:** `slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ISOSCELES_TRIANGLE, Pt(x), Pt(y), Pt(w), Pt(h))`

| 属性 | python-pptx | 備考 |
|---|---|---|
| `background` | `shape.fill.solid()` + `fore_color.rgb` | |
| 枠線 | `shape.line.fill.background()` | 常に非表示 |
| `data-pptx-rotation` | `shape.rotation = float(val)` | 省略時は `0`（上向き） |

**`data-pptx-rotation` 値と向き：**

| 値 | 向き | 主な用途 |
|---|---|---|
| `0` | ▲ 上向き（デフォルト） | — |
| `90` | ▶ 右向き | 横型フロー矢印 |
| `180` | ▼ 下向き | — |
| `270` | ◀ 左向き | — |

**矢印寸法規則（軸:矢頭高 = 3:5）：**

| SIZE | 軸高 h | 矢頭幅 w | 矢頭高 H |
|---|---|---|---|
| S | 22px | 24px | 36px |
| M ★標準 | 36px | 40px | 60px |
| L | 58px | 64px | 96px |
| 縦（↑） | 19px（軸長） | — | 32px |

縦方向矢印（向上指示）は accent4(#C87800) を使用。

**python-pptx処理：**
```python
def handle_triangle(slide, el, base_dir=None):
    s = parse_style(el.get('style', ''))
    shape = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ISOSCELES_TRIANGLE,
        Pt(px(s['left'])), Pt(px(s['top'])),
        Pt(px(s['width'])), Pt(px(s['height'])))
    bg = get_bg_color(s)
    if bg:
        shape.fill.solid()
        shape.fill.fore_color.rgb = bg
    shape.line.fill.background()
    rotation = float(el.get('data-pptx-rotation', '0'))
    if rotation:
        shape.rotation = rotation
```

**HTMLプレビュー用の推奨記述：**

PPTXの三角形はCSS `background` では視覚的に再現できないため、`clip-path` で補助描写する（Converterは `clip-path` を無視）。

```html
<!-- 右向き矢印の例 -->
<div data-pptx-type="triangle"
     data-pptx-rotation="90"
     style="position:absolute; left:860px; top:570px; width:64px; height:64px;
            background:#0088C8;
            clip-path: polygon(0% 0%, 100% 50%, 0% 100%);"></div>
```

---

### 3-6. image（URL参照）

GeminiがWeb上の公開画像を参照する場合に使用。
`src` が `http` で始まる場合、ConverterがダウンロードしてPPTXに埋め込む。

```html
<img data-pptx-type="image"
     src="https://example.com/product-photo.jpg"
     style="position:absolute; left:960px; top:98px; width:882px; height:700px;">
```

**変換先:** `slide.shapes.add_picture(str(img_path), Pt(x), Pt(y), Pt(w), Pt(h))`

| 属性 | python-pptx | 備考 |
|---|---|---|
| `src` | URLの場合は一時ファイルにダウンロード | `http` で始まる場合 |
| `width/height` | サイズ | px → Pt() |

```python
# Converter側: URL画像のダウンロード処理
import urllib.request, tempfile

def resolve_image(src: str, base_dir: Path) -> Path:
    if src.startswith('http'):
        suffix = Path(src.split('?')[0]).suffix or '.jpg'
        tmp = Path(tempfile.mktemp(suffix=suffix))
        urllib.request.urlretrieve(src, tmp)
        return tmp
    return (base_dir / src).resolve()
```

ダウンロード失敗時はグレー矩形プレースホルダーを配置し、警告を出して続行。

---

### 3-6. placeholder（画像プレースホルダー）

AIが「ここに画像があるべき」と判断した場合に使用。
Webアプリ側でユーザーにアップロードまたは画像生成を促す。

```html
<div data-pptx-type="placeholder"
     data-placeholder-label="製品イメージ"
     data-placeholder-hint="スマートフォンを持つビジネスパーソン・明るいオフィス・水平構図"
     style="position:absolute; left:960px; top:98px; width:882px; height:700px;
            background:#F0F0F0;">
  <!-- HTML仮描写: グレー背景 + ラベル表示 -->
  <div style="position:absolute; left:50%; top:50%; transform:translate(-50%,-50%);
              text-align:center;">
    <div style="width:80px; height:80px; background:#CCCCCC; margin:0 auto 16px;"></div>
    <div style="font-size:27px; color:#AAAAAA;">製品イメージ</div>
  </div>
</div>
```

**属性：**

| 属性 | 内容 |
|---|---|
| `data-placeholder-label` | Webアプリのアップロード画面に表示するラベル |
| `data-placeholder-hint` | 推奨サイズ・構図・内容のヒント（画像生成プロンプトにも転用可） |

**Converter側の処理：**
```python
def handle_placeholder(slide, el, base_dir=None):
    s = parse_style(el.get('style', ''))
    # グレー矩形を配置
    shape = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Pt(px(s['left'])), Pt(px(s['top'])),
        Pt(px(s['width'])), Pt(px(s['height'])))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(0xCC, 0xCC, 0xCC)
    shape.line.fill.background()
    # レポートに記録
    label = el.get('data-placeholder-label', '画像')
    hint  = el.get('data-placeholder-hint', '')
    missing_images.append({'label': label, 'hint': hint})
```

**変換後レポート例：**
```
⚠ 画像プレースホルダーが 2箇所あります：
  - Slide 03:「製品イメージ」
    ヒント: スマートフォンを持つビジネスパーソン・明るいオフィス・水平構図
  - Slide 07:「グラフ画像」
→ Webアプリ上でファイルをアップロードするか、画像生成をご利用ください。
```

**Webアプリのアップロードフロー：**
```
① HTML貼り付け
② placeholder 検出 → アップロードUI表示
   ┌─────────────────────────────────────┐
   │ 📷 製品イメージ（Slide 03）          │
   │ ヒント: ビジネスパーソン・オフィス  │
   │ [ファイルを選択] [AI画像生成]        │
   └─────────────────────────────────────┘
③ アップロード済み画像を埋め込んでPPTX生成
   （スキップ可 → グレー矩形のまま出力）
```

**要素構造のルール:**

```
section
├── div[data-pptx-type="rect"]         ← Converter処理
├── div[data-pptx-type="text"]         ← Converter処理
├── div[data-pptx-type="placeholder"]   ← Converter処理
│   ├── div（HTMLレビュー用他·data-pptx-typeなし） ← 完全無視
│   └── div（HTMLレビュー用他·data-pptx-typeなし） ← 完全無視
└── div[data-pptx-type="ignore"]       ← Converterスキップ
```

placeholderの内側にはグレー矩形・ラベルテキストなどHTMLプレビュー用の仕込みを自由に記述できる。
**ただし内側の要素に `data-pptx-type` を付けないこと。**
Converterは `section.select('[data-pptx-type]')` で取得するため、
空白の入れ子構造になっても二重処理は発生しない。

---

### 3-6. ignore

```html
<!-- ページ番号（PPTレイアウト管理） -->
<div data-pptx-type="ignore"
     style="position:absolute; left:1793px; top:26px; width:89px; height:43px;
            font-family:'Noto Sans JP',sans-serif; font-size:27px; color:#767676;
            line-height:43px; text-align:right;">
  01
</div>
```

**変換先:** スキップ（何もしない）

PPTレイアウト管理の要素に必ず付与する：
- 著作権フッター（全レイアウト共通）
- ページ番号（04_body）
- タイトル下ディバイダーライン（04_body）

---

### 3-7. ドロップシャドウ（data-pptx-shadow）

PPTプリセット1シャドウを使用する。`data-pptx-shadow="preset1"` 属性を追加する。

```html
<!-- シャドウあり -->
<div data-pptx-type="rect"
     data-pptx-shadow="preset1"
     style="position:absolute; left:38px; top:136px; width:520px; height:400px;
            background:#D6EEF8;
            box-shadow:3px 3px 5px rgba(0,0,0,0.40);"></div>
```

**PPTプリセット1の設定値：**

| 設定 | 値 | HTML換算 |
|---|---|---|
| 色 | #000000 | #000000 |
| 透明度 | 60% | opacity 40% → rgba(0,0,0,0.4) |
| サイズ | 100% | — |
| ぼかし | 4pt | 11px HTML |
| 角度 | 45° | offset-x: 6px / offset-y: 6px |
| 距離 | 3pt | 8px HTML |

**換算式：**
```
PPTX pt → HTML px: pt / 72 × 192（1920px ÷ 10inch）
3pt → 8px / 4pt → 11px
offset = 距離px × cos/sin(45°) ≈ 距離px × 0.707
```

**CSS（HTML仮描写用）：**
```css
box-shadow: 3px 3px 5px rgba(0,0,0,0.40);
```

**python-pptx処理：**
```python
def handle_shadow(shape):
    """data-pptx-shadow="preset1" → PPTプリセット1シャドウを適用"""
    shadow = shape.shadow
    shadow.inherit = False
    # プリセット1はXML直書きで適用
    # （python-pptx標準APIはinherit制御のみ）
    spPr = shape._element.find('.//' + qn('p:spPr'))
    # 必要に応じてXMLレベルでeffectLstを設定
```

※ python-pptxのshadow APIはinherit制御のみ対応。
プリセット1の詳細数値はPOTXテンプレート側で設定済みであれば継承される場合もある。

---

## 4. 未知のtypeの扱い

```python
handler = HANDLERS.get(ptype)
if handler:
    handler(slide, el, base_dir)
else:
    print(f'⚠ warn: 未知のdata-pptx-type="{ptype}" — スキップします')
```

---

## 5. Z順（重ね順）

- **DOM順 = Z順**（後の要素が前面）
- 背景rect → 前景rect → テキスト の順で記述

---

## 6. 色のフォーマット

- **`#RRGGBB` 形式のみ使用**
- `rgba()` / `hsla()` / 色名 / `#RGB` 短縮形 は使用禁止
- グラデーション全禁止（マテリアルデザイン準拠・奥行きはシャドウで表現）

```python
def hex_to_rgb(val: str) -> RGBColor:
    val = val.strip().lstrip('#')
    return RGBColor(int(val[0:2],16), int(val[2:4],16), int(val[4:6],16))
```

---

## 7. テキスト配置ルール

### 7-1. ボックス幅の計算

```
width = 推定テキスト幅 + font-size × 3（バッファ）

推定テキスト幅（Noto Sans JP 日本語）= 文字数 × font-size × 0.95
推定テキスト幅（Latin混在）= 日本語文字数 × font-size × 0.95
                            + Latin文字数 × font-size × 0.65
```

### 7-2. ボックス高さとTop位置の計算

**高さ（height）:**
```
height = font-size × 行数

例: 1行 / 48px → height = 48px
例: 2行 / 48px → height = 96px
例: 1行 / 59px → height = 59px
```

**Top位置:**
```
テキストのみ配置（スロットなし）: top = 配置したいY座標
スロット（rect）内に配置: top = slot.top + (slot.height − font-size) / 2

例: スロットrect top:490 h:77, font-size:43 → text top:507

※ スライドタイトル（font-size:59）は上端8px固定（top:8）。上記式の適用外。
```

line-height は常に font-size と同値にすること（PPTX行間「1行」と等価）。

### 7-3. text-align別のleft計算

**左揃え（left）— 最多用途**
```
left  = 配置したいX座標（通常はsafe-side=38px）
width = 推定テキスト幅 + font-size × 2
```

**中央配置（center）— 表紙・A系見出し**
```
W    = 推定テキスト幅 + font-size × 2
left = (1920 - W) / 2
```

**右揃え（right）— ページ番号**
```
W    = 推定テキスト幅 + font-size × 2
left = 1882 - W   （right端1882px基準: 1920 - safe-side 38px）
```

### 7-4. 複数行テキスト（brタグ）

```html
<!-- 2行: font-size:48px → height = 48 × 2 = 96px -->
<div data-pptx-type="text"
     style="position:absolute; left:38px; top:200px; width:1600px; height:96px;
            font-family:'Noto Sans JP',sans-serif; font-size:48px; font-weight:400;
            color:#1A1A1A; line-height:48px; text-align:left; letter-spacing:0;">
  1行目のテキストをここに記述する<br>
  2行目のテキストをここに記述する
</div>
```

```python
# パーサー側：<br>を段落分割として処理
inner = el.decode_contents()
lines = re.split(r'<br\s*/?>', inner, flags=re.IGNORECASE)

for i, line_html in enumerate(lines):
    text = BeautifulSoup(line_html, 'html.parser').get_text(strip=True)
    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
    p.alignment = align
    p.line_spacing = 1.0  # シングルスペース固定
```

---

## 8. CSSスタイルパースのユーティリティ

```python
def parse_style(style_str: str) -> dict:
    result = {}
    for item in (style_str or '').split(';'):
        item = item.strip()
        if ':' in item:
            k, v = item.split(':', 1)
            result[k.strip()] = v.strip()
    return result

def px(val_str: str) -> float:
    return float(val_str.replace('px', '').strip())

def font_pt(css_px_str: str) -> float:
    """48px → 18.0pt（liber8_ppter: × 0.375）"""
    return float(css_px_str.replace('px', '').strip()) * 0.375

def get_bg_color(style_dict: dict) -> RGBColor | None:
    val = style_dict.get('background') or style_dict.get('background-color') or ''
    val = val.strip()
    if val.startswith('#'):
        return hex_to_rgb(val)
    return None
```

---

## 9. パーサー骨格

```python
from pathlib import Path
from bs4 import BeautifulSoup
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
import re

prs = Presentation()
prs.slide_width  = Pt(1920)
prs.slide_height = Pt(1080)

with open('slides.html', encoding='utf-8') as f:
    soup = BeautifulSoup(f, 'html.parser')

HANDLERS = {
    'text':   handle_text,
    'rect':   handle_rect,
    'line':   handle_line,
    'oval':   handle_oval,
    'image':  handle_image,
    'ignore': lambda *_: None,
}

for section in soup.select('section[data-pptx-layout]'):
    layout_name = section['data-pptx-layout']
    layout = next(
        (l for l in prs.slide_layouts if l.name == layout_name),
        prs.slide_layouts[6]  # フォールバック: blank
    )
    slide = prs.slides.add_slide(layout)

    base_dir = Path('slides.html').parent
    for el in section.select('[data-pptx-type]'):
        ptype = el['data-pptx-type']
        handler = HANDLERS.get(ptype)
        if handler:
            handler(slide, el, base_dir)
        else:
            print(f'⚠ warn: 未知のtype="{ptype}" — スキップ')

prs.save('output.pptx')
```

---

## 10. ファイル構成（変換時）

```
📁 作業フォルダ/
  ├── slides.html       ← A4H2P形式HTMLファイル
  ├── converter.py      ← HTMLパーサー + PPTX生成スクリプト
  └── assets/           ← 画像・SVGファイル
```

---

## 11. よくあるミスと回避方法

| ミス | 原因 | 回避方法 |
|---|---|---|
| テキストがPPTXでずれる | `line-height` をfont-sizeと違う値にした | `line-height = font-size`（常に等値） · `height = font-size × 行数` |
| ヘッダー内テキストが上下ずれる | font-size / height を変更後にtopを再計算しなかった | **font-size・height・line-heightを変更した場合は必ずtopを再計算すること** · `top = slot.top + (slot.height − font-size) / 2` |
| 文字が切れる | widthが小さすぎる | `width = 推定テキスト幅 + font-size × 2` |
| フォントが反映されない | `font-family: sans-serif` のみ記述 | 必ず `'Noto Sans JP',sans-serif` と具体名を先頭に |
| 色が変換されない | rgba()や色名を使った | `#RRGGBB`形式のみ使用 |
| 要素が重ならない | Z順を考慮していない | 背景rect → 前景rect → textの順でDOM記述 |
| 角丸が出ない | `border-radius` を書いた | Converterは`border-radius`を実装しない。円形は`data-pptx-type="oval"`を使う |
| レイアウトが見つからない | layout名のタイポ | `01_title` / `02_structure_dark` / `03_structure_light` / `04_body` の4種のみ |
| PPTXサイズがおかしい | `Pt(1920)` を忘れた | `prs.slide_width = Pt(1920)` / `prs.slide_height = Pt(1080)` を必ず設定 |
