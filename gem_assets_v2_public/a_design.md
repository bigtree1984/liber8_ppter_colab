# liber8_ppter デザイン仕様書
# AIへの指示例：「a_design.mdの約束事を遵守してスライドHTMLを生成してください」

更新日: 2026-06-16

---

## このファイルについて

このファイルは **PowerPoint（.pptx）に変換可能なHTMLを生成するための
デザインシステム仕様書** です。

HTMLはブラウザで表示するだけでなく、python-pptxを使って
PowerPointファイルに変換することを前提としています。
そのためCSS の `flex` / `grid` / `rgba` / `border-radius` / グラデーションなど、
PPTXに変換できない表現は**使用禁止**です。
色指定は `#RRGGBB` 形式のみ、全要素は `position:absolute` で記述してください。

### 関連ドキュメント

| ファイル | 役割 | 読む人 |
|---|---|---|
| **a_design.md**（このファイル） | カラー・トークン値・禁止事項の辞書（What） | AI（Gemini Gem） |
| **a4h2p_spec.md** | HTML記述の技術仕様・変換ルール（How） | AI（Gemini Gem） |
| **a4h2p_sample.html** | 実際のスライドサンプル（座標の具体例） | AI（Gemini Gem） |
| **design_catalogue.dc.html** | スライドパターン一覧・設計ルール注記（参照用） | AI（Gemini Gem） / 人間（設計者） |

HTMLを生成する際は、このファイル（a_design.md）＋ a4h2p_spec.md ＋ a4h2p_sample.html ＋ design_catalogue.dc.html の
4ファイルを合わせて参照してください。

---

## 1. カラートークン

### 1-1. 共通カラー（A案・B案 共通・変更禁止）

| トークン名 | PPTXスロット | HEX | 用途 |
|---|---|---|---|
| text-primary | dk1 | #1A1A1A | 本文・最重要テキスト |
| text-secondary | dk2 | #767676 | サブテキスト・キャプション（WCAG AA） |
| bg-primary | lt1 | #FFFFFF | メイン背景・白余白 |
| bg-secondary | lt2 | #F5F5F5 | スライド背景・セクション区切り |
| brand-blue | accent1 | #0088C8 | ブランドカラー・メインアクセント（固定） |
| alert-red | accent2 | #CC0028 | アラート・警告・最重要強調 |

### 1-2. アクセントカラー

| トークン名 | PPTXスロット | HEX | グレー輝度Y | 用途 |
|---|---|---|---|---|
| yellow | accent3 | #F0C000 | 76% | ハイライト・エネルギー・注目喚起 |
| gold | accent4 | #C87800 | 57% | セカンダリアクセント・プレミアム感 |
| navy | accent5 | #004F8C | 32% | 見出し強調・ダークアンカー・A系背景 |
| sky-tint | accent6 | #D6EEF8 | 93% | 背景ティント・KPIカード背景 |

### 1-3. カラー使用ルール

- テキストボックス内のインライン装飾（spanでの色変え）は**対象外**
- テキストボックス単位でトークンを適用する
- **#RRGGBB形式のみ使用**（rgba・色名・短縮形 禁止）
- B系タイトル下ディバイダーは accent1（#0088C8）固定
- グレースケール印刷での識別：accent2(Y19%) accent5(Y32%) accent1(Y52%) accent4(Y57%) accent3(Y76%) accent6(Y93%)

---

## 2. タイポグラフィトークン

フォント: **Noto Sans JP** 固定
Line-height: **1（固定）**
Letter-spacing: **0（標準）**
PPTXのBold: weight 700 → True / weight 400 → False

| トークン | PPTX pt | HTML px | Weight | デフォルト色 | PPTXスロット | 用途 |
|---|---|---|---|---|---|---|
| type-display | 38pt | 101px | 700 | #1A1A1A | dk1 | 大見出し（表紙・扉） |
| type-title | 22pt | 59px | 700 | #1A1A1A | dk1 | スライドタイトル（B系） |
| type-heading | 20pt | 53px | 700 | #1A1A1A | dk1 | 見出し・強調 |
| type-body | 18pt | 48px | 400 | #1A1A1A | dk1 | 本文・箇条書き（統一） |
| type-body-sm | 16pt | 43px | 400 | #767676 | dk2 | 本文小・第2階層箇条書き |
| type-caption | 12pt | 32px | 400 | #767676 | dk2 | キャプション・注釈 |
| type-label | 10pt | 27px | 400 | #767676 | dk2 | ラベル・出典・著作権表示 |

### フォントサイズ換算式

```
PPTX pt × 8/3 ≈ HTML px
HTML px × 0.375 = PPTX pt
```

### テキストボックス高さ計算

```
height = font-size(px) × 行数
例: 1行 / 48px → height = 48px
例: 2行 / 48px → height = 96px
```

> 詳細・計算根拠 → `a4h2p_spec.md` Section 0-3（SSOT）

### テキストボックス幅計算

```
width = 推定テキスト幅 + font-size × 3
推定テキスト幅（Noto Sans JP 日本語）= 文字数 × font-size × 0.95
推定テキスト幅（Latin混在）= 日本語文字数 × font-size × 0.95 + Latin文字数 × font-size × 0.65
```

---

## 3. スペーシングトークン

8pxグリッド基準。単位は HTML px（PPTX px は÷2）。

| トークン | HTML px | PPTX px | 主な用途 |
|---|---|---|---|
| space-1 | 8px | 4px | アイコン・バッジ内パディング |
| space-2 | 16px | 8px | 隣接要素間・リスト item 間 |
| space-3 | 24px | 12px | eyebrow とタイトルの間 |
| space-4 | 32px | 16px | タイトルと本文の間・caption 上余白 |
| space-5 | 40px | 20px | 箇条書きブロック間・カード内パディング |
| space-6 | 48px | 24px | セクション区切り |
| space-8 | 64px | 32px | カラム間ギャップ・大ブロック間余白 |
| space-10 | 80px | 40px | 参考値 |
| space-12 | 96px | 48px | 大セクション間 |
| space-15 | 120px | 60px | A系スライド上下余白 |
| space-20 | 160px | 80px | 扉・QA・表紙の大余白 |

---

## 4. セーフゾーン定義

> ⚠ **SSOT = `a4h2p_spec.md` Section 1-B**。以下はデザイン参照用の抜粋。数値に差異がある場合は spec.md を優先すること。

キャンバス: **1920 × 1080 px**
PPTX実寸: **25.4 × 14.29 cm（10 × 5.63 in）**
換算: HTML px ÷ 2 = PPTX px / PPTX px ÷ 37.8 = cm

### B系スライド（04_body）実測値

| ゾーン | HTML px | PPTX cm | 備考 |
|---|---|---|---|
| 左右マージン | 38px | 0.50cm | 実測値 |
| タイトルゾーン高さ | 94px | 1.245cm | タイトル＋ページ番号 |
| タイトル下ディバイダー | 4px | — | accent1固定・ignore |
| コンテンツ開始Y | 98px | — | ディバイダー直下 |
| 著作権ゾーン高さ | 34px | 0.445cm | 最下部・ignore |
| コンテンツ幅 | 1844px | 24.40cm | 1920 - 38×2 |
| コンテンツ高さ | 948px | 12.57cm | 1080 - 98 - 34 |

### A系スライド（02_structure_dark / 03_structure_light）設計値

| ゾーン | HTML px | PPTX cm | 備考 |
|---|---|---|---|
| 左右マージン | 38px | 0.50cm | B系と共通 |
| 上マージン | 120px | 1.59cm | 設計値 |
| 著作権ゾーン高さ | 34px | 0.445cm | B系と共通・ignore |
| コンテンツ幅 | 1844px | 24.40cm | B系と共通 |
| コンテンツ高さ | 926px | 12.26cm | 1080 - 120 - 34 |

---

## 5. PPTXレイアウト定義

スライドマスター名: `liber8_ppter`

| レイアウト名 | 対象 | 背景 | 備考 |
|---|---|---|---|
| 01_title | 公式表紙 | ダーク | 会社規程。公式資料の表紙のみ |
| 02_structure_dark | ダーク背景A系＋非公式表紙 | ダーク | 区切り・QA・結び・社内勉強会表紙 |
| 03_structure_light | ライト背景A系 | ホワイト | 目次・資料共有 |
| 04_body | B系本文系全般 | ホワイト | テンプレート07〜18 |

PPTレイアウトが管理する要素（HTMLでは `data-pptx-type="ignore"` で仮描写）:
- 著作権フッター（全レイアウト共通）
- ページ番号（04_body）
- タイトル下ディバイダーライン（04_body）

---

---

## 6. レイアウト別カラー適用ルール

| レイアウト | 背景 | メインテキスト | サブテキスト | アクセント | 備考 |
|---|---|---|---|---|---|
| 01_title | 公式規程 | 公式規程 | 公式規程 | 公式規程 | 会社規程に従う |
| 02_structure_dark | accent5 | lt1(#FFFFFF) | lt1 60%透過 | accent3 | ダーク背景・白テキスト |
| 03_structure_light | lt1(#FFFFFF) | dk1(#1A1A1A) | dk2(#767676) | accent1 | 白背景・通常テキスト |
| 04_body | lt1(#FFFFFF) | dk1(#1A1A1A) | dk2(#767676) | accent1 | 白背景・タイトルはdk1 |

---

## 7. 分子別トークン早見表

### A系分子

| 分子 | 要素 | font-size | font-weight | color |
|---|---|---|---|---|
| M-A1 ダーク全面背景 | eyebrow | 27px | 400 | lt1(#FFFFFF) 45%透過 → #739EC0相当 |
| M-A1 ダーク全面背景 | type-display line1 | 101px | 700 | #FFFFFF |
| M-A1 ダーク全面背景 | type-display line2 | 101px | 700 | accent3 |
| M-A1 ダーク全面背景 | アクセントバー | rect h:6px | — | accent4 |
| M-A1 ダーク全面背景 | subtitle | 48px | 400 | #FFFFFF 75%透過 |
| M-A3 A系タイトルブロック | eyebrow | 27px(type-label相当) | 400 | dk2(#767676) |
| M-A3 A系タイトルブロック | 見出し | 101px(type-display) | 700 | accent5 |
| M-A3 A系タイトルブロック | 左アクセントバー | rect w:10px | — | accent1 |
| M-A4 セクション番号 | SECTION label | 27px | 400 | lt1 40%透過 |
| M-A4 セクション番号 | 番号（自由演技） | 自由 | 700 | accent3 |
| M-A4 セクション番号 | ルール | rect h:4px | — | accent2 |
| M-A4 セクション番号 | タイトル | 53px(type-heading相当) | 700 | #FFFFFF |

### B系分子

| 分子 | 要素 | font-size | font-weight | color |
|---|---|---|---|---|
| M-B0 ヘッドライン | 背景バンド | rect h:154px · top:108px | — | accent6(#D6EEF8) |
| M-B0 ヘッドライン | キーメッセージ（1行） | 48px(type-body) · top:161px | 700 | accent5(#004F8C) |
| M-B0 ヘッドライン | キーメッセージ（2行） | 48px(type-body) · top:137px | 700 | accent5(#004F8C) |
| M-B0 ヘッドライン | コンテンツ開始Y | BODY_TOP=272（バンド下端262+余白10px） | — | — |

| 分子 | 要素 | font-size | font-weight | color |
|---|---|---|---|---|
| M-B1 スライドヘッダー | type-title | 59px | 700 | dk1(#1A1A1A) |
| M-B1 スライドヘッダー | ページ番号 | 27px | 400 | dk2(#767676) · ignore |
| M-B1 スライドヘッダー | ディバイダー | rect h:4px | — | accent1(#0088C8) · ignore |
| M-B3 箇条書き（第1階層） | ドット | oval 20×20px | — | accent1(#0088C8) |
| M-B3 箇条書き（第1階層） | テキスト | 48px(type-body) | 400 | dk1(#1A1A1A) |
| M-B3 箇条書き（第2階層） | ドット | oval 14×14px | — | dk2(#767676) |
| M-B3 箇条書き（第2階層） | テキスト | 43px(type-body-sm) | 400 | dk2(#767676) |
| M-B4 データカード | カード背景 | rect | — | accent6 |
| M-B4 データカード | ラベル | 27px(type-label相当) | 400 | accent1 | **text-align:center、left=カード左端・width=カード幅**（フル幅テキストボックスで座標センタリング不要） |
| M-B4 データカード | 数値（通常） | 自由演技 | 700 | accent5 | **text-align:right**、right端=カード幅60%位置（left=カードleft+幅×0.6-width） |
| M-B4 データカード | 数値（警告） | 自由演技 | 700 | accent2(#CC0028) | **text-align:right**（同上） |
| M-B4 データカード | 単位 | 48px(type-body) | 400 | dk2(#767676) | **text-align:left**・left=カードleft+幅×0.6（60:40境目固定） |
| ~~M-B5 削除~~ | → L-2H等のLayoutに統合 | — | — | — |
| M-B6 タイムライン行 | 時刻 | 32px(type-caption相当) | 400 | dk2(#767676) |
| M-B6 タイムライン行 | 縦バー | rect w:3px | — | accent1 / accent3 |
| M-B6 タイムライン行 | 行タイトル | 48px(type-body) | 700 | dk1(#1A1A1A) |
| M-B6 タイムライン行 | サブテキスト | 43px(type-body-sm) | 400 | dk2(#767676) |
| M-B7 人物カード | カード背景 | rect | — | lt2(#F5F5F5) |
| M-B7 人物カード | アイコン背景 | rect | — | accent1 / accent4 |
| M-B7 人物カード | 名前 | 53px(type-heading) | 700 | dk1(#1A1A1A) |
| M-B7 人物カード | 部署 | 43px(type-body-sm) | 400 | dk2(#767676) |
| M-B7 人物カード | テーマ | 48px(type-body) | 700 | accent1(#0088C8) |
| M-B8 比較テーブル | ヘッダー行背景 | rect | — | dk1(#1A1A1A) |
| M-B8 比較テーブル | A案列ヘッダー | rect | — | accent1(#0088C8) |
| M-B8 比較テーブル | セルテキスト | 43px(type-body-sm) | 400 | dk1(#1A1A1A) |
| M-B8 比較テーブル | 交互行bg | rect | — | lt2(#F5F5F5) |
| M-B9 円形カード | カード（PPTX=oval） | oval 380〜500px | — | accent1/accent5/accent6 |
| M-B9 円形カード | テキスト | 48px(type-body) | 700 | lt1(#FFFFFF) または accent5 |
| M-B10 Descriptionカード | アクセントライン（variant B） | rect h:8px | — | accent1/accent2 |
| M-B10 Descriptionカード | タイトル | 48px(type-body) | 700 | dk1(#1A1A1A) |
| M-B10 Descriptionカード | description | 43px(type-body-sm) | 400 | dk2(#767676) |
| M-B10 Descriptionカード | シャドウ | data-pptx-shadow="preset1" | — | — |
| M-B11 ゾーン見出し | 左アクセントバー（variant A） | rect w:8px **h:85px**（= 53×1.6 テキスト高さに揃える） | — | accent1(#0088C8) |
| M-B11 ゾーン見出し | 見出しテキスト | 53px(type-heading) | 700 | dk1(#1A1A1A) |
| M-B11 ゾーン見出し | バンド背景（variant B） | rect h:全幅 | — | accent1(#0088C8) |
| M-B11 ゾーン見出し | バンドテキスト（variant B） | 53px(type-heading) | 700 | lt1(#FFFFFF) |
| M-B12 標準テーブル | ヘッダー行背景 | rect | — | dk1(#1A1A1A) |
| M-B12 標準テーブル | ヘッダーテキスト | 27px(type-label) | 700 | lt1(#FFFFFF) |
| M-B12 標準テーブル | 左端列背景 | rect | — | lt2(#F5F5F5) |
| M-B12 標準テーブル | 左端列テキスト | 43px(type-body-sm) | 700 | dk1(#1A1A1A) |
| M-B12 標準テーブル | データ行境界 | rect h:1px | — | #E8E8E8 |
| M-B12 標準テーブル | 強調セル | — | 700 | accent1(#0088C8) |

### B系Organism

| Organism | 要素 | サイズ | weight | color | 備考 |
|---|---|---|---|---|---|
| O-B9 グリッドフロー3×2 | カード背景 | rect 588×433px | — | lt1(#FFFFFF) | gap:40px |
| O-B9 グリッドフロー3×2 | アクセントバー | rect h:8px w:588px | — | accent1/accent4 |
| O-B9 グリッドフロー3×2 | カードタイトル | 48px(type-body) | 700 | dk1(#1A1A1A) |
| O-B9 グリッドフロー3×2 | カード本文 | **32px(type-caption)** | 400 | dk2(#767676) | カード幅が狭いためtype-captionを使用（スケール内対応） |

#### O-B5 横型フロー+グルーピング 配置ロジック

> 個別座標値・配置ロジックの詳細は **`design_catalogue.dc.html` O-B5セクション（SSOT）** を参照。

---

## 8. 矢印（At-4）寸法規則

> 矢印寸法の詳細・比率規則は **`a4h2p_spec.md` Section 3-5（SSOT）** を参照。
> 縦方向矢印（向上指示）は accent4(#C87800) を使用。

---

## 8-B. Tone of Voice（文章表現ルール）

すべてのテキストは以下のルールに従う：

| ルール | 内容 | 例 |
|---|---|---|
| 体言止め | タイトル・見出し・箇条書きは名詞で終える | ◯「議事録作成の自動化」 ✕「議事録作成を自動化する」 |
| 読点を打たない | 句読点「、」「。」は原則使わない | ◯「AI導入で工数を90%削減」 ✕「AI導入で、工数を90%削減した。」 |
| 簡潔・端的 | 1要素1メッセージ。冗長な修飾を避ける | ◯「欠品率30%改善」 ✕「欠品率が大幅に改善されました」 |
| 数値は具体的に | 効果・実績は数字で示す | ◯「月40時間削減」 ✕「大幅に削減」 |

**例外:**
- ヘッドラインは結論を示すため、述語を含む短文も可（読点なし・「。」は任意）
- 引用（B09）は発言をそのまま記載するため、このルールの対象外

---

## 9. 禁止事項

- `position:relative` / `static` / `fixed`（全要素 `position:absolute` 必須）
- `display:flex` / `display:grid` / `float`
- `right:` / `bottom:` 座標指定
- `rgba()` / `hsla()` / 色名 / `#RGB` 短縮形
- `gradient`（グラデーション全禁止・マテリアルデザイン準拠のため奥行きはシャドウで表現）
- **`border-radius` — 完全禁止（Converterは一切実装しない）**
  - PPTXのrect/textboxに角丸を付与するAPIは使用しない設計方針
  - 円形が必要な場合は `data-pptx-type="oval"` で代替
  - HTMLプレビュー用途でも記述しないこと（Converterの実装をシンプルに保つため）
- `font-family` に汎用ファミリー名のみ記載（`sans-serif` 等、Converterが無視するため）
- テキストボックス内インラインspan装飾
- `letter-spacing` の非ゼロ指定
- `text-align:center` / `text-align:right`（原則テキストボックス内は `text-align:left`。デザイン上の中央配置はtextboxの座標で表現する。**例外：KPIカードのKPI名ラベル・乗算式フロー（Molecule）の円内ラベル・ポジショニングマップの軸ラベルおよびプロット要素ラベル・比較テーブル（O-B1/O-B3）の列ヘッダー名および評価記号セル（◎△○）は `text-align:center` を許容する**）

### 推奨機能

- **ドロップシャドウ（T-4）**: `data-pptx-shadow="preset1"` + CSS `box-shadow:3px 3px 5px rgba(0,0,0,0.40)` を併記
  → PPTプリセット1（#000000・透明度60%・サイズ100%・ぼかし4pt・角度45°・距離3pt）に対応
  → 枠線（border）は原則使用しない。シャドウで境界を表現する（マテリアルデザイン準拠）

### 画像の扱い

画像は以下の2パターンを使い分ける。**`<img>` タグはWeb公開画像の埋め込み時のみ使用可**：

| パターン | タグ | type | 使うシーン |
|---|---|---|---|
| **placeholder** | `<div>` | `data-pptx-type="placeholder"` | AIが「ここに画像があるべき」と判断した場合。WebアプリでユーザーがアップロードまたはAI画像生成 |
| **image（URL）** | `<img>` | `data-pptx-type="image"` | Web上の公開画像URLを参照・埋め込む場合 |

```html
<!-- placeholder の例 -->
<div data-pptx-type="placeholder"
     data-placeholder-label="製品イメージ"
     data-placeholder-hint="スマートフォンを持つビジネスパーソン・明るいオフィス・水平構図"
     style="position:absolute; left:960px; top:98px; width:882px; height:700px;
            background:#F0F0F0;"></div>

<!-- image（URL）の例 -->
<img data-pptx-type="image"
     src="https://example.com/photo.jpg"
     style="position:absolute; left:960px; top:98px; width:882px; height:700px;">
```
