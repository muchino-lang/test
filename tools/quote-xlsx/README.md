# 見積xlsx 出力ツール

Academy 保証つきPCの見積（卸価格・エンド売価・コミッション）を、
毎回Excelにコピペせずにブラウザで入力して **そのまま送れる .xlsx** として書き出すツール。

## 使い方

`index.html` をブラウザで開くだけ（ダブルクリックでOK。サーバー不要・オフライン可）。

1. **明細** … 機種名・卸価格（税込）・エンド売価（税込）を入力（行は何行でも追加可）
   - Excelやスプレッドシートから3列コピーして「まとめて貼り付け」に貼れば一括で入ります（タブ/カンマ区切り）
   - 税抜・コミッションは入力しながら画面上でも自動計算されます
2. **出力** … ファイル名を決めて「xlsxをダウンロード」

出力されるファイルは `template.xlsx` と同じ書式（フォント・網掛け・罫線・列幅・行高）で、
C・E・F・G列は **数式のまま** 入ります。受け取った側が B・D 列を直せば再計算されます。

| 列 | 内容 |
|---|---|
| A | 保証（機種名） |
| B | 卸価格（税込） ← 入力 |
| C | 卸価格（税抜） `=ROUNDUP(B/1.1,-1)` |
| D | エンド売価（税込） ← 入力 |
| E | エンド売価（税抜） `=ROUNDUP(D/1.1,-1)` |
| F | コミッション（税込） `=D-B` |
| G | コミッション（税抜） `=ROUNDUP(F/1.1,-1)` |

入力内容はブラウザの localStorage に残るので、閉じても次回そのまま続きから使えます。

## 書式を変えたいとき

1. `template.xlsx` を Excel で開いて、フォント・色・列幅などを編集して保存
2. `python3 tools/quote-xlsx/build.py` を実行（`index.html` が再生成されます）

`build.py` は template.xlsx から styles / theme / workbook などの書式パーツを取り出して
`src/index.src.html` に埋め込みます。**`index.html` は生成物なので直接編集しないこと**
（編集するのは `src/index.src.html` と `template.xlsx`）。

## 構成

```
tools/quote-xlsx/
├── index.html          ← 生成物。これを開いて使う（単体で完結・外部ライブラリなし）
├── src/index.src.html  ← 画面とxlsx生成ロジックの元ファイル
├── template.xlsx       ← 書式の元になるExcel
└── build.py            ← template.xlsx の書式を src に埋め込んで index.html を出力
```
