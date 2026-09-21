# 社内ドキュメント（HTML 版）

Word で作成された社内ドキュメントを HTML 化したものです。

## 構成

| パス | 内容 |
| --- | --- |
| `docs/index.html` | ドキュメント一覧（ここから各ページへ） |
| `docs/sales-guideline.html` | 営業チームの働き方ガイドライン |
| `docs/remote-work-policy.html` | R∞PC事業部 在宅勤務に関する方針（素案） |
| `docs/solution-sales-schedule.html` | ソリューション営業部 勤務スケジュール運用ルール |
| `docs/assets/style.css` | 共通スタイル（レスポンシブ / ダークモード / 印刷対応） |
| `docs/assets/export.js` | Word / Excel 出力ボタン |
| `docs/_source/*.docx` | 元の Word ファイル（更新履歴の比較用） |
| `tools/export.py` | .docx / .xlsx の一括生成スクリプト |

`docs/index.html` をブラウザで開くだけで表示できます（サーバー不要）。

## Word / Excel 出力

### 1. ブラウザのボタンから（手軽な方法）

各ドキュメントのページ上部にあるボタンから出力できます。

- **Word で出力** … `.doc` 形式でダウンロード（Word でそのまま開けます）
- **Excel で出力** … `.xls` 形式でダウンロード（区分・項目・内容の3列）
- **印刷 / PDF** … ブラウザの印刷ダイアログから PDF 保存

追加のインストールは不要です。Excel で `.xls` を開くときに形式確認のダイアログが
出る場合がありますが、「はい」で開けます。

### 2. スクリプトから（.docx / .xlsx が必要な場合）

```bash
pip install python-docx openpyxl
python3 tools/export.py                      # 全ドキュメントを dist/ に出力
python3 tools/export.py docs/sales-guideline.html   # 1件だけ出力
```

`dist/` に `.docx` と `.xlsx` が生成されます（`dist/` は Git 管理対象外）。

## 更新のしかた

内容を変更するときは `docs/*.html` を直接編集してください。
Word 側で更新した場合は、新しい `.docx` を渡してもらえれば HTML に反映します。
