# 動作確認テスト

`index.html` の主要フローを Playwright で通す統合テストです。

```
npm i playwright
node smoke.mjs
```

Chromium が `/opt/pw-browsers/chromium` に無い環境では、`smoke.mjs` 冒頭の
`executablePath` を各自の Chromium/Chrome のパスに変更してください。

カバー範囲：スキル表の読み込み・重点3項目・業務内容の連続入力と件数表示・
候補の自動提案・貼り付け枠からの取り込み・モ・ゲ・ジョの生成と折りたたみ・
「制約なし」チップ・保存時のコピーと内容保持・再保存の重複防止・宣言の記録・
スキルマップのIDとピン止めとフィルタ・スキル表CSVの往復と初期化・
AIモーダルの2モード・履歴・評価面談の実践率・下書きの自動復元。
