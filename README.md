# 履歴書・職務経歴書ビルダー

就労移行支援の利用者向け。JIS準拠フォーマット(template.xlsx)への完全準拠Excel出力(.xlsx)と、Word出力(.docx)に対応。

## デプロイ手順
1. このフォルダをGitHubリポジトリにpush
2. Vercelで「New Project」→ リポジトリを選択 → そのままDeploy(Viteとして自動認識されます)

## ローカル開発
```
npm install
npm run dev
```

## 構成
- `public/template.xlsx` … 履歴書フォーマット(写真枠・印刷範囲設定済み)。差し替える場合はsrc/excelExport.jsのセル座標を合わせること
- `src/excelExport.js` … exceljsでテンプレートに書き込み(.xlsx)
- `src/wordExport.js` … docxライブラリで職務経歴書+履歴書を生成(.docx)
- `src/App.jsx` … 入力UI(空白期間の自動検出・テンプレ文サポート付き)

## データの扱い
入力データはサーバーに送信されず、ブラウザ内で完結します。保存はJSONエクスポートで行ってください。
