# 履歴書・職務経歴書ビルダー

就労移行支援の利用者向け。exceljsによるJIS準拠レイアウトの履歴書Excel出力(.xlsx)と、職務経歴書のWord出力(.docx)に対応。

## デプロイ手順
1. このフォルダをGitHubリポジトリにpush
2. Vercelで「New Project」→ リポジトリを選択 → そのままDeploy(Viteとして自動認識されます)

## ローカル開発
```
npm install
npm run dev
```

## 構成
- `src/excelExport.js` … exceljsで履歴書をゼロから構築(.xlsx)。レイアウト変更はこのファイルを編集
- `src/wordExport.js` … docxライブラリで職務経歴書を生成(.docx)
- `src/App.jsx` … 入力UI(空白期間の自動検出・テンプレ文サポート付き)

## データの扱い
入力データはサーバーに送信されず、ブラウザ内で完結します。保存はJSONエクスポートで行ってください。
