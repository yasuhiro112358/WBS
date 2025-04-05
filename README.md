# WBS

## About development
- To code, use VBE (VB Editor) in Excel.
- To manage the code, exporting it to `src` folder.
- Basically not import back the code to VBAProject file (.xlsm).

## 文字コードについて
- Excel VBAのコードは、各ロケールに応じた文字コードでエクスポートされる
  例：日本語ロケールの場合、Shift-JISでエクスポートされる。

## 運用手順
### VBEからのエクスポート
- エクスポート
- 文字コードをUTF-8に変換
- Gitにpush

### VBEへのインポート（対応中）
- Gitからpull
- 文字コードをUTF-8から各開発環境（VBE）に合わせて変換
- インポート





