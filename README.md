# IsogawaRyo.github.io

個人紹介ページのソースコードです。コンテンツはすべて `content.xlsx` で管理し、変換スクリプトで `content.json` を生成してから `index.html` が読み込みます。

## ワークフロー

1. `content.xlsx` を Excel で開き、必要な情報（自己紹介・基本情報・連絡先など）を編集します。  
   - 列は `Section / Key / Value / Link` の4つです。  
   - 既存のセクションを利用するか、新規セクションを追加して構いません。
2. 変更を保存したら、以下のスクリプトで JSON を再生成します。（`content.json` は自動生成ファイルなので手動編集しません）

```bash
scripts/build_content.py
```

3. GitHub Pages などで配信される `index.html` は `content.json` を fetch して画面に反映します。JSON を更新し忘れると変更が反映されないので注意してください。

## ローカルプレビュー

リポジトリ直下で簡易サーバーを立ち上げ、ブラウザでアクセスしてください。

```bash
python -m http.server 8000
# ブラウザで http://localhost:8000 を開く
```
