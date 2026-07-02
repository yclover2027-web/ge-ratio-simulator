# ジェネリック医薬品 使用率シミュレーター

使用薬剤一覧CSVを読み込み、ジェネリック医薬品の使用率をシミュレーションするWebアプリです。

## ローカルで起動する

```bash
npm start
```

起動後、次のURLを開きます。

```text
http://127.0.0.1:8000/
```

Windowsでは `start_app.bat` をダブルクリックしても起動できます。

## Vercelで公開する

1. このフォルダをGitHubリポジトリに置きます。
2. VercelでそのリポジトリをImportします。
3. Framework Presetは `Other` のままで公開します。
4. Build CommandとOutput Directoryは空欄で問題ありません。

`/api/latest-mhlw-master` がVercelのサーバー側処理として動き、厚生労働省の最新版Excelを取得します。

## 注意

- APIキーやパスワードなどの機密情報はコードに直接書かないでください。
- 最新マスターデータの自動取得にはインターネット接続が必要です。
- 厚生労働省ページの構造やファイル名が変わった場合は、`api/latest-mhlw-master.js` と `server.mjs` の抽出処理を見直してください。
