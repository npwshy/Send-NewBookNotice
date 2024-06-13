# Send-NewBookNotice
Send a notice of new book release (Kobo)
楽天 Kobo の API を使い、続刊があったら案内をだすスクリプト

## 使い方

### 準備

どこかにフォルダを作成し、コードをクローンまたはコピー。
cache サブフォルダを作成。

```
mkdir cache
```

sample-booklist.csv を参考に案内がほしい本のタイトル（先頭からの一部でOK）、ジャンル（省略可）を CSV に入れる。キーワード列は省略可。タイトルに入れた検索文字だけでは正しく最新刊が探せない時にキーワード列に詳細なキーワードを入れる。最新刊発売日は自動で更新されるので入力不要。

### 実行

Windowsのコマンドプロンプトで

```
pwsh .\Send-NewBookNotice.ps1 <リストファイルCSV> <楽天API ID> <案内を受け取るメールアドレス>
```

