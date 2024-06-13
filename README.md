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

### 実行

Windowsのコマンドプロンプトで

```
pwsh .\Send-NewBookNotice.ps1 <リストファイルCSV> <楽天API ID> <案内を受け取るメールアドレス>
```

