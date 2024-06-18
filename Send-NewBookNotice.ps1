#
# 続刊を探して案内を出す
#
using module .\KoboAPI.psm1
using module .\NotifyMail.psm1

Param(
[Parameter()][string]$BookList,
[Parameter()][string]$AppId,
[string]$MailTo,
[string]$Encoding = "oem"
)

Set-StrictMode -Version latest
$ErrorActionPreference = "stop"

#
# CSV ヘッダー定義
#
class BookListHeader {
    static $Title = ”タイトル";
    static $Genre = "ジャンル";
    static $Keyword = "キーワード";
    static $SalesDate = "最新刊発売日";
}


#
# CSV から詮索リストをロード
#
$csvdata = Import-Csv $script:BookList -Encoding $script:Encoding | ForEach-Object { @{
        Title = $_.[BookListHeader]::Title;
        Genre = $_.[BookListHeader]::Genre;
        Keyword = $_.[BookListHeader]::Keyword;
        SalesDate = [DateTime]($_.[BookListHeader]::SalesDate ? $_.[BookListHeader]::SalesDate : 0);
    }
}

#
# Kobo API で最新刊を検索、未知のもののみをピックアップ
#
$api = [KoboAPI]::New()
$api.Init($script:AppId)

$newbooklist = @()
foreach ($item in $csvdata) {
    if ($b = $api.SearchKobo($item.Title, $item.Genre, $item.Keyword)) {
        if ([DateTime]$b.salesDate -gt $item.SalesDate) {
            $newbooklist += ,$b
        }
    } else {
        write-host "検索しましたがみつかりませんでした: $($item.Title)"
    }
}

if ($newbooklist) {
    #
    # CSVデータの最新刊発売日を更新
    #
    $newbooklist | ForEach-Object {
        $newbook = $_;
        $item = $csvdata | Where-Object { $newbook.title -match "^$($_.Title)" }
        $item.SalesDate = $newbook.salesDate
    }

    #
    # 案内通知
    #
    $params = [NotifyMailParams]::New()
    $params.To = $script:MailTo
    $params.Subject = "楽天Kobo続刊のお知らせ"
    $params.NewBookList = $newbooklist
    $notifyer = [NotifyMail]::New($params)
    $notifyer.Send()

    #
    # CSV ファイルを更新
    #
    # まずはバックアップ
    $ext = Split-Path -Extension $BookList
    $basename = $BookList -replace "$ext$",''
    3 .. 0 |ForEach-Object {
        $n = $_;
        $n1 = $_ + 1;
        Move-Item -Path "$basename-$n$ext" -Destination "$basename-$n1$ext" -Force -ErrorAction SilentlyContinue
    }
    Move-Item -Path "$BookList" -Destination "$basename-0$ext" -Force -ErrorAction SilentlyContinue

    #
    # 保存
    $csvdata |ForEach-Object {
        @{
            [BookListHeader]::Title = $_.Title;
            [BookListHeader]::Genre = $_.Genre;
            [BookListHeader]::Keyword = $_.Keyword;
            [BookListHeader]::SalesDate = $_.SalesDate;
        }
    } |Select-Object ([BookListHeader]::Title),([BookListHeader]::Genre),([BookListHeader]::Keyword),([BookListHeader]::SalesDate) |
        Export-Csv -Path $BookList -Encoding $script:Encoding
}

exit
