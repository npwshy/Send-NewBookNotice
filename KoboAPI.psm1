#
# Rakuten Kobo API
#

class GenreInfo {
    [string] $GenreId;
    [string] $GenreName;

    GenreInfo($h) {
        $this.GenreId = $h.koboGenreId
        $this.GenreName = $h.koboGenreName
    }
}

class KoboAPI {
    [GenreInfo[]] $GenreInfo;
    [DateTime] $LastAPICall = 0;
    [int] $APICallInternval = 1000; # milliseconds
    [string]$AppId;
    $Encoding = [System.Text.Encoding]::GetEncoding('UTF-8');
    [string]$NGWord = "分冊 話売り 単話版 連載版 巻セット 限定無料";
    [string]$NGWordParam = "NGKeywords=" + [System.Web.HttpUtility]::UrlEncode($this.NGWord, $this.Encoding);

    [string] $CacheDir = ".\cache";
    [DateTime] $Expire = [DateTime]::Now.AddHours(-3);
    [DateTime] $GenreDataExpire = [DateTime]::Now.AddDays(-14);
    $HashFunc = (New-Object System.Security.Cryptography.SHA256CryptoServiceProvider);

    LoadGenre() {
        $fp = "genre.xml"
        if (-not (Test-Path $fp) -or (Get-Item $fp).LastWriteTime -lt $this.GenreDataExpire) {
            $j = $this.GetGenreData()
            $j.children.child |Export-Clixml $fp
        }
        $this.GenreInfo = Import-Clixml $fp |ForEach-Object { [GenreInfo]::New($_) }
    }

    [hashtable] GetGenreData() {
        $res = Invoke-WebRequest "https://app.rakuten.co.jp/services/api/Kobo/GenreSearch/20131010?applicationId=$($this.AppId)&koboGenreId=101"
        return $res| ConvertFrom-Json -Depth 10 -AsHashtable
    }

    [string] GetGenreIdParam([string]$genre) {
        if (-not $genre) {
            return ""
        } elseif ($genre -match '^\d') {
            return  "koboGenreId=$genre"
        } else {
            $g = $this.GenreInfo |Where-Object { $_.GenreName -match $genre } |Select-Object -First 1
            return $g ? "koboGenreId=$($g.GenreId)" : ""
        }
    }

        Init([string]$id) {
            $this.AppId = $id

            $this.LoadGenre()
        }

    [Object] SearchKobo([string]$title, [string]$genre, [string]$keywords) {
        $url = @("https://app.rakuten.co.jp/services/api/Kobo/EbookSearch/20170426?applicationId=$($this.AppId)")
        $url += ,"title=$([System.Web.HttpUtility]::UrlEncode($title, $this.Encoding))"
        $url += ,$this.GetGenreIdParam($genre)
        if ($keywords) {
            $url += ,"keyword=$([System.Web.HttpUtility]::UrlEncode($keywords, $this.Encoding))"
        }
        $url += ,"formatVersion=2"
        $url += ,"elements=title,seriesName,author,salesDate,itemUrl,koboGenreId"
        $url += ,"field=1"

        $res = $this.APIGet($url -join('&'))
        if ($res.Contains('Items')) {
#            $r = $res.Items |Sort-Object @{Expression={$_.salesDate};Descending=1} |Where-Object { $_.seriesName -match $title }
            $r = $res.Items |Sort-Object @{Expression={$_.salesDate};Descending=1} |Where-Object { $_.title -match "^$title" }
            if ($r) {
                return $r |Select-Object -First 1
            } else {
                #write-host "$($res.Items.Count)件が検索にヒットしましたシリーズタイトルが一致するものはありませんでした"
                #$res.Items |%{ write-host "$($_.title) [$($t_.seriesName)]"}
                return $null
            }
        } else {
            #Write-Host "検索結果なし: $title in $genre, $keywords"
            return $null
        }
    }


    [hashtable] APIGet($url) {
        $fp = $this.GetCacheFilename($url)
        if (Test-Path $fp) {
            if ((Get-Item $fp).LastWriteTime -gt $this.Expire) {
                #--- キャッシュファイルがあり、新しい
                return (Get-Content $fp) |ConvertFrom-Json -AsHashtable
            }
        }

        #--- キャッシュがない、あるいは古い
        $this.Throttling()
        $res = Invoke-WebRequest $url -Method Get
        $this.LastAPICall = [DateTime]::Now

        if ($res.StatusCode -eq 200) {
            $res.Content |Out-File $fp -Encoding utf8
            return $res.Content |ConvertFrom-Json -AsHashtable
        }

        return $null
    }

    [string] GetCacheFilename($url) {
        $u = [URI]$url
        $hashcode = ($this.HashFunc.ComputeHash([Text.Encoding]::UTF8.GetBytes($url)) |ForEach-Object { $_.ToString("x2") }) -join('')
        return Join-Path $this.CacheDir ($u.host + "_" + $hashcode)
    }

    Throttling() {
        $waitms = $this.APICallInternval - (([DateTime]::Now - $this.LastAPICall).TotalMilliseconds)
        if ($waitms -gt 0) {
            Start-Sleep -Milliseconds $waitms
        }
    }
}