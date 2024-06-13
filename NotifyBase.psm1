#
# Notify Base
#

class NotifyParams {
    $NewBookList;
}

class NotifyBase {
    [NotifyParams] $Params;

    NotifyBase($p) { $this.Params = $p }

    [string[]] GetHTMLMessage() {
        $msg = @()
        $msg += ,'<html>'
        $msg += ,'<body>'
        $msg += $this.GetTextMessage() |ForEach-Object { "$_<br/>" }
        $msg += ,'</body>'
        $msg += ,'</html>'
        return $msg
    }

    [string[]] GetTextMessage() {
        $msg = @()
        $msg += ,""
        $msg += ,"楽天Koboに続刊が出ました"
        $msg += ,""

        $this.Params.NewBookList |ForEach-Object {
            $msg += ,"タイトル: <a href=""$($_.itemUrl)"">$($_.title)</a>"
            $msg += ,"発売日: $($_.salesDate)"
            $msg += ,""
        }

        return $msg
    }
}