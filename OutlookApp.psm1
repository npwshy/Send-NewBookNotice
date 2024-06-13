#
# Outlook Application interface
#

class OutlookApp {
    $App;
    $OutlookIsRunning;
    $LastSendItem = $null;

    OutlookApp() {
        $this.OutlookIsRunning = Get-Process -Name OUTLOOK -ErrorAction ignore
        $this.App = New-Object -ComObject Outlook.Application
    }

    SendMail([string]$to, [string]$cc, [string]$subject, [string]$msg) {
        $item = $this.App.CreateItem(0) # 0 = MailItem
        $item.To = $to
        if ($cc) {
            $item.Cc = $cc
        }
        $item.Subject = $subject
        $item.BodyFormat = 1 # 1 = HTML
        $item.HTMLBody = $msg
        $this.LastSendItem = @{To=$to; Subject=$subject; SentOn=[DateTime]::Now}
        $item.Send()
    }

    Quit() {
        $this.WaitSendingItem()
        if (-not $this.OutlookIsRunning) {
            $this.App.Quit()
            Get-Process -Name OUTLOOK -ErrorAction ignore |Stop-Process
        }
    }

    WaitSendingItem() {
        if (-not $this.LastSendItem) { return }

        $sentItemsFolder = $this.App.GetNamespace("MAPI").GetDefaultFolder(5)  # Send Items

        $cond = "[SentOn] >= '$($this.LastSendItem.SentOn.ToString('yyyy\/M\/d HH:mm'))' And [Subject] = '$($this.LastSendItem.Subject)' And [To] = '$($this.LastSendItem.To)'"
        while (1) {
            $items = $sentItemsFolder.Items.Restrict($cond)
            if ($items.Count -gt 0) {
                break
            }
            start-sleep -s 2
        }
    }
}