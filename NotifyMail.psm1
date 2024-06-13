#
# Notify by email
#
using module .\NotifyBase.psm1
using module .\OutlookApp.psm1

class NotifyMailParams : NotifyParams {
    [string]$To;
    [string]$Cc;
    [string]$Subject;
}

class NotifyMail : NotifyBase {
    NotifyMail($p) : base($p) {}

    Send() {
        $msg = $this.GetHTMLMessage() -join("`n")
        $mailer = [OutlookApp]::New()
        $mailer.SendMail($this.Params.To, $this.Params.Cc, $this.Params.Subject,  $msg)
        $mailer.Quit()
    }
}
