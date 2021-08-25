#code came from website
#https://adamtheautomator.com/powershell-email/
#slightly modified

Function Send-MailkitMessage {
  [CmdletBinding(
      SupportsShouldProcess = $true,
      ConfirmImpact = "Low"
  )] # Terminate CmdletBinding

  Param(
    [Parameter( Position = 0, Mandatory = $True )][String]$To,
    [Parameter( Position = 1, Mandatory = $True )][String]$Subject,
    [Parameter( Position = 2, Mandatory = $True )][String]$Body,
    [Parameter( Position = 3 )][Alias("ComputerName")][String]$SmtpServer = $PSEmailServer,
    [Parameter( Mandatory = $True )][String]$From,
    [String]$CC,
    [String]$BCC,
    [Switch]$BodyAsHtml,
    [switch]$Secure,
    $Credential,
    [Int32]$Port = 25
  )

  begin{
    
    Add-Type -Path "C:\Program Files\PackageManagement\NuGet\Packages\MailKit.2.15.0\lib\net47\MailKit.dll"
    Add-Type -Path "C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.2.15.0\lib\net47\MimeKit.dll" 
  }

  Process {
    $SMTP     = New-Object MailKit.Net.Smtp.SmtpClient
    $Message  = New-Object MimeKit.MimeMessage

    If ($BodyAsHtml) {
      $TextPart = [MimeKit.TextPart]::new("html")
    } Else {
      $TextPart = [MimeKit.TextPart]::new("plain")
    }
    
    $TextPart.Text = $Body

    $Message.From.Add($From)
    $Message.To.Add($To)
    
    If ($CC) {
      $Message.CC.Add($CC)
    }
    
    If ($BCC) {
      $Message.BCC.Add($BCC)
    }

    $Message.Subject = $Subject
    $Message.Body    = $TextPart

    if ($Secure) {$SMTP.Connect($SmtpServer, $Port,[MailKit.Security.SecureSocketOptions]::StartTls, $False)} else {$SMTP.Connect($SmtpServer, $Port, $False)}

    If ($Credential) {
      $SMTP.Authenticate($Credential.UserName, $Credential.GetNetworkCredential().Password)
    }

    If ($PSCmdlet.ShouldProcess('Send the mail message via MailKit.')) {
      $SMTP.Send($Message)
    }

    $SMTP.Disconnect($true)
    $SMTP.Dispose()
  }
}