#original code came from website
#https://adamtheautomator.com/powershell-email/
#slightly modified

#this variable is just a place holder in case the dll's loaded in the begining do not match the version of dotnet installed. This variable will just show the latest dotnet version installed.
if ($iswindows -eq $true){ $latestdotnet = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP' -Recurse | Get-ItemProperty -Name version -ErrorAction SilentlyContinue | where-object {$_.pschildname -eq "Client" -or $_.pschildname -eq "Full"} | Sort-Object -Property version  | Select-Object version, pschildname -Last 1}

#may need a way to confirm linux dot net install???

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
    
    if ($PSVersionTable.PSVersion.Major -lt 7) {
        Write-Error "This function requires powershell version 7 or greater. Exiting function...."
        exit
    }

    #need to load dll for both mailkit and mimekit packages.
    #too much to write code to dynamically grab the correct .net dll so this has been hard coded
    #If you recieve a .net error please check the $latestdotnet and change below to correct version.
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