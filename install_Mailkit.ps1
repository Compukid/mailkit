#package manager will not find mailkit in windows powershell 5, so far only works on version 7
if ($PSVersionTable.PSVersion.Major -ge 7){
    if ( (get-package -name mailkit -ErrorAction SilentlyContinue) -eq $null) {

        write-host "mailkit not installed" 

        #Confirm nuget repository is added
        if ( (Get-PackageSource -providername NuGet -ErrorAction SilentlyContinue) -eq $null) {  
            #Set Tls security and add nuget repository
            [net.servicepointmanager]::SecurityProtocol = [net.securityprotocoltype]::Tls12
            Register-PackageSource -Name MyNuGet -Location https://www.nuget.org/api/v2 -ProviderName NuGet
        }

        #install mimekit (depenecy) and mailkit manually. Need the SkipDependencies switch to bypass an error. By installing both apps individually it will meet requirments.
        Install-Package mimekit -ProviderName NuGet -SkipDependencies
        Install-Package mailkit -ProviderName NuGet -SkipDependencies

    }
}else{
    Write-Error "This script must be run on powershell 7 or greater!" 
}