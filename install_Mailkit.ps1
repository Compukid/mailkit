
if ( (get-package -name mailkit) -eq $null) {

    write-host "mailkit not installed" 

    #Confirm nuget repository is added
    if ( (Get-PackageSource -providername NuGet) -eq $null) {  
        #Set Tls security and add nuget repository
        [net.servicepointmanager]::SecurityProtocol = [net.securityprotocoltype]::Tls12
        Register-PackageSource -Name MyNuGet -Location https://www.nuget.org/api/v2 -ProviderName NuGet
    }

    #install mimekit (depenecy) and mailkit. Need the SkipDependencies switch to bypass an error. By installing both apps it will meet requirments
    Install-Package mimekit -ProviderName NuGet -SkipDependencies
    Install-Package mailkit -ProviderName NuGet -SkipDependencies

}