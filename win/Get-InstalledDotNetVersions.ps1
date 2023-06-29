#Author: mas
#Date: 29.11.21

function Get-InstalledDotNetVersions {
    <#  
    .SYNOPSIS  
        Returns installed .NET versions by registry.
    
    .DESCRIPTION
        This function retrieves the installed .NET versions by querying the Windows registry.
    
    .OUTPUTS
        [System.Object]
        Returns an object containing information about the installed .NET versions.
    
    .NOTES 
        Author: Marcel Schubert
        LastEdit: 30.11.2021
        URL: https://stackoverflow.com/questions/3487265/powershell-script-to-return-versions-of-net-framework-on-a-machine 
    #>

    $Lookup = @{
        378389 = [version]'4.5'
        378675 = [version]'4.5.1'
        378758 = [version]'4.5.1'
        379893 = [version]'4.5.2'
        393295 = [version]'4.6'
        393297 = [version]'4.6'
        394254 = [version]'4.6.1'
        394271 = [version]'4.6.1'
        394802 = [version]'4.6.2'
        394806 = [version]'4.6.2'
        460798 = [version]'4.7'
        460805 = [version]'4.7'
        461308 = [version]'4.7.1'
        461310 = [version]'4.7.1'
        461808 = [version]'4.7.2'
        461814 = [version]'4.7.2'
        528040 = [version]'4.8'
        528049 = [version]'4.8'
    }

    $netFrameworkPath = 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP'

    $installedVersions = Get-ChildItem $netFrameworkPath -Recurse |
        Get-ItemProperty -name Version, Release -EA 0 |
        Where-Object { $_.PSChildName -match '^(?!S)\p{L}'} |
        Select-Object -Property @{
            Name = ".NET Framework"
            Expression = {$_.PSChildName}
        }, 
        @{
            Name = "Product"
            Expression = {$Lookup[$_.Release]}
        }, 
        Version, Release

    return $installedVersions
}
