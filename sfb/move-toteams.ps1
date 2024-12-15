function Get-LogsOlderThan{
    [cmdletbinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullorEmpty()]
        [string] $path,
        [int] $days = 30
    )
    if ($days -lt 1) {
        throw "you cannot rm files that are younger than 1"
    }

    if (-not (Test-Path -Path $path -PathType Container)){
        throw "path is not valid"
    }
    Write-Verbose "finding log files older than days=$days"
    
    $compareDate = (Get-Date).AddDays($days*-1)
    Write-Verbose "Comparison Date=$compareDate"

    
    $items = Get-ChildItem -Path "$path\*"
    Write-Verbose "Total items in path count=$($items.Count)"

    
    $filteredItems = $items | Where-Object {
        ($_.Extension -eq ".txt" -or $_.Extension -eq ".log") -and $_.CreationTime -lt $compareDate
    }
    Write-Verbose "Total filtered items in path count=$($filteredItems.Count)"

    return $filteredItems
}

# Start ########################################################################################################################################################

$LOG_DIR = "<logdir>"
$GROUP = "<ad group to migrate>"
$SVC = "<entra service user with permissions>"
$TARGET = "sipfed.online.lync.com"


Start-Transcript -OutputDirectory $LOG_DIR -NoClobber -IncludeInvocationHeader

try{

    $oldLogs = Get-LogsOlderThan -path $LOG_DIR -days 31 -Confirm:$false -ErrorAction stop 
    
    # $oldLogs | Remove-Item
    $oldLogs = $null

    $DC = (Get-ADDomainController -Discover -NextClosestSite -ErrorAction Stop).HostName[0]
    Write-Host "found dc=$dc"


    $CREDS = Get-StoredCredential -Target $SVC
    if ($creds -eq $null){
        throw "stored credentials not found"
    }


    If ((Get-Module MicrosoftTeams -ListAvailable) -eq $Null){
        throw "MicrosoftTeams module is not installed"
    } 


    If ((Get-Module SkypeForBusiness -ListAvailable) -eq $Null){
        throw "SkypeForBusiness module is not installed"
    } 

    Import-Module MicrosoftTeams -ErrorAction Stop
    Import-Module SkypeForBusiness -ErrorAction Stop

    Write-Host "connecting to ms teams"
    $connection = Connect-MicrosoftTeams -Credential $CREDS
    if ($connection -eq $null){
        throw "connect to ms teams failed"
    }

    # this creates a list of cs users which are in the migration group and have not been migrated
    $groupMembers = Get-ADGroupMember $GROUP -Server $DC -ErrorAction Stop |
        Where-Object { $_.ObjectClass -eq 'User' } |
        Get-ADUser -Properties UserPrincipalName -ErrorAction Stop | 
        Select-Object -ExpandProperty UserPrincipalName | 
        ForEach-Object {
            $upn = $_
            Get-CsUser -Identity $upn
        }
    $groupMembers = $groupMembers | Where-Object HostingProvider -ne $TARGET

    Write-Host "migrating $(($groupMembers).Count) users"

    $groupMembers | ForEach-Object {
        $member = $_
        Write-Host "migrating $($member.UserPrincipalName)"

        Move-CsUser -Identity $member.UserPrincipalName `
            -Target $TARGET `
            -DomainController $DC `
            -Credential $CREDS `
            -Confirm:$false `
            -Verbose
    }

} catch {
    Write-Host "An error occured that cannot be handled gracefully, script execution was aborted"
    Write-Host $_
} finally{
    Disconnect-MicrosoftTeams -Verbose
    Stop-Transcript
}


