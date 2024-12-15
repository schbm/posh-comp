function Get-ExchangeMailBoxSizeInfo {
    <#
    .Synopsis
    Returns mailbox information of the supplied identity.
    .Description
    Returns mailbox information of the supplied identity. 
    IssueWarningQuota: Warning trigger
    ProhibitSendQuota: Send limit
    ProhibitSendReceiveQuota: Receive limit
    MaxReceiveSize:
    MaxSendSize:
    UseDatabaseQuotaDefaults:
    ItemCount:
    TotalDeletedSize:
    TotalSize:

    .Parameter Identity
    [string] The identity name of the mailbox e.g 'marcel.schubert'.
    .Parameter Json
    [bool] Returns the result json formatted if set to true.
    .Example
    #Show-ExchangeMailBoxSize -Identity 'marcel.schubert'

    .Notes 
    Name: Marcel Schubert 
    Author: Marcel Schubert
    LastEdit: 30.11.2021
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false)]
        [ValidateNotNullOrEmpty()]
        [switch]$Json = $false
    )

    process {
            Write-Verbose "Setting global error action to stop for compliance with the Get-Mailbox cmdlet"
            $OldPreferences = $global:ErrorActionPreference
            $global:ErrorActionPreference = 'Stop'
        try {
            Write-Verbose "Getting mailbox and mailbox statistics with name $Identity"
            $Mailbox = Get-Mailbox -Identity $Identity
            $MailboxStats = $Mailbox | Get-MailboxStatistics

            $Output = @{
                IssueWarningQuota = $Mailbox.IssueWarningQuota
                ProhibitSendQuota = $Mailbox.ProhibitSendQuota
                ProhibitSendReceiveQuota = $Mailbox.ProhibitSendReceiveQuota
                MaxReceiveSize = $Mailbox.MaxReceiveSize
                MaxSendSize = $Mailbox.MaxSendSize
                UseDatabaseQuotaDefaults = $Mailbox.UseDatabaseQuotaDefaults
                ItemCount = $MailboxStats.ItemCount
                TotalDeletedSize = $MailboxStats.TotalDeletedItemSize
                TotalSize = $MailboxStats.TotalItemSize
            }

            if($Json) {
                $Output = ConvertTo-Json $Output
            }
            
            Return $Output
        } finally {
            $Mailbox = $null
            $MailboxStats = $null
            $Output = $null 
            Write-Verbose "Setting back global error preference action"
            $global:ErrorActionPreference = $OldPreferences 
        }
    }
}


function Set-ExchangeMailboxSize {
    <#
    .Synopsis
    Sets a new mailbox quota in MB.
    .Description
    Sets a new mailbox quota in MB. Converts the value to real bytes.
    ProhibitSendReceiveQuota is always unlimited as defined as standard.
    .Parameter Identity
    [string] The identity name of the mailbox e.g 'marcel.schubert'.

    .Parameter NewMailboxSizeMB
    [double] The new size of the mailbox in MB

    .Example
    #Set-ExchangeMailboxSize -Identity marcel.schubert -NewMailboxSizeMB 900

    .Notes 
    Name: Marcel Schubert 
    Author: Marcel Schubert
    LastEdit: 30.11.2021
    #>
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()]
            [string]$Identity,    
            [Parameter(Mandatory=$false,ValueFromPipeline=$false)]
            [ValidateNotNullOrEmpty()]
            [double]$NewMailboxSizeMB = 600
        )
    process {
            Write-Verbose "Setting global error action to stop for compliance with the Get-Mailbox cmdlet"
            $OldPreferences = $global:ErrorActionPreference
            $global:ErrorActionPreference = 'Stop'
            try {
                if($NewMailboxSizeMB -gt 9000 -or $NewMailboxSizeMB -lt 300) {
                    throw "Mailbox size cannot exeed 2400MB or be less than 300MB"
                }
   
                Write-Verbose "Converting from MB to Bytes"
                $NewMailboxSizeBytes = Convert-TMBToByte $NewMailboxSizeMB
                $NewMailboxSizeWarningBytes = Convert-TMBToByte ($NewMailboxSizeMB-60)
                Write-Verbose "New Mailbox size bytes: $NewMailboxSizeBytes"
                Write-Verbose "New Mailbox size warning bytes: $NewMailboxSizeWarningBytes"

                $MailboxSizeConfig = @{
                    Identity = $Identity
                    IssueWarningQuota = $NewMailboxSizeWarningBytes
                    ProhibitSendQuota = $NewMailboxSizeBytes
                    ProhibitSendReceiveQuota = 'unlimited'
                    UseDatabaseQuotaDefaults = $false
                    ErrorAction = 'Stop'
                }
                 
                Write-Verbose "Changing Mailboxsize"
                $CurrQuota = (Get-Mailbox -Identity $MailboxSizeConfig.Identity).ProhibitSendQuota
                write-host "Setting from $CurrQuota to $NewMailboxSizeBytes Bytes"
                Set-Mailbox @MailboxSizeConfig
            } finally {
                Write-Verbose "Setting back global error preference action"
                $global:ErrorActionPreference = $OldPreferences
            }
    }
}

Function Reset-ExchangeMailboxSize {
    <#
    .Synopsis
    Sets the quota settings back to the default unlimited values
    .Description
    Sets the quota settings back to the default unlimited values
    IssueWarningQuota = 'unlimited'
    ProhibitSendQuota = 'unlimited'
    ProhibitSendReceiveQuota = 'unlimited'
    UseDatabaseQuotaDefaults = $true
    ErrorAction = 'Stop'

    .Parameter Identity
    [string] The mailbox identity to reset. e.g 'marcel.schubert'
    .Example
    # Reset-MailboxSizeExchange -Identity marcel.schubert

    .Notes 
    Name: Marcel Schubert 
    Author: Marcel Schubert
    LastEdit: 30.11.2021
    #>
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()]
            [string]$Identity
        )
    process {
            try {
                Write-Verbose "Setting global error action to stop for compliance with the Get-Mailbox cmdlet"
                $OldPreferences = $global:ErrorActionPreference
                $global:ErrorActionPreference = 'Stop'

                $MailboxSizeConfig = @{
                    Identity = $Identity
                    IssueWarningQuota = 'unlimited'
                    ProhibitSendQuota = 'unlimited'
                    ProhibitSendReceiveQuota = 'unlimited'
                    UseDatabaseQuotaDefaults = $true
                    ErrorAction = 'Stop'
                }
                Write-Verbose "Changing Mailboxsize"
                Set-Mailbox @MailboxSizeConfig
            } finally {
                Write-Verbose "Setting back global error preference action"
                $global:ErrorActionPreference = $OldPreferences
            }
    }
}

Function Enable-ExchangeMailboxSendAs {
    <#
    .Synopsis
    Adds SendAs permissions for the Identity "ToEnable" to the root identity.
    Only for on-premises exchange.
    .Description
    Uses Add-ADPermission cmdlet to add 'SendAs' permissions to the active directory
    maulbox / user object

    .Parameter RootIdentity
    [string] The mailbox identity to reset. e.g 'marcel.schubert'

    .Parameter IdentityToEnable
    [string] The mailbox identity to reset. e.g 'michael.hug'
    .Example
    # Enable-SendAsExchange RootIdentity marcel.schubert -IdentityToEnable michael.hug

    .Notes 
    Name: Marcel Schubert 
    Author: Marcel Schubert
    LastEdit: 30.11.2021
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$RootIdentity,
        [Parameter(Mandatory=$true,ValueFromPipeline=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$IdentityToEnable
    )

    process {
        try {
            Write-Verbose "Setting global error action to stop for compliance with the Get-Mailbox cmdlet"
            $OldPreferences = $global:ErrorActionPreference
            $global:ErrorActionPreference = 'Stop'
            
            Write-Verbose "Getting root identity"
            $rootMailbox = Get-Mailbox -Identity $RootIdentity
        
            $rootMailbox | Add-ADPermission -User $IdentityToEnable -ExtendedRights "Send As"
        } finally {
            Write-Verbose "Setting back global error preference action"
            $global:ErrorActionPreference = $OldPreferences
        }
    }
}
