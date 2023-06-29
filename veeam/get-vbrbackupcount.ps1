#Author: mas
#Date: 29.11.21

Function Get-VbrBackupCount {
    Param(
        [parameter(Mandatory=$true,ValueFromPipeline=$true)]
        $backup  
    )

    $storages = $backup.GetAllStorages()

    $output = @()

    foreach($storage in $storages){
        $outItem = [PSCustomObject]@{ 
                Date = $storage.CreationTime
                IsFullFast = $storage.IsFullFast
                IsIncrementalFast = $storage.IsIncrementalFast
                IsFull = $storage.IsFull
                FilePath = $storage.FilePath
                VM = $($storage.FilePath.ToString() -split '\.')[0]
                EffectiveSize = $storage.Stats.BackupSize / 1TB
            }

        $output += $outItem
    }

    return $output | Group-Object VM | Select Count, Name

}

Function Get-VbrAllBackupCount {
    $jobs = Get-VbrJob | ? TypeToString -EQ "VMware Backup Copy"

    $jobs | Foreach-Object {
        Write-Host "Checking $($_.Name)"
        Get-VbrBackup -Name $_.Name | Get-VbrBackupCount
    }
}
