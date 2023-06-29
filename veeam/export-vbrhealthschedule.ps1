#Author: mas
#Date: 29.11.21

function Export-VbrHlthScheduleCsv {
    <#
    .Synopsis
    Export veeam the health check options to csv file.

    .Description
    Get the health check options of each backup and agent backup job. Export these to a csv file under the defined path.
    The csv uses the ';' sepperator. It can be edited and reimportet by Import-VeeamHealthScheduleCsv

    .Parameter Path
    The full path with filename, where the csv is gonna be created

    .Example
    # Export-VeeamHealthScheduleCsv -Path "C:\_Scripts\Veeam\Daily-Schedule\HealthCheckSchedule.csv"

    .Notes 
    Name: Marcel Schubert 
    Author: Marcel Schubert
    LastEdit: 30.11.2021
    #>
    [CmdletBinding()]
    param (
        [ValidateNotNullOrEmpty()]
        [string] $Path #Path to the exported csv
    )

    #Get all veeam backup and agent backup jobs
    $VBRJobs = Get-VBRJob -ErrorAction Stop | where-object {$_.JobType -eq "Backup" -or $_.JobType -eq "EpAgentBackup"}
    #Filter DSS (dorado snapshot) jobs
    $VBRJobs = $VBRJobs | Where-Object Name -notlike "*DSS"
    #new empty list
    $results = @()

    #iterate all jobs
    foreach($job in $VBRJobs){
        #fetch job options
        $Options = $job.getoptions()
        
        #parse recheck days from veeam type to string
        $recheckDays = ""
        $i=1 
        $count = ($Options.GenerationPolicy.RecheckDays | Measure-Object).Count
        foreach($day in $Options.GenerationPolicy.RecheckDays){
            $recheckDays += [string]$day #convert to string
            if($i -lt $count){
                $recheckDays += ","
            }
            $i++
        }

        #parse recheck months from veeam type to string
        $recheckMonths = ""
        $i=1
        $count = ($Options.GenerationPolicy.RecheckBackupMonthlyScheduleOptions.Months | Measure-Object).Count
        foreach($month in $Options.GenerationPolicy.RecheckBackupMonthlyScheduleOptions.Months){
            $recheckMonths += [string]$month #convert to string
            if($i -lt $count){
                $recheckMonths +=","
            }
            $i++
        }

        $session = $job.FindLastSession()
        if($null -eq $session){
            Throw "Couldnt fetch last session of job $($job.Name)"
        }

        #Build a custom object to easily convert to csv
        $results += [PSCustomObject]@{
            Name = $job.Name
            #Time = $job.ScheduleOptions.OptionsDaily.TimeLocal
            #RetentionPolicy = $job.BackupStorageOptions.RetainCycles
            #ActiveFull = $job.BackupStorageOptions.EnableFullBackup
            #SynFull = $job.BackupTargetOptions.TransformFulltoSyntethic
            #Genpolicycompactfullkind = $Options.GenerationPolicy.CompactFullBackupScheduleKind
            #Genpolicycompactfulldays = $Options.GenerationPolicy.CompactFullBackupMonthlyScheduleOptions.DayOfWeek
            #Genpolicycompactfullenable = $Options.GenerationPolicy.EnableCompactFull
            EnableRecheck = $Options.GenerationPolicy.EnableRechek
            #Weekly, Daily
            RecheckScheduleKind = $Options.GenerationPolicy.RecheckScheduleKind
            #Multiple: Monday-Sunday
            DailyScheduleOptionsRecheckDays = $recheckDays 
            #Monday-Sunday
            MonthlyScheduleOptionsDayOfWeek = $Options.GenerationPolicy.RecheckBackupMonthlyScheduleOptions.DayOfWeek
            #First, Second, Third, Fourth, Last
            MonthlyScheduleOptionsDayNumberInMonth = $Options.GenerationPolicy.RecheckBackupMonthlyScheduleOptions.DayNumberInMonth
            #MonthlyScheduleOptionsDayOfMonth = $Options.GenerationPolicy.RecheckBackupMonthlyScheduleOptions.DayOfMonth
            #Multiple: January-December
            MonthlyScheduleOptionsMonths = $recheckMonths
            #Additional information about curr backoup processing size
            LastSessTotProcessedSizeGB = [int]($session.Info.Progress.TotalUsedSize/ 1GB)
        }
    }

    #export the csv to the defined path
    $results | Export-Csv -Path $Path -Delimiter ";" -NoTypeInformation
}