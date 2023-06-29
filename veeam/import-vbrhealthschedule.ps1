function Import-VbrHlthScheduleCsv {
    <#
    .Synopsis
    Import veeam the health check options from exported csv file.

    .Description
    Gets the health check options of each backup and agent backup job, exported through Export-VeeamHealthScheduleCsv.
    Imports these from a csv file with the defined path. Each value is validated. Invalid inputs will skip job settings from beeing pushed to veeam.
    The csv uses the ';' sepperator. It can be edited and reimportet by Import-VeeamHealthScheduleCsv

    .Parameter Path
    The full path with filename, where the csv is gonna be created

    .Example
    # Import-VeeamHealthScheduleCsv -Path "C:\_Scripts\Veeam\Daily-Schedule\HealthCheckSchedule.csv"
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

    $importedJobs = Import-Csv -Path $Path -Delimiter ";" -ErrorAction Stop

    #TODO Import check if empty? more granual warnings

    foreach($job in $importedJobs){

        #if($job.EnableRecheck -eq $false){
        #    continue
        #}

        #Check days input. Replace spaces and remove duplicates. Literal content is checked later.
        $days = $job.DailyScheduleOptionsRecheckDays -replace '\s','' -split "," | Sort-Object  -Unique
        #double check amount of inputs
        if(($days | Measure-Object).Count -gt 7){
            Write-Host "More than 7 days specified. Skipping" $job.Name
            continue
        }
        #Check months input. Replace spaces and remove duplicates. Literal content is checked later.
        $months = $job.MonthlyScheduleOptionsMonths -replace '\s','' -split "," | Sort-Object  -Unique
        #Double check amount of inputs
        if(($months | Measure-Object).Count -gt 12){
            Write-Warning "More than 12 months specified. Skipping $($job.Name)"
            continue
        }

        #Check amount of enable reckeck inputs
        if(($job.EnableRecheck | Measure-Object).Count -gt 1){
            Write-Warning "More than 1 EnableRecheck specified. Skipping $($job.Name)"
            continue
        }

        #Check amount of RecheckScheduleKind inputs
        if(($job.RecheckScheduleKind | Measure-Object).Count -gt 1){
            Write-Warning "More than 1 RecheckScheduleKind specified. Skipping $($job.Name)" 
            continue
        }

        #Check amount of MonthlyScheduleOptionsDayOfWeek inputs
        if(($job.MonthlyScheduleOptionsDayOfWeek | Measure-Object).Count -gt 1){
            Write-Warning "More than 1 MonthlyScheduleOptionsDayOfWeek specified. Skipping $($job.Name)"
            continue
        }
        
        #Check amount of MonthlyScheduleOptionsDayNumberInMonth inputs
        if(($job.MonthlyScheduleOptionsDayNumberInMonth | Measure-Object).Count -gt 1){
            Write-Warning "More than 1 MonthlyScheduleOptionsDayNumberInMonth specified. Skipping $($job.Name)"
            continue
        }
        ## Check literal input
        #If during check $skip is changed to $true, at the end
        #of the validation the job entry is skipped from beeing edited.
        $skip = $false

        #Check each day entry and cast the correct veeam type
        foreach($day in $days){
            switch($day){
                "Monday"{ $day = [System.DayOfWeek]::Monday}
                "Tuesday"{ $day = [System.DayOfWeek]::Tuesday}
                "Wednesday"{ $day = [System.DayOfWeek]::Wednesday}
                "Thursday"{ $day = [System.DayOfWeek]::Thursday}
                "Friday"{ $day = [System.DayOfWeek]::Friday}
                "Saturday"{ $day = [System.DayOfWeek]::Saturday}
                "Sunday"{ $day = [System.DayOfWeek]::Sunday}
                Default {
                    Write-Host "Error changing" $job.Name "day data is wrong:" $day
                    $skip = $true
                    break
                }
            }
        }
        
        #Check each month entry and cast the correct veeam type
        foreach($month in $months){
            switch($month){
                "January"{$month = [Veeam.Backup.Common.EMonth]::January}
                "February"{$month = [Veeam.Backup.Common.EMonth]::February}
                "March"{$month = [Veeam.Backup.Common.EMonth]::March}
                "April"{$month = [Veeam.Backup.Common.EMonth]::April}
                "May"{$month = [Veeam.Backup.Common.EMonth]::May}
                "June"{$month = [Veeam.Backup.Common.EMonth]::June}
                "July"{$month = [Veeam.Backup.Common.EMonth]::July}
                "August"{$month = [Veeam.Backup.Common.EMonth]::August}
                "September"{$month = [Veeam.Backup.Common.EMonth]::September}
                "October"{$month = [Veeam.Backup.Common.EMonth]::October}
                "November"{$month = [Veeam.Backup.Common.EMonth]::November}
                "December"{$month = [Veeam.Backup.Common.EMonth]::December}
                Default {
                    Write-Host "Error changing" $job.Name "month data is wrong:" $month
                    $skip = $true
                    break
                }
            }
        }

        #Check if schedule kind is correct
        switch($job.RecheckScheduleKind){
            "Monthly"{}
            "Daily"{}
            Default {
                Write-Host "Skipping" $job.Name "Schedule kind is wrong"
                $skip = $true
            }
        }

        #Check each day input. Can only be one value and has not to be casted.
        switch($job.MonthlyScheduleOptionsDayOfWeek){
            "Monday"{}
            "Tuesday"{}
            "Wednesday"{}
            "Thursday"{}
            "Friday"{}
            "Saturday"{}
            "Sunday"{}
            Default {
                Write-Host "Skipping" $job.Name "MonthlyScheduleOptionsDayOfWeek is wrong:" $job.MonthlyScheduleOptionsDayOfWeek
                $skip = $true
            }
        }

        #Check schedule day number in month input. Can only be one value.
        switch($job.MonthlyScheduleOptionsDayNumberInMonth){
            "First"{}
            "Second"{}
            "Third"{}
            "Fourth"{}
            "Last"{}
            Default {
                Write-Host "Skipping" $job.Name "MonthlyScheduleOptionsDayNumberInMonth is wrong:" $job.MonthlyScheduleOptionsDayNumberInMonth
                $skip = $true
            }
        }  

        #Skip if error was found
        if($skip -eq $true){
            Write-Warning "Wrong input data. Skipping $($job.Name)" 
            continue
        }

        #Get the current job options and compare them to the csv input.
        $vbrJob = Get-VBRJob -Name $job.Name
        if($null -eq $vbrJob){
            Throw "No job could be fetched"
        }
        $compareOptions = $vbrJob.GetOptions() #alias current options

        if($null -eq $compareOptions){
            Throw "Error compareOptions from fetched job is null"
        }

        $Options = (Get-VBRJob -Name $job.Name).GetOptions() # Cannot just copy object. It will be a reference, TODO: more error checking?
        $Options.GenerationPolicy.EnableRechek = $job.EnableRecheck
        $Options.GenerationPolicy.RecheckScheduleKind = $job.RecheckScheduleKind
        $Options.GenerationPolicy.RecheckBackupMonthlyScheduleOptions.DayOfWeek = $job.MonthlyScheduleOptionsDayOfWeek
        $Options.GenerationPolicy.RecheckBackupMonthlyScheduleOptions.DayNumberInMonth = $job.MonthlyScheduleOptionsDayNumberInMonth
        $Options.GenerationPolicy.RecheckBackupMonthlyScheduleOptions.Months = $months
        $Options.GenerationPolicy.RecheckDays = $days 

        Write-Host "Checking for changes in job" $job.Name
        #Compare the options and only if they differ, push them to veeam
        $changes = $false

        if($compareOptions.GenerationPolicy.EnableRechek -ne $Options.GenerationPolicy.EnableRechek){
            Write-Host "Changes in EnableRechek"
            Write-Host "Old:" $compareOptions.GenerationPolicy.EnableRechek
            Write-Warning $Options.GenerationPolicy.EnableRechek
            $changes = $true
        }
        if($compareOptions.GenerationPolicy.RecheckScheduleKind -ne $Options.GenerationPolicy.RecheckScheduleKind){
            Write-Host "Changes in RecheckScheduleKind"
            Write-Host "Old:" $compareOptions.GenerationPolicy.RecheckScheduleKind
            Write-Warning $Options.GenerationPolicy.RecheckScheduleKind
            $changes = $true
        }
        if($compareOptions.GenerationPolicy.RecheckBackupMonthlyScheduleOptions.DayOfWeek -ne $Options.GenerationPolicy.RecheckBackupMonthlyScheduleOptions.DayOfWeek){
            Write-Host "Changes in RecheckBackupMonthlyScheduleOptions"
            Write-Host "Old:" $compareOptions.GenerationPolicy.RecheckBackupMonthlyScheduleOptions.DayOfWeek
            Write-Warning $Options.GenerationPolicy.RecheckBackupMonthlyScheduleOptions.DayOfWeek
            $changes = $true
        }

        if($compareOptions.GenerationPolicy.RecheckBackupMonthlyScheduleOptions.DayNumberInMonth -ne $Options.GenerationPolicy.RecheckBackupMonthlyScheduleOptions.DayNumberInMonth){
            Write-Host "Changes in DayNumberInMonth"
            Write-Host "Old:" $compareOptions.GenerationPolicy.RecheckBackupMonthlyScheduleOptions.DayNumberInMonth 
            Write-Warning $Options.GenerationPolicy.RecheckBackupMonthlyScheduleOptions.DayNumberInMonth
            $changes = $true
        }
        Write-Host "Old / New months count"
        Write-Host ($Options.GenerationPolicy.RecheckBackupMonthlyScheduleOptions.Months | Measure-Object).Count
        Write-Host ($compareOptions.GenerationPolicy.RecheckBackupMonthlyScheduleOptions.Months | Measure-Object).Count
        if(($Options.GenerationPolicy.RecheckBackupMonthlyScheduleOptions.Months | Measure-Object).Count -ne ($compareOptions.GenerationPolicy.RecheckBackupMonthlyScheduleOptions.Months | Measure-Object).Count){
            Write-Host "Changes in Months"
            Write-Host "Old:" $compareOptions.GenerationPolicy.RecheckBackupMonthlyScheduleOptions.Months 
            Write-Host "New:" $Options.GenerationPolicy.RecheckBackupMonthlyScheduleOptions.Months
            $changes = $true
        }

        Write-Host "Old / New recheck days count"
        Write-Host ($compareOptions.GenerationPolicy.RecheckDays | Measure-Object).Count
        Write-Host ($Options.GenerationPolicy.RecheckDays | Measure-Object).Count
        if(($compareOptions.GenerationPolicy.RecheckDays | Measure-Object).Count -ne ($Options.GenerationPolicy.RecheckDays | Measure-Object).Count){
            Write-Host "Changes in RecheckDays"
            Write-Host "Old:" $compareOptions.GenerationPolicy.RecheckDays
            Write-Host "New:" $Options.GenerationPolicy.RecheckDays
            $changes = $true
        }

        if($changes){
            Write-Host "-----------------------------------------------------------------"
            $decision = Get-UIChoicePrompt -Title "Veeam - Edit health check options" -Content "Do you really want to edit health check options for job: $($job.Name)"
             #TODO confirmation
            if(!$decision){
                Set-VBRJobOptions -Job $vbrJob -Options $Options -Verbose
            } else {
                Write-Host "Did not accept. Skipping jobs"
            }
        } else{
            Write-Host "No changes in job" $job.Name
        }
    }
    Write-Host "Finished"
}
