# Author: Marcel.schubert

enum LogTypes {
    Error = 0
    Warning = 1
    Info = 2
    Debug = 3
}
class Logger {
    [string] hidden $fullPath #full path of log file
    [LogTypes] $logLevel
    static [string] hidden $DATE_FORMAT = "dd/MM/yyyy HH:mm" #date format used in log
    static [string] hidden $DELIMITER_FORMAT = "`t"

    Logger([string]$p, [LogTypes]$ll, [bool]$verbose){
        if($verbose){$DebugPreference = 'Continue'}
        #Enable debug to see what is happening
        #before the logger is able to write
        Write-Debug 'called logger constructor'

        Write-Debug "logger setting log level: $ll, value: $($ll.value__)"
        $this.logLevel = $ll
        $this.fullPath = $p

        #Create log file if it does not exist
        if(!(Test-Path $this.fullPath -PathType Leaf)){
            Write-Debug 'logger creating new log file'
            
            New-Item -Path $this.fullPath  -ItemType "file"

            #Create the first entry in the log file which
            #then is used automaticaly by the next export-csv cmdlet calls
            $firstEntryVars = @{
                Path = $this.fullPath
                Delimiter = [Logger]::DELIMITER_FORMAT
                InputObject = ([PSCustomObject]@{
                                    Date =(Get-Date -Format ([Logger]::DATE_FORMAT))
                                    Type='[Info]'
                                    Message ='Create log file'  
                                }) 
                Force = $true
                NoTypeInformation = $true
            }

            Export-Csv @firstEntryVars

        } else {
            Write-Debug 'logger log file already exists. Exiting logger constructor'
        }
    }

    [void] Write([string]$message, [LogTypes]$logType, [bool]$verbose){
        if($verbose){$DebugPreference = 'Continue'}
        #Check if the current log level is equal or
        #above the supplied log level if not
        #do nothing and return
        if($this.logLevel.value__ -lt $logType.value__){
            return
        }

        #Check empty parms
        if($message -eq ''){
            Write-Error 'no log written, empty string provided'
            return
        }

        #Check if log file still exists
        if(!(Test-Path -Path $this.fullPath)){
            Write-Error 'no log written, can not access log file: moved, deleted, network connectivity or permissions changed'
            throw "log file is inaccessable"
        }

        $type = '[Type]'
        switch($logType){
            Info { $type = '[Info]'}
            Error{ $type = '[Error]'}
            Debug {$type = '[Debug]'}
            Warning { $type = '[Warning]'}
        }

        $output = [PSCustomObject]@{
            Date = (Get-Date -Format ([Logger]::DATE_FORMAT))
            Type = $type
            Message = $message  
        }

        $exportVars = @{
            Path = $this.fullPath
            Delimiter = [Logger]::DELIMITER_FORMAT
            Append = $true
        } 

        Export-Csv @exportVars -InputObject $output
        $cliOut = $output.Type + [Logger]::DELIMITER_FORMAT + [Logger]::DELIMITER_FORMAT + $output.Message
        Write-Debug  $cliOut
    }
}

Function Get-TLogger{
    <#
        .Synopsis
            Create a new logger object.
            Constructor tries to create log file and creates first entries.
        .Parameter Path
            [String] Full path of the log file.
        .Parameter Logtype
            [LogTypes] Sets the log level. Can be changed. Default is INFO
        .Outputs
            [Logger]
        .Notes 
            Name: Marcel Schubert 
            LastEdit: 30.11.2021
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [string] $Path,
        [LogTypes] $LogType = [LogTypes]::Info
    )
    $verbose = $PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent
    return [Logger]::New($Path, $LogType, $verbose)
}

Function Write-TLogger {
    <#
        .Synopsis
            Logs message using supplied logger instance
        .Parameter Logger
            [Logger]
        .Parameter Message
            [string]
        .Outputs LogType
            [LogTypes] The log level of the log entry
        .Notes 
            Name: Marcel Schubert 
            LastEdit: 30.11.2021
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$True)]
        [Logger] $Logger,
        [Parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [string] $Message,
        [LogTypes] $LogType = [LogTypes]::Info
    )
    $verbose = $PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent
    $Logger.Write($Message, $LogType, $verbose)
}
