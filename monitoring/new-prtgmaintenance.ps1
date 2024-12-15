# Author: Marcel.schubert
# Logger from module MAS-Tools
# Achtung Spaghetti

#TODO: Offload logger to ext. module
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

Function Get-Logger{
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

Function Write-Logger {
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

#################################################################################
#Glob vars
##########

$account = "<service user upn>"                                  #Ivanti service account
$prtgCredFile = "<path to cred file>"                            #Credentials for the svcivanti user #encr rsa
$prtgHost = "<prtg api fqdn>"                                    #fqdn of the prtg api server
$logPath = "<log path>"                                          #Log location
$scheduledTaskPath = "<scheduled task uri>"                      #Logical location of the scheduled tasks, created by ivanti
$startTimeScheduledTaskPath = "\Ivanti\Security Controls\Scans"  #Workaround for Get-ScheduledTaskInfo as it does not work with double backslash
$isVerbose = $false                                              #LEGACY: For console output, legacy requirement for logger. Enable console output with $true
$waitTime = 2                                                    #Sleep time in seconds between REST requests
$maintenanceDuration = 2                                         #The maintenance duration in hours
$rtLoggerLevel = "INFO"                                          #Log level of the logger see enum definition
 

$returnCode = 0
##########

#################################################################################
#Legacy: Edit accepted tls methods
##########
$AllProtocols = [System.Net.SecurityProtocolType]'Tls11,Tls12'
[System.Net.ServicePointManager]::SecurityProtocol = $AllProtocols


#################################################################################
#Legacy: Simple get request for fetching passhash etc.
##########
Function Get-WebRequest {
    Param (
        $WebRequestURL
    )
        try {
            $webRequest = [System.Net.WebRequest]::Create($WebRequestURL)
            $webRequest.ContentType = "application/xml"
            $webRequest.Method = "GET"
        
            $response = $webRequest.GetResponse()
            $stream = New-Object System.IO.StreamReader($response.GetResponseStream())
            $data = $stream.ReadToEnd()
        } catch {
            Write-Error "Error get request" #TBD proper error handling
        }
        return $data
}


#################################################################################
#Runtime logger setup
##########
$rtLogger = Get-Logger -Path $logPath -Logtype $rtLoggerLevel -Verbose
Write-Logger -Logger $rtLogger -Message "Logger init" -Verbose:$isVerbose
Write-Logger -Logger $rtLogger -Message "Starting Maintenance Script" -Verbose:$isVerbose
Write-Logger -Logger $rtLogger -Message "Test" -LogType Warning -Verbose:$isVerbose

#################################################################################
#Import ivanti module
##########
Import-Module STProtect –PassThru

#################################################################################
#Initial auth setup
##########
#TODO: Password decrypt; unnötig
[Byte[]] $key = (1..32)
$prtgPassword = Get-Content $prtgCredFile | ConvertTo-SecureString -Key $key
$strPRTGCred = New-Object -TypeName System.Management.Automation.PSCredential -Argumentlist $account, $PRTGPassword
$decryptedPW = $strPRTGCred.GetNetworkCredential().password


#################################################################################################################################################################################################################
#
# Main Code
#
#################################################################################################################################################################################################################

##########
#Authenticate user in ivanti. Get passhash needed for interacting with prtg api.
##########
$passHashURL = "https://$prtgHost/api/getpasshash.htm?username=$account&password=$decryptedPW"
Write-Logger -Logger $rtLogger -Message "Getting passhash with: https://prtg.ari-ag.mgt/api/getpasshash.htm?username=$account&password=" -Verbose:$isVerbose
$passwordHash = Get-WebRequest -WebRequestURL $PassHashURL
Write-Logger -Logger $rtLogger -Message "$passwordHash" -Verbose:$isVerbose

##########
#Get PRTG Devices max. uINT! 
##########
$deviceSourceURL = "https://prtg.ari-ag.mgt/api/table.xml?content=devices&output=csvtable&columns=device,host,objid&username=$account&passhash=$passwordHash&count=65535"

# use cache?
#if(Test-Path $prtgDeviceData) {
#    Remove-Item $prtgDeviceData
#}

Write-Logger -Logger $rtLogger -Message ("Getting device data with: https://prtg.ari-ag.mgt/api/table.xml?"+
"content=devices&output=csvtable&columns=device,host,objid&username=$account&passhash=&count=65535") -Verbose:$isVerbose
$deviceDataResult = (Get-WebRequest -WebRequestURL $deviceSourceURL) | ConvertFrom-CSV # use as cache? | Set-Content -Path $Output -Force
Write-Logger -Logger $rtLogger -Message "Response: Got $(($deviceDataResult | Measure-Object).Count) devices" -Verbose:$isVerbose

##########
#Get scheduled tasks, exits early if no tasks are found
##########
Write-Logger -Logger $rtLogger -Message "Getting ivanti windows scheduled patch tasks:" -Verbose:$isVerbose
$scheduledTasks = Get-ScheduledTask | Where {$_ -match "$scheduledtaskPath" -and ($_.State -eq "Ready" -and $_.Description -NotLike "*VMware Templates*")}
if($scheduledTasks -eq $null){
    Write-Logger -Logger $rtLogger -Message "Permission error, no scheduled tasks found" -LogType Error -Verbose:$isVerbose
    return 1
}
Write-Logger -Logger $rtlogger -Message "$($scheduledTasks | % {$_.Description+';'})" -LogType Debug -Verbose:$isVerbose
Write-Logger -Logger $rtlogger -Message "Found $(($scheduledTasks | Measure-Object).Count)" -Verbose:$isVerbose

##########
#Check if ivanti Get-MachineGroup returns error due to permission failure etc.
#Exits normally if no machine groups are found. 
#Exits with code 1 if an error occured.
##########
Write-Logger -Logger $rtLogger -Message "Checking ivanti get machine groups" -Verbose:$isVerbose  
try {
    $machineGroups = Get-MachineGroup
    if($machineGroups -eq $null){
        Write-Logger -Logger $rtLogger -Message "No machine groups returned" -LogType Error -Verbose:$isVerbose   
        return 0
    }
} catch {
    Write-Logger -Logger $rtLogger -Message "Could not fetch ivanti machine groups" -LogType Error -Verbose:$isVerbose
    Write-Logger -Logger $rtLogger -Message "$_" -LogType Error -Verbose:$isVerbose
    return 1
}
Write-Logger -Logger $rtLogger -Message "Ok" -LogType Info -Verbose:$isVerbose 

##########
#Main iteration.
#Iterate through each scheduled task get the name and
#find the corresponding ivanti machine group.
#Map the existing devices to the prtg device names and find out their prtg id.
#Set the maintenance time of the fetched id's.
##########
Write-Logger -Logger $rtLogger -Message "Iterating through tasks" -Verbose:$isVerbose

Foreach($scheduledTask in $scheduledTasks){
    $patchGroupName = $scheduledTask.Description #scheduled task description = ivanti machinegroup name
    Write-Logger -Logger $rtLogger -Message "$patchGroupName" -Verbose:$isVerbose
    Write-Logger -Logger $rtLogger -Message "Getting corresponding ivanti group" -Verbose:$isVerbose

    #Match name with the ivanti machine groups
    # EDIT THE DOMAIN SUFFIXES FIX!
    $patchGroupClients = (Get-MachineGroup -Name $patchGroupName).Filters.Name | ForEach-Object {$_ -replace "<DOMAINSUFFIX>", "" ` -replace "<DOMAINSUFFIX>", "" ` -replace "<DOMAINSUFFIX>", ""}
    Write-Logger -Logger $rtLogger -Message "Found clients:" -Verbose:$isVerbose
    Write-Logger -Logger $rtLogger -Message "$patchGroupClients" -Verbose:$isVerbose

    $startTime = ($scheduledTask | Get-ScheduledTaskInfo).NextRunTime

    #If the scheduled task has no planned next run time, skip the group, log the error and return 1 at the end of all the iterations.
    if($startTime -ne $null){
        
        $fmtStartTime = $startTime
        $fmtEndTime = $startTime.AddHours($maintenanceDuration)

        Write-Logger -Logger $rtLogger -Message "Next runtime is at: $fmtStartTime" -Verbose:$isVerbose
        Write-Logger -Logger $rtLogger -Message "Runs until: $fmtEndTime" -Verbose:$isVerbose

        ##########
        #GIterate through the clients and match with corresponding prtg id's
        #Add the device ids to the $setForDevices list. Which is then used to multi edit alls id's at once.
        ##########
        $setForDevices = @()
        Foreach($patchGroupClient in $patchGroupClients){

            $prtgDevices = $deviceDataResult | Where {$_.Device -eq $patchGroupClient}
            $prtgDevicesCount = ($prtgDevices | Measure-Object).Count

            if($prtgDevicesCount -gt 1){
                Write-Logger -Logger $rtLogger -Message "Multiple IDs for $patchGroupClient" -LogType Warning -Verbose:$isVerbose
                Write-Logger -Logger $rtLogger -Message "Found: $(($prtgDevices | Measure-Object).Count)" -LogType Warning -Debug -Verbose:$isVerbose
            } elseif ($prtgDevicesCount -eq 1){
                Write-Logger -Logger $rtLogger -Message "Single IDs for $patchGroupClient found" -LogType Debug -Verbose:$isVerbose
            }
        
            foreach($prtgDevice in $prtgDevices){
                if($prtgDevice.id -eq $null){
                    Write-Logger -Logger $rtLogger -Message "No ID found for $patchGroupClient $prtgDevice" -LogType Error -Verbose:$isVerbose
                    $returnCode = 1
                    continue
                }
                $setForDevices += $prtgDevice.id
            }   
        }
        
        if(($setForDevices | Measure-Object).Count -eq 0){
             Write-Logger -Logger $rtLogger -Message "No IDs found for $($scheduledTask.Description)" -LogType Error -Verbose:$isVerbose
             continue
        }
        
        #Reformat list. Add delimiter , at the end of each item and add them to string $idString
        $idString = "" #final string of multiple prtg id's separated by ,
        $iii = 1 #used for list reformatting
        $setForDevices | % {if($iii -ne ($setForDevices | Measure-Object).Count){$idString += $_ +",";}else{$idString += $_}; $iii++} 
        
        Write-Logger -Logger $rtLogger -Message "Action: Setting maintenance window Meta: $account, $passwordHash, $idString, $fmtStartTime, $fmtEndTime" -Verbose:$isVerbose

       $apiCall = ("https://prtg.ari-ag.mgt/editsettings?scheduledependency=0&want_maintenable=1&maintenable_=1&want_maintstart=1&maintstart_=$($fmtStartTime.ToString('yyyy-MM-dd-HH-mm-ss'))"+
        "&maintstart__picker=$($fmtStartTime | % {$_.ToString('yyyy-MM-dd+HH')})%3A$($fmtStartTime.Minute)"+
        "&want_maintend=1&maintend_=$($fmtEndTime.ToString('yyyy-MM-dd-HH-mm-ss'))"+
        "&maintend__picker=$($fmtEndTime | % {$_.ToString('yyyy-MM-dd+HH')})%3A$($fmtEndTime.Minute)"+
        "&id=$idString&domultiedit=1&username=$account&passhash=$passwordHash")


        Write-Logger -Logger $rtLogger -Message "Action: Multi-edit Meta: $apiCall" -Verbose:$isVerbose
        $retResponse = $null

        try{ #Try get; catch any errors
            $response = Invoke-WebRequest -Uri $apiCall -UseBasicParsing
            # This will only execute if the Invoke-WebRequest is successful.
            $statusCode = $response.StatusCode
            Write-Logger -Logger $rtLogger -Message "Http status code: $statusCode" -LogType Info -Verbose:$isVerbose
            Write-Logger -Logger $rtLogger -Message "Response: $response" -LogType Debug -Verbose:$isVerbose

            #Check response
            $responseWords = $response.Content.Split(" ")
            $changedObjCount = $responseWords[[array]::indexof($responseWords,"object(s)")-1]
            Write-Logger -Logger $rtLogger -Message "$changedObjCount objects were changed" -LogType Info -Verbose:$isVerbose

            $delta = ($setForDevices | Measure-Object).Count - $changedObjCount
            if($delta -ne 0){
                Write-Logger -Logger $rtLogger -Message "$delta objects could not be changed" -LogType Error -Verbose:$isVerbose
                $returnCode = 1
            } else {
                Write-Logger -Logger $rtLogger -Message "ok" -LogType Info -Verbose:$isVerbose
            }

        }catch{ #Set return value to error and write error to log
            $statusCode = $_.Exception.Response.StatusCode.value__
            Write-Logger -Logger $rtLogger -Message "HTTP error code: $statusCode" -LogType Error -Verbose:$isVerbose
            Write-Logger -Logger $rtLogger -Message "$_" -LogType Error -Verbose
            Write-Logger -Logger $rtLogger -Message "$($_.Exception.Response)" -LogType Error -Verbose:$isVerbose
            $returnCode = 1
        }

        #Read-Host -Prompt "waiting to continue" #TODO!

    } else {
        Write-Logger -Logger $rtLogger -Message "$($scheduledTask.TaskName) No start time defined" -LogType Error -Verbose:$isVerbose
        $returnCode = 1
    }
}
Write-Logger -Logger $rtLogger -Message "Finished, returning with opcode $returnCode" -Verbose:$isVerbose
$host.SetShouldExit($returnCode); exit