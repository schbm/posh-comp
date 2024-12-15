						                                  #
# Purpose:	Backup of important Skype for Business Server 2015    #
#       Front End Data which is needed for a successfull restore  #
#		Including:					                              #  	
#			- CsConfiguraton Export			                      #
#			- CsLisConfiguration Export		                      #
#			- CsRgsConfiguration Export		                      #
#			- IIS-Metadata Configuration Export	                  #
#			- Lync 2013 CsUserData						          #
#			- Export Voice Configuration                          #
#           - Export Persistent Chat Data                         #
#			- Keep last 7 Backupfiles (tbd)                       #


#Variable Definition
$BackupPath = "CHANGE"
$BackupFiles = 14

$ScriptName = $Myinvocation.InvocationName

#Initialize Eventlog for reporting and Debugging
#------------------------------------------------
$evt=new-object System.Diagnostics.EventLog("Lync Server")
$evt.Source="Skype for Business Server 2015 Backup Skript"
$infoevent=[System.Diagnostics.EventLogEntryType]::Information
$warnevent=[System.Diagnostics.EventLogEntryType]::Warning
$errorevent=[System.Diagnostics.EventLogEntryType]::Error


#Get Skpe for Business Server 2015 Pool FQDN and SQL Server Information
#----------------------------------------------------------------------
Import-Module SkypeForBusiness
$SysInfo = Get-WmiObject -Class Win32_ComputerSystem
$ComputerFqdn = “{0}.{1}” -f $sysinfo.Name, $sysinfo.Domain
$LyncPoolname = Get-CsComputer -Identity $ComputerFqdn | Select-Object -ExpandProperty Pool
$LyncUserDB = Get-CsService -Identity userserver:$lyncpoolname | Select-Object -ExpandProperty UserDatabase
$LyncSqlServer = Get-CsService -Identity $LyncUserDB | Select-Object -ExpandProperty PoolFqdn
$LyncSqlInstance = Get-CsService -Identity $LyncUserDB | Select-Object -ExpandProperty SqlInstanceName
$SQLServer = $LyncSqlServer + "\" + $LyncSqlInstance


#Test if Poolname is set correctly
#---------------------------------
$PoolName = Get-CsService -ApplicationServer | where {$_.PoolFqdn -eq $LyncPoolname} | ft PoolFqdn
if($PoolName -eq $NULL){
	$evt.WriteEntry("Beim Ausführen des Backup Scripts für Skype for Business Server 2015 ist ein Fehler aufgetreten!
	Der im Backup Script angegebene Skype for Business Server 2015 Poolname konnte nicht gefunden werden.
	Überprüfen Sie den Skype for Business Server 2015 Poolnamen im Backup Script" + " ""$ScriptName""" + "!",$errorevent,0)
	Exit
}

#Test if Backup Path exists, define Backup Folder and how many Files to keep ($total)
#------------------------------------------------------------------------------------
$Date = get-date -format "yyyMMdd_HHmmss" 
$Date = $Date -replace „/“, „-“
$Servername = gc env:computername
$DestFolder = $Servername + "_" + $date
$storage = new-item -path $BackupPath -name $DestFolder -type directory
$FinalBackupPath = ($BackupPath + "\" + $DestFolder + "\")

$ChkFile = Test-path $FinalBackupPath
if($ChkFile -eq $false){
	$evt.WriteEntry("Beim Ausführen des Backup Scripts für Skype for Business Server 2015 ist ein Fehler aufgetreten!
	Der im Backup Script angegebene Backup Pfad ist falsch oder existiert nicht!
	Überprüfen Sie den Backup Pfad im Backup Script" + " ""$ScriptName""" + "!",$errorevent,0)
	Exit
}
Else{
    if((ls $BackupPath).count -gt $BackupFiles){
        ls $BackupPath |sort-object -Property {$_.CreationTime} | Select-Object -first 1 | Remove-Item -force -recurse
    }
}

#Define a File for Important Backup Notes
$BackupNote = ($FinalBackupPath + "IMPORTANT_BACKUP_NOTE.log")

#Perform Skype for Business Server 2015 CsConfiguration Backkup
#--------------------------------------------------------------
Export-CsConfiguration -Filename ($FinalBackupPath + "CsConfiguration.zip")

#Test if Backup File was written
$ChkFile = Test-path ($FinalBackupPath + "CsConfiguration.zip")
if($ChkFile -eq $false){
	$evt.WriteEntry("Beim Ausführen des Backup Scripts für Skype for Business Server 2015 ist ein Fehler aufgetreten!
	Der Export der Skype for Business Server 2015 Konfiguration ist fehlgeschlagen. Das Backup File wurde nicht erstellt!",$errorevent,0)
	Exit
}
Else{
	$evt.WriteEntry("Der Export der Skype for Business Server 2015 Konfiguration wurde erfolgreich durchgeführt.",$infoevent,0)
}

#Perform Skype for Business Server 2015 CsLisConfiguration Backup
#-------------------------------------------
Export-CsLisConfiguration -Filename ($FinalBackupPath + "CsLisConfiguration.zip")

#Test if Backup File was written
$ChkFile = Test-path ($FinalBackupPath + "CsLisConfiguration.zip")
if($ChkFile -eq $false){
	$evt.WriteEntry("Beim Ausführen des Backup Scripts für Skype for Business Server 2015 ist ein Fehler aufgetreten!
	Der Export der LIS Konfiguration ist fehlgeschlagen. Das Backup File wurde nicht erstellt!",$errorevent,0)
	Exit
}
Else{
	$evt.WriteEntry("Der Export der LIS Konfiguration wurde erfolgreich durchgeführt.",$infoevent,0)
}

#Perform Skype for Business Server 2015 Persistent Chat Backup
#--------------------------------------------------------------
$CheckPChat = Get-CsService -PersistentChatServer
if($CheckPChat -ne $Null){
    $PChatSql = Get-CsService -PersistentChatDatabase
    foreach($PChatPool in $PChatSql){
        $PChatSqlFqdn = ""
        $PChatSqlFqdn = ($PChatPool.PoolFqdn + "\" + $PChatPool.SqlInstanceName)
        Export-CsPersistentChatData -DBInstance $PChatSqlFqdn -FileName ($FinalBackupPath + "CsPChatConfiguration.zip") -ErrorAction SilentlyContinue
    }
}

#Test if Backup File was written
$ChkFile = Test-path ($FinalBackupPath + "CsPChatConfiguration.zip")
if($ChkFile -eq $false){
	$evt.WriteEntry("Beim Ausführen des Backup Scripts für Skype for Business Server 2015 ist ein Fehler aufgetreten!
	Der Export der Persistent Chat Konfiguration ist fehlgeschlagen. Das Backup File wurde nicht erstellt!",$errorevent,0)
	Exit
}
Else{
	$evt.WriteEntry("Der Export der Persistent Chat Konfiguration wurde erfolgreich durchgeführt.",$infoevent,0)
}

#Perform Skype for Business Server 2015 Response Group Service Backup
#-----------------------------------------------
Export-CsRGSConfiguration -Source applicationserver:$LyncPoolname -Filename ($FinalBackupPath + "RGSConfiguration.zip")

#Test if Backup File was written
$ChkFile = Test-path ($FinalBackupPath + "RGSConfiguration.zip")
if($ChkFile -eq $false){
	$evt.WriteEntry("Beim Ausführen des Backup Scripts für Skype for Business Server 2015 ist ein Fehler aufgetreten!
	Der Export der RGS Konfiguration ist fehlgeschlagen. Das Backup File wurde nicht erstellt!",$errorevent,0)
	Exit
}
Else{
	$evt.WriteEntry("Der Export der RGS Konfiguration wurde erfolgreich durchgeführt.",$infoevent,0)
}


#Perform Skype for Business Server 2015 Export Voice Configuration
#--------------------------------------------
$VoiceTemp = new-item -path "$($env:SystemRoot)\temp" -name "VoiceXML" -type directory
$VoiceBackupPath = $VoiceTemp.Fullname + "\"
Get-CsDialPlan | Export-Clixml -Path ($VoiceBackupPath + "Dialplan.xml")
Get-CsVoicePolicy | Export-Clixml -Path ($VoiceBackupPath + "VoicePolicy.xml")
Get-CsPstnUsage | Export-Clixml -Path ($VoiceBackupPath + "PstnUsage.xml")
Get-CsVoiceRoute | Export-Clixml -Path ($VoiceBackupPath + "VoiceRoute.xml")
Get-CsTrunkConfiguration | Export-Clixml -Path ($VoiceBackupPath + "TrunkConfiguration.xml")

#Create Zip File and remove old xml Files
$ZipFileName = "$($env:SystemRoot)\temp\VoiceConfiguration.zip"

if (test-path $ZipFileName) { 
  	echo "Zip file already exists at $ZipFileName" 
  	return 
} 

set-content $ZipFileName ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18)) 
(dir $ZipFileName).IsReadOnly = $false 
$ZipFile = (new-object -com shell.application).NameSpace($ZipFileName)
$ZipFile.CopyHere($VoiceTemp.FullName)

Start-Sleep -s 2

Move-Item $ZipFileName $FinalBackupPath -force
Get-Item $VoiceTemp | remove-item -force -Recurse 

#Test if Backup File was written
$ChkFile = Test-path ($FinalBackupPath + "VoiceConfiguration.zip")
if($ChkFile -eq $false){
	$evt.WriteEntry("Beim Ausführen des Backup Scripts für Skype for Business Server 2015 ist ein Fehler aufgetreten!
	Der Export der Voice Konfiguration ist fehlgeschlagen. Das Backup File wurde nicht erstellt!",$errorevent,0)
	Exit
}
Else{
	$evt.WriteEntry("Der Export der Voice Konfiguration wurde erfolgreich durchgeführt.",$infoevent,0)
}


#Perform Skype for Business Server 2015 Export User DB
#-----------------------------------------------------
#Check if UCS is activ on Lync Server
$Ucs = (Get-CsUserServicesPolicy).ucsallowed
$Data = Get-CsUserServicesPolicy
#If UCS is active on Skype for Business Server 2015 write IMPORTANT NOTE into Backup Path
If($ucs -match "True"){
    $Titel = "***IMPORTANT NOTE Export CsUserData***"
    $Note = "UCS is activated on Skype for Business Server 2015! Export CsUserData might not available for all Users."
    Write $Titel,$Note,$Data,"" | Out-File $BackupNote -Append
}

#Export User Data
Export-CsUserData -PoolFqdn $LyncPoolname -FileName ($FinalBackupPath + "CsUserData.zip")

#Test if Export was successfull
$ChkFile = Test-path ($FinalBackupPath + "CsUserData.zip")
if($ChkFile -eq $false){
	$evt.WriteEntry("Beim Ausführen des Backup Scripts für Skype for Business Server 2015 ist ein Fehler aufgetreten!
	Der Export der Userdaten (CsUserData) ist fehlgeschlagen. Das Backup File wurde nicht erstellt!",$errorevent,0)
	Exit
}
Else{
	$evt.WriteEntry("Der Export der Userdaten (CsUserData) wurde erfolgreich durchgeführt.",$infoevent,0)
}

#Check for Skype for Business Kerberos Accounts
#----------------------------------------------
#Check if Kerberos Accounts are activ on Skype for Business Server
$Kerb = Get-CsKerberosAccountAssignment
#If Kerberos Accounts are active on Skype for Business Server write IMPORTANT NOTE into Backup Path
If($Kerb -ne $Null){
    $Titel = "***IMPORTANT NOTE Skype for Business Kerberos Accounts***"
    $Note = "Kerberos Accounts are activated on Skype for Business Server 2015!"
    Write $Titel,$Note,$Kerb,"" | Out-File $BackupNote -Append
}