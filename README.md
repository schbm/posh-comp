# posh-comp
Compilation of 'useful' posh helpers i wrote for work.

> There is a possibility that some functions need some change to get them working because they are more or less directly copied from work.

More may follow if i am not too lazy to edit them for public use.

## Windows
```powershell
Get-LeaseInfoFromMac
```
 This function searches for an ```IP``` lease on the DHCP servers based on the supplied MAC address. It returns a ```PSCustomObject``` with the following fields: ```Lease``` (the found lease object), ```ServerDns``` (DNS name of the DHCP server where the lease resides), and ```ServerIp``` (IP address of the DHCP server).

```powershell
Get-InstalledDotNetVersions
```
Returns installed .NET versions by registry.

```powershell
Get-UIChoicePrompt
```
Prompts a closed question with the ability to decide.

## Veeam (V11a)

```powershell
Export-VbrHlthScheduleCsv
```
Export veeam the health check options to csv file.
(Editing required, DSS and Copy jobs are excluded)

```powershell
Import-VbrHlthScheduleCsv
```
Import veeam the health check options from exported csv file (directly apply settings to the jops).

```powershell
Get-VbrBackupCount
```

## Util

```powershell
Foreach-ObjectFast
```
Improved version of the ```foreach-object``` cmdlet. Without the unnesecary overhead.

```powershell
Where-ObjectFast
```
Improved version of the ```where-object``` cmdlet. Without the unnesecary overhead.

## Skype for Business (15)
```New-SfbNumberReport``` Is a legacy script for generating HTML reports of the telephone number usage in a skype for business env given a 
respective number inventory.

## IO

```powershell
Get-Encoding
```
Get file encoding. (Legacy)
