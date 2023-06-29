#SOS this is legacy code
#SOS this is legacy code
#SOS this is legacy code

$Firma = "BLA"
$Date = Get-Date -format d
$File_Path = split-path $SCRIPT:MyInvocation.MyCommand.Path -parent
$File = $File_Path + "\number_report_$Firma" + ".htm" #target file
 
If (Test-Path $File){
    Remove-Item $File -Force
}
 
If (-not (Get-Module ActiveDirectory)){ 
    Import-Module ActiveDirectory -Force
}
 
$ADDomain = Get-ADDomain | Select-Object DistinguishedName
$SearchBase = "CN=Configuration," + $ADDomain.DistinguishedName
 
# import number blocks
[XML]$DID_Liste = Get-Content "$File_Path\list.xml"
 
$Style = @'
<style type=text/css>
 
body { 
    font-family: Lato, sans-serif;
}
 
ul.tab {
    list-style-type: none;
    margin: 0;
    padding: 0;
    overflow: hidden;
    border: 1px solid #ccc;
    background-color: #f1f1f1;
}
 
/* Float the list items side by side */
ul.tab li {
    float: left;
}
 
/* Style the links inside the list items */
ul.tab li a {
    display: inline-block;
    color: black;
    text-align: center;
    padding: 14px 16px;
    text-decoration: none;
    transition: 0.3s;
    font-size: 17px;
}
 
/* Change background color of links on hover */
ul.tab li a:hover {
    background-color: #ddd;
}
 
/* Create an active/current tablink class */
ul.tab li a:focus, .active {
    background-color: #ccc;
}
 
/* Style the tab content */
.tabcontent {
    display: none;
    padding: 6px 12px;
    -webkit-animation: fadeEffect 1s;
    animation: fadeEffect 1s;
}
 
@-webkit-keyframes fadeEffect {
    from {opacity: 0;}
    to {opacity: 1;}
}
 
@keyframes fadeEffect {
    from {opacity: 0;}
    to {opacity: 1;}
}
 
button.accordion {
    background-color: #eee;
    color: #444;
    cursor: pointer;
    padding: 18px;
    width: 100%;
    border: none;
    text-align: left;
    outline: none;
    font-size: 15px;
    transition: 0.4s;
}
 
button.accordion.active, button.accordion:hover {
    background-color: #ddd;
}
 
button.accordion:after { 
    content: '\02795'; 
    font-size: 13px; 
    color: #777; 
    float: right; 
    margin-left: 5px; 
}
button.accordion.active:after { 
    content: '\2796';
}
 
div.panel {
    padding: 0 18px;
    background-color: white;
    max-height: 0;
    overflow: hidden;
    transition: 0.6s ease-in-out;
    opacity: 0;
}
 
div.panel.show {
    opacity: 1;
    max-height: 100%;
    overflow-y: auto;
}
 
table { 
    border-collapse: collapse; 
}
 
th, td { 
    text-align: left;
    padding: 8px;
    border: 1px;
    solid: #000;
}
 
tr:nth-child(even){
    background-color: #f2f2f2;
}
 
th { 
    background-color: black;
    color: #ffffff;
}
 
table[id^=DDI] td:first-child {
    text-align: left;
    padding: 8px;
    cursor: pointer;
}
 
* {
  box-sizing: border-box;
}
 
#Suchen_Eingabe {
    background-image: url('http://www.w3schools.com/css/searchicon.png');
    background-position: 10px 12px;
    background-repeat: no-repeat;
    width: 100%;
    font-size: 16px;
    padding: 12px 20px 12px 40px;
    border: 1px solid #ddd;
    margin-bottom: 12px;
}
 
</style>
'@
 
$Style | Out-File $File -append
 
"<!DOCTYPE html>" | Out-File $File -append
"<html>" | Out-File $File -append
 
"<script src='https://ajax.googleapis.com/ajax/libs/jquery/2.1.3/jquery.min.js'></script>" | Out-File $File -append
"<script src='https://code.highcharts.com/highcharts.js'></script>" | Out-File $File -append
"<script src='https://code.highcharts.com/modules/data.js'></script>" | Out-File $File -append
 
Function CheckNumber($Line_URI){
    
    $Tel_Reservation = tel:++$Line_URI
    $Tel_Number = tel:++$Line_URI
    $Extension = $Line_URI.Substring($Line_URI.Length - 4)
    $Line_URI = tel:+$Line_URI #+ ";ext=" + $Extension
      
    # Active Directory Lookup
    $Result_Users = Get-ADObject -Filter {Objectclass -eq "User" -and msRTCSIP-Line -eq $Line_URI -and msRTCSIP-UserEnabled -eq $True } -Properties Displayname, telephoneNumber | Select-Object Displayname, telephoneNumber
    $Result_Common_Area_Phones = Get-ADObject -Filter {ObjectClass -eq "Contact" -and msRTCSIP-OwnerUrn -eq "urn:device:commonareaphone" -and msRTCSIP-Line -like $Tel_Number -and msRTCSIP-UserEnabled -eq $True } -Properties Displayname, telephoneNumber | Select-Object Displayname, telephoneNumber
    $Result_Analog_Devices = Get-ADObject -Filter {ObjectClass -eq "Contact" -and msRTCSIP-OwnerUrn -eq "urn:device:analogphone" -and msRTCSIP-Line -like $Tel_Number -and msRTCSIP-UserEnabled -eq $True } -Properties Displayname, telephoneNumber | Select-Object Displayname, telephoneNumber
    $Result_Response_Group_Services = Get-ADObject -SearchBase $SearchBase -Filter {ObjectClass -eq "Contact" -and msRTCSIP-OwnerUrn -eq "urn:application:RGS" -and msRTCSIP-Line -like $Tel_Number -and msRTCSIP-UserEnabled -eq $True } -Properties Displayname, telephoneNumber | Where-Object Displayname -notcontains "Announcement Service" | Where-Object Displayname -notcontains "RGS Presence Watcher" | Select-Object Name, Displayname, telephoneNumber
    $Result_DialIn_Conferencing_Access_Numbers = Get-ADObject -SearchBase $SearchBase -Filter {ObjectClass -eq "Contact" -and msRTCSIP-OwnerUrn -eq "urn:application:Caa" -and msRTCSIP-Line -like $Line_URI -and msRTCSIP-UserEnabled -eq $True } -Properties Displayname, telephoneNumber | Select-Object Displayname, telephoneNumber
    $Result_Trusted_Application_Endpoints = Get-ADObject -SearchBase $SearchBase -Filter {ObjectClass -eq "Contact" -and msRTCSIP-OwnerUrn -like "urn:application:*" -and msRTCSIP-Line -like $Tel_Number -and msRTCSIP-UserEnabled -eq $True } -Properties Displayname, telephoneNumber, msRTCSIP-OwnerUrn | Where-Object msRTCSIP-OwnerUrn -notcontains "urn:application:RGS" | Where-Object msRTCSIP-OwnerUrn -ne "urn:application:Caa" | Select-Object Displayname, telephoneNumber
    
    $User = $Result_Users.Displayname
    $Common_Area_Phone = $Result_Common_Area_Phones.Displayname
    $Analog_Devices = $Result_Analog_Devices.Displayname
    $Response_Group_Service = $Result_Response_Group_Services.Displayname
    $DialIn_Conferencing_Access_Number = $Result_DialIn_Conferencing_Access_Numbers.Displayname
    $Trusted_Application_Endpoint = $Result_Trusted_Application_Endpoints.Displayname
   
    If (($Result_Users | Measure-Object ).Count -eq 1) {
    
       "<tr><td style=background-color:red;color:white>$Line_URI</td><td>$User</td><td>Skype for Business</td><td>Vergeben</span></td></tr>" | Out-File $File -append  
    }
    
    If (($Result_Common_Area_Phones | Measure-Object ).Count -eq 1) {
    
       "<tr><td style=background-color:red;color:white>$Line_URI</td><td>$Common_Area_Phone</td><td>Common Area Phone</td><td>Vergeben</td></tr>" | Out-File $File -append
    }
    
    If (($Result_Response_Group_Services | Measure-Object ).Count -eq 1) {
 
       "<tr><td style=background-color:red;color:white>$Tel_Number</td><td>$Response_Group_Service</td><td>Response Group Service</td><td>Vergeben</td></tr>" | Out-File $File -append
    }
 
    If (($Result_DialIn_Conferencing_Access_Numbers | Measure-Object ).Count -eq 1) {
 
       "<tr><td style=background-color:red;color:white>$Tel_Number</td><td>$DialIn_Conferencing_Access_Number</td><td>Conferencing Dial-In</td><td>Vergeben</td></tr>" | Out-File $File -append
    }
 
    If (($Result_Trusted_Application_Endpoints | Measure-Object ).Count -eq 1) {
 
       "<tr><td style=background-color:red;color:white>$Tel_Number</td><td>$Trusted_Application_Endpoint</td><td>Trusted Application End Point</td><td>Vergeben</td></tr>" | Out-File $File -append
    }
 
    If (($Result_Analog_Devices | Measure-Object ).Count -eq 1) {
 
       "<tr><td style=background-color:red;color:white>$Tel_Number</td><td>$Analog_Devices</td><td>Analog Device</td><td>Vergeben</td></tr>" | Out-File $File -append
    }
 
    Elseif ((($Result_Users | Measure-Object ).Count -eq 0) -and (($Result_Common_Area_Phones | Measure-Object ).Count -eq 0) -and (($Result_Response_Group_Services | Measure-Object ).Count -eq 0) -and (($Result_DialIn_Conferencing_Access_Numbers | Measure-Object ).Count -eq 0) -and (($Result_Trusted_Application_Endpoints | Measure-Object ).Count -eq 0) -and (($Result_Analog_Devices | Measure-Object ).Count -eq 0)) {
   
               
        
        If ($Device=$DID_Liste.DID.Analog_Devices_Reservation.Device | Where-Object {$_.Number -eq $Tel_Reservation})
        { 
            $Device = $Device.Name
            If($Device -match "FAX"){
                "<tr><td style=background-color:Orange;color:white>$Tel_Reservation</td><td>$Device</td><td>Device Reservation</td><td>Vergeben</td></tr>" | Out-File $File -append
            }Else{
 
                "<tr><td style=background-color:Blue;color:white>$Tel_Reservation</td><td>$Device</td><td>Device Reservation</td><td>Vergeben</td></tr>" | Out-File $File -append
            }
        }
 
        Else
        {
            "<tr><td style=background-color:green;color:white>$Line_URI</td><td></td><td></td><td>Verf�gbar</td></tr>" | Out-File $File -append
        }
 
    }
}
 
"<body>" | Out-File $File -append
"<h1>Telefon Nummern Inventar der $Firma | Erstellt am $Date</h1>" | Out-File $File -append
 
"<ul class=tab>" | Out-File $File -append
"<li><a href=javascript:void(0) class=tablinks onclick=openCity(event,'Numbers') id='defaultOpen'>Nummern Blocks</a></li>" | Out-File $File -append
"<li><a href=javascript:void(0) class=tablinks onclick=openCity(event,'Statistik')>Statistik</a></li>" | Out-File $File -append
"<li><a href=javascript:void(0) class=tablinks onclick=openCity(event,'Chart')>Chart</a></li>" | Out-File $File -append
"<li><a href=javascript:void(0) class=tablinks onclick=openCity(event,'Reservationen')>Reservationen</a></li>" | Out-File $File -append
"<li><a href=javascript:void(0) class=tablinks onclick=openCity(event,'Suchen')>Suchen</a></li>" | Out-File $File -append
 
$Standort_Liste=$DID_Liste.DID.Number_Blocks.Number_Block.ForEach({[PSCustomObject]$_}) | Sort-Object Standort -Unique | Select-Object Standort
 
Foreach ($Standort_Name in $Standort_Liste){
    
    [String]$Standort_Display_Name = $Standort_Name.standort
    "<li><a href=javascript:void(0) class=tablinks onclick=openCity(event,'$Standort_Display_Name')>$Standort_Display_Name</a></li>" | Out-File $File -append
}
 
"<li><a href=javascript:void(0) class=tablinks onclick=openCity(event,'About')>About</a></li>" | Out-File $File -append
"</ul>" | Out-File $File -append
 
Foreach ($Standort_Name in $Standort_Liste){
 
    [String]$Standort_Display_Name = $Standort_Name.standort
 
    # DID Lookup pro Standort
    $DIDs= $DID_Liste.DID.Number_Blocks.Number_Block.ForEach({[PSCustomObject]$_}) | Where-Object {$_.Standort -eq $Standort_Display_Name }
 
    "<div id=$Standort_Display_Name class=tabcontent>" | Out-File $File -append
   
    Foreach ($DIDGroup in $DIDs){
   
        [String]$Start_Range = $DIDGroup.StartRange
        [String]$End_Range = $DIDGroup.EndRange
        [String]$DID_Description = $DIDGroup.Description
 
        [Int64]$StartRange = [String]$Start_Range.Replace('+','')
        [Int64]$EndRange = [String]$End_Range.Replace('+','')
 
        "<button class=accordion><b>Nummern Block:</b> $DID_Description [ $Start_Range bis $End_Range ]</button>" | Out-File $File -append
        "<div class=panel>" | Out-File $File -append
        "<p>" | Out-File $File -append  
        "<table id='DDI$StartRange'>" | Out-File $File -append
        "<tr><th style=width:120px>Nummer</th><th style=width:250px>Displayname</th><th style=width:250px>Verwendung</th><th style=width:80px>Verf�gbar</th></tr>" | Out-File $File -append
 
        # Loop Nummern Blocks
        For ($DID=($StartRange); $DID -le ($EndRange); $DID++)
        {
            CheckNumber([String]$DID)
        }
    
        "</table>" | Out-File $File -append
       
 
$Style= @"
<style>
        
        table#$StartRange td:first-child {
        text-align: left;
        padding: 8px;
        cursor: pointer;
     }
 
</style>
"@
 
$Script = @"
<script>
 
    var tbl = document.getElementById('DDI$StartRange');   
 
    if (tbl !=  null) {   
 
        for (var i = 1; i < tbl.rows.length; i++) {   
        tbl.rows[i].cells[0].onclick = function () { getval(this); };   
 
        }   
 
    }   
</script>
"@
 
$Script | Out-File $File -append
 
        "<p>" | Out-File $File -append
        "</div>" | Out-File $File -append   
    
    }
 
    "</div>" | Out-File $File -append
}
 
# Mitarbeiter Suchen
 
$Users=Get-ADUser -Filter {msRTCSIP-Line -like tel:* -and msRTCSIP-UserEnabled -eq $True } -Properties Displayname, sn, givenname, telephoneNumber, mobile, l | Sort-Object givenname
 
"<div id=Suchen class=tabcontent>" | Out-File $File -append
"<input style=width:660px type=text id=Suchen_Eingabe onkeyup=Suchen() placeholder='Suchen von Mitarbeiter Nummern...'>" | Out-File $File -append
 
"<Table id=Mitarbeiter>" | Out-File $File -append
"<thead>" | Out-File $File -append
"<tr><th style=width:200px>Mitarbeiter</th><th style=width:160px>Festnetz Nummer</th><th style=width:160px>Mobile Nummer</th><th style=width:140px>Standort</th></tr>" | Out-File $File -append
"<tbody>" | Out-File $File -append
 
ForEach ($Mitarbeiter in $Users){
 
    $Name = $Mitarbeiter.DisplayName
    $Telefon = $Mitarbeiter.telephoneNumber
    $Mobile = $Mitarbeiter.mobile
    $Standort = $Mitarbeiter.l
    
    "<tr><td>$Name</td><td>$Telefon</td><td>$Mobile</td><td>$Standort</td></tr>" | Out-File $File -append
}
 
"</tbody>" | Out-File $File -append
"</thead>" | Out-File $File -append
"</Table>" | Out-File $File -append
"</div>" | Out-File $File -append
 
# About
"<div id=About class=tabcontent>" | Out-File $File -append
"<p><strong>AR Informatik AG</strong></p>" | Out-File $File -append
"</div>" | Out-File $File -append
 
# Reservationen
"<div id=Reservationen class=tabcontent>" | Out-File $File -append
"<p>Analog Device Reservationen</p>" | Out-File $File -append
"<Table>" | Out-File $File -append
"<tr><th style=width:300px>Analog Device</th><th style=width:250px>Nummer</th></tr>" | Out-File $File -append
 
foreach ($Device in $DID_Liste.DID.Analog_Devices_Reservation.Device)
{
    $DeviceName = $Device.Name
    $DeviceNumber = $Device.Number
 
    "<tr><td>$DeviceName</td><td style=width:250px>$DeviceNumber</td></tr>" | Out-File $File -append
}
 
"</Table>" | Out-File $File -append
"</div>" | Out-File $File -append
 
# Telefon Nummern Statistik erstellen
$Anzahl_Users=0
$Anzahl_Common_Area_Phones=0
$Anzahl_Analog_Phones=0
$Anzahl_Response_Group_Services=0
$Anzahl_Dial_In_Nummern=0
$Anzahl_Trusted_Application_End_Points=0
 
# Anzahl Users
$Anzahl_Users = Get-ADObject -Filter {Objectclass -eq "User" -and msRTCSIP-Line -like tel:* -and msRTCSIP-UserEnabled -eq $True } -Properties Displayname, msRTCSIP-Line, msRTCSIP-UserEnabled
$Anzahl_Users = ($Anzahl_Users | Measure-Object ).Count
 
# Anzahl Common Area Phones
$Anzahl_Common_Area_Phones = Get-ADObject -Filter {ObjectClass -eq "Contact" -and msRTCSIP-OwnerUrn -eq "urn:device:commonareaphone" -and msRTCSIP-Line -like tel:* -and msRTCSIP-UserEnabled -eq $True } -Properties Displayname, msRTCSIP-Line, msRTCSIP-UserEnabled
$Anzahl_Common_Area_Phones = ($Anzahl_Common_Area_Phones | Measure-Object ).Count
 
# Anzahl Analog Devices
$Anzahl_Analog_Phones = Get-ADObject -Filter {ObjectClass -eq "Contact" -and msRTCSIP-OwnerUrn -eq "urn:device:analogphone" -and msRTCSIP-Line -like tel:* -and msRTCSIP-UserEnabled -eq $True } -Properties Displayname, msRTCSIP-Line, msRTCSIP-UserEnabled
$Anzahl_Analog_Phones = ($Anzahl_Analog_Phones | Measure-Object ).Count
 
# Anzahl Response Group Services
$Anzahl_Response_Group_Services = Get-ADObject -SearchBase $SearchBase -Filter {ObjectClass -eq "Contact" -and msRTCSIP-OwnerUrn -eq "urn:application:RGS" -and msRTCSIP-Line -like tel:* -and msRTCSIP-UserEnabled -eq $True } -Properties * | Where-Object Displayname -notcontains "Announcement Service" | Where-Object Displayname -notcontains "RGS Presence Watcher"
$Anzahl_Response_Group_Services = ($Anzahl_Response_Group_Services | Measure-Object ).Count
 
# Anzahl Conference Dial In Nummern
$Anzahl_Dial_In_Nummern = Get-ADObject -SearchBase $SearchBase -Filter {ObjectClass -eq "Contact" -and msRTCSIP-OwnerUrn -eq "urn:application:Caa" -and msRTCSIP-Line -like tel:* -and msRTCSIP-UserEnabled -eq $True } -Properties * | Select-Object *
$Anzahl_Dial_In_Nummern = ($Anzahl_Dial_In_Nummern | Measure-Object ).Count
 
# Anzahl Trusted Application Endpoint
$Anzahl_Trusted_Application_End_Points = Get-ADObject -SearchBase $SearchBase -Filter {ObjectClass -eq "Contact" -and msRTCSIP-OwnerUrn -like "urn:application:*" -and msRTCSIP-Line -like tel:* -and msRTCSIP-UserEnabled -eq $True } -Properties * | Where-Object msRTCSIP-OwnerUrn -ne "urn:application:RGS" | Where-Object msRTCSIP-OwnerUrn -ne "urn:application:Caa"
$Anzahl_Trusted_Application_End_Points = ($Anzahl_Trusted_Application_End_Points | Measure-Object ).Count
 
$Total = ($Anzahl_Users+$Anzahl_Common_Area_Phones+$Anzahl_Analog_Phones+$Anzahl_Response_Group_Services+$Anzahl_Dial_In_Nummern+$Anzahl_Trusted_Application_End_Points)
 
"<div id='Statistik' class=tabcontent>" | Out-File $File -append
"<p>Statistik der vergebenen Skype for Business Telefon Nummern</p>" | Out-File $File -append
"<Table>" | Out-File $File -append
"<tr><th style=width:300px>Verwendung</th><th style=width:250px>Vergebene Nummern</th></tr>" | Out-File $File -append
"<tr><td>Anzahl Users</td><td style=width:250px>$Anzahl_Users</td></tr>" | Out-File $File -append
"<tr><td style=width:250px>Anzahl Common Area Phones</td><td style=width:250px>$Anzahl_Common_Area_Phones</td></tr>" | Out-File $File -append
"<tr><td style=width:250px>Anzahl Analog Phones</td><td style=width:250px>$Anzahl_Analog_Phones</td></tr>" | Out-File $File -append
"<tr><td style=width:250px>Anzahl Response Group Services</td><td style=width:250px>$Anzahl_Response_Group_Services</td></tr>" | Out-File $File -append
"<tr><td style=width:250px>Anzahl Dial-In Nummern</td><td style=width:250px>$Anzahl_Dial_In_Nummern</td></tr>" | Out-File $File -append
"<tr><td style=width:250px>Anzahl Trusted Application End Points</td><td style=width:250px>$Anzahl_Trusted_Application_End_Points</td></tr>" | Out-File $File -append
"<tr><th style=width:250px>Total vergebene Nummern</th><th style=width:250px>$Total</th></tr>" | Out-File $File -append
"</Table>" | Out-File $File -append
"</div>" | Out-File $File -append
 
# Nummern Blocks
"<div id='Numbers' class=tabcontent>" | Out-File $File -append
"<p>Inventar aller Nummern Blöcke</p>" | Out-File $File -append
"<Table>" | Out-File $File -append
"<tr><th style=width:200px>Standort</th><th style=width:200px>Start Block</th><th style=width:200px>End Block</th><th style=width:200px>Anzahl Nummern</th></tr>" | Out-File $File -append
 
$Standort_Liste=$DID_Liste.DID.Number_Blocks.Number_Block.ForEach({[PSCustomObject]$_}) | Sort-Object Standort | Select-Object Standort, StartRange, EndRange, Description
 
Foreach ($Standort_Name in $Standort_Liste){
    
    $Standort_Display_Name = $Standort_Name.Standort
    $Start_Range = $Standort_Name.StartRange
    $End_Range = $Standort_Name.EndRange
    $Description = $Standort_Name.Description
 
    [Int64]$StartRange = [String]$Start_Range.Replace('+','')
    [Int64]$EndRange = [String]$End_Range.Replace('+','')
 
    $AnzahlDID = $EndRange-$StartRange + 1
    
    "<tr><td>$Standort_Display_Name</td><td>$Start_Range</td><td>$End_Range</td><td>$AnzahlDID</td></tr>" | Out-File $File -append
}
 
$TotalDID=0
 
foreach ($DID in $DID_Liste.DID.Number_Blocks.Number_Block) {
    
    $StartRange = $DID.StartRange.Replace('+','')
    $EndRange = $DID.EndRange.Replace('+','')
    $TotalBlock = $EndRange - $StartRange + 1
    $TotalDID=$TotalDID+$TotalBlock
}
 
"<tr><td></td><td></td><th>Total Nummern</th><th>$TotalDID</th></tr>" | Out-File $File -append
 
"</Table>" | Out-File $File -append
"</div>" | Out-File $File -append
 
# Chart
"<div id=Chart class=tabcontent>" | Out-File $File -append
 
$Script = @'
 
   <script language="JavaScript">
 
$(document).ready(function() { 
   var data = {
      table: 'Chart'
   };
   var chart = {
      type: 'column'
   };
   var title = {
      text: 'Telefon Nummern Inventar'   
   };      
   var yAxis = {
      allowDecimals: false,
      title: {
         text: 'Anzahl DID'
      }
   };
   var tooltip = {
      formatter: function () {
         return '<b>' + this.series.name + '</b><br/>' +
            this.point.y + ' ' + this.point.name.toLowerCase();
      }
   };
   var credits = {
      enabled: false
   };  
      
   var json = {};   
   json.chart = chart; 
   json.title = title; 
   json.data = data;
   json.yAxis = yAxis;
   json.credits = credits;  
   json.tooltip = tooltip;  
   $('#Chart').highcharts(json);
});
 
</script>
 
'@
 
$Script | Out-File $File -append
 
Function CountDID($Line_URI){
    
    $Tel_Number = tel:++$Line_URI
    $Extension = $Line_URI.Substring($Line_URI.Length - 4)
    $Line_URI = tel:+$Line_URI#+ ";ext=" + $Extension
      
    # Active Directory Lookup
    $Result_Users = Get-ADObject -Filter {Objectclass -eq "User" -and msRTCSIP-Line -eq $Line_URI -and msRTCSIP-UserEnabled -eq $True } 
    $Result_Common_Area_Phones = Get-ADObject -Filter {ObjectClass -eq "Contact" -and msRTCSIP-OwnerUrn -eq "urn:device:commonareaphone" -and msRTCSIP-Line -like $Tel_Number -and msRTCSIP-UserEnabled -eq $True } 
    $Result_Analog_Devices = Get-ADObject -Filter {ObjectClass -eq "Contact" -and msRTCSIP-OwnerUrn -eq "urn:device:analogphone" -and msRTCSIP-Line -like $Tel_Number -and msRTCSIP-UserEnabled -eq $True } 
    $Result_Response_Group_Services = Get-ADObject -SearchBase $SearchBase -Filter {ObjectClass -eq "Contact" -and msRTCSIP-OwnerUrn -eq "urn:application:RGS" -and msRTCSIP-Line -like $Tel_Number -and msRTCSIP-UserEnabled -eq $True } 
    $Result_DialIn_Conferencing_Access_Numbers = Get-ADObject -SearchBase $SearchBase -Filter {ObjectClass -eq "Contact" -and msRTCSIP-OwnerUrn -eq "urn:application:Caa" -and msRTCSIP-Line -like $Tel_Number -and msRTCSIP-UserEnabled -eq $True } 
    $Result_Trusted_Application_Endpoints = Get-ADObject -SearchBase $SearchBase -Filter {ObjectClass -eq "Contact" -and msRTCSIP-OwnerUrn -like "urn:application:*" -and msRTCSIP-Line -like $Tel_Number -and msRTCSIP-UserEnabled -eq $True } -Properties msRTCSIP-OwnerUrn | Where-Object msRTCSIP-OwnerUrn -notcontains "urn:application:RGS" | Where-Object msRTCSIP-OwnerUrn -ne "urn:application:Caa"
      
    If (($Result_Users | Measure-Object ).Count -eq 1) { Return 1 }
    If (($Result_Common_Area_Phones | Measure-Object ).Count -eq 1) { Return 1}
    If (($Result_Response_Group_Services | Measure-Object ).Count -eq 1) { Return 1 }
    If (($Result_DialIn_Conferencing_Access_Numbers | Measure-Object ).Count -eq 1) { Return }
    If (($Result_Trusted_Application_Endpoints | Measure-Object ).Count -eq 1) { Return 1 }
    If (($Result_Analog_Devices | Measure-Object ).Count -eq 1) { Return 1 }
    If ((($Result_Users | Measure-Object ).Count -eq 0) -and (($Result_Common_Area_Phones | Measure-Object ).Count -eq 0) -and (($Result_Response_Group_Services | Measure-Object ).Count -eq 0) -and (($Result_DialIn_Conferencing_Access_Numbers | Measure-Object ).Count -eq 0) -and (($Result_Trusted_Application_Endpoints | Measure-Object ).Count -eq 0) -and (($Result_Analog_Devices | Measure-Object ).Count -eq 0)) { Return 0 }
}
 
"<body>" | Out-File $File -append
 
"<Table id='Chart' style='visibility:hidden'>" | Out-File $File -append
"<thead>" | Out-File $File -append
"<tr><th></th><th>Total DID</th><th>Verfügbare DID</th></tr>" | Out-File $File -append
"</thead>" | Out-File $File -append
"<tbody>" | Out-File $File -append
 
$Standort_Liste=$DID_Liste.DID.Number_Blocks.Number_Block.ForEach({[PSCustomObject]$_}) | Sort-Object Standort -Unique | Select-Object Standort, StartRange, EndRange, Description
 
Foreach ($Standort_Name in $Standort_Liste){
 
    [String]$Standort_Display_Name = $Standort_Name.standort
 
    # DID Lookup pro Standort
    $DIDs = $DID_Liste.DID.Number_Blocks.Number_Block.ForEach({[PSCustomObject]$_}) | Where-Object {$_.Standort -eq $Standort_Display_Name }
    
    $AnzahlDID=0 
    $TotalDID=0
    $VergebeneDID=0
    $VerfuegbareDID=0
 
    Foreach ($DIDGroup in $DIDs){
   
        [String]$Standort = $DIDGroup.Standort
        [Int64]$StartRange = $DIDGroup.StartRange.Replace('+','')
        [Int64]$EndRange = $DIDGroup.EndRange.Replace('+','')
 
        $AnzahlDID=$EndRange-$StartRange+1
        $TotalDID=$TotalDID+$AnzahlDID
                
        # Loop Nummern Blocks
        For ($DID=($StartRange); $DID -le ($EndRange); $DID++)
        {
            $VergebeneDID=$VergebeneDID-(CountDID([String]$DID))
        }
  
    }
        $VerfuegbareDID=$TotalDID+$VergebeneDID
 
    "<tr><th>$Standort_Display_Name</th><td>$TotalDID</td><td>$VerfuegbareDID</td></tr>" | Out-File $File -append
}
 
"</tbody>" | Out-File $File -append
"</Table>" | Out-File $File -append
"</div>" | Out-File $File -append
 
# HTML Scripts
$Script = @'
 
<script>
 
function Suchen() {
   
    var searchText = document.getElementById('Suchen_Eingabe').value;
    var targetTable = document.getElementById('Mitarbeiter');
    var targetTableColCount;
            
    //Loop through table rows
    for (var rowIndex = 0; rowIndex < targetTable.rows.length; rowIndex++) {
        var rowData = '';
 
        //Get column count from header row
        if (rowIndex == 0) {
           targetTableColCount = targetTable.rows.item(rowIndex).cells.length;
           continue; //do not execute further code for header row.
        }
                
        //Process data rows. (rowIndex >= 1)
        for (var colIndex = 0; colIndex < targetTableColCount; colIndex++) {
            rowData += targetTable.rows.item(rowIndex).cells.item(colIndex).textContent;
        }
 
        //If search term is not found in row data
        //then hide the row, else show
        if (rowData.indexOf(searchText) == -1)
            targetTable.rows.item(rowIndex).style.display = 'none';
        else
            targetTable.rows.item(rowIndex).style.display = 'table-row';
    }
}
 
function openCity(evt, cityName) {
    var i, tabcontent, tablinks;
    tabcontent = document.getElementsByClassName('tabcontent');
    for (i = 0; i < tabcontent.length; i++) {
        tabcontent[i].style.display='none';
    }
    tablinks = document.getElementsByClassName('tablinks');
    for (i = 0; i < tablinks.length; i++) {
        tablinks[i].className = tablinks[i].className.replace(' active', '');
    }
    document.getElementById(cityName).style.display='block';
    evt.currentTarget.className +=' active';
}
 
// Get the element with id=defaultOpen and click on it
document.getElementById(defaultOpen).click();
</script>
 
'@
 
$Script | Out-File $file -append
 
$Script = @'
 
<script>
var acc = document.getElementsByClassName('accordion');
var i;
for (i = 0; i < acc.length; i++) {
    acc[i].onclick = function(){
        this.classList.toggle('active');
        this.nextElementSibling.classList.toggle('show');
  }
}
 
// Get the element with id='defaultOpen' and click on it
document.getElementById('defaultOpen').click();
 
</script>
'@
 
$Script | Out-File $File -append
          
$Script = @'
 
<script>
            
        function getval(cel) {   
 
            window.alert("LineURI wurde ins Clipboard kopiert !");
            window.clipboardData.setData('Text', cel.innerHTML);
        }   
 
</script>
'@
 
$Script | Out-File $File -append
 
"</body>" | Out-File $File -append
"</html>" | Out-File $File -append
 
Write-host "Telefon Nummern Inventar der $Firma erstellt" -ForegroundColor Green
