 
<#
AssetDetails PowerShell module (hughowens@gmail.com) 
Copyright (C) 2016 Hugh Owens 
 
This program is free software: you can redistribute it and/or modify 
it under the terms of the GNU General Public License as published by 
the Free Software Foundation, either version 3 of the License, or 
(at your option) any later version. 
 
This program is distributed in the hope that it will be useful, 
but WITHOUT ANY WARRANTY; without even the implied warranty of 
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the 
GNU General Public License for more details. 
 
You should have received a copy of the GNU General Public License 
along with this program. If not, see <http://www.gnu.org/licenses/>. 
#>

<#
 .Synopsis
  Gathers various information about computer system.
 .Description
  Gathers various information about computer system for Audit purposes
 .Parameter ConfigFile
  Location of the config file to use.
 .Example
   # Show a default asset details
   Get-AssetDetails
 .Example
   # Get Asset details with aid of a config file
   Get-AssetDetails -configfile .\Config.xml
#>

function Get-AssetDetails{
	param (
		[string]$configFile = "$(Split-Path -parent $PSCommandPath)\config.xml",
		[String]$UseActiveDirectory = $true,
		[String]$UseOutlookProfile = $false,
		[String]$CompanyName = "",
		[String]$Notes = "",
		[String]$PurchaseDate = $null,
		[String]$PurchaseCost = $null,
		[String]$Export = $null,
		[String]$Status = $null,
		[String]$Supplier = $null,
		[switch]$NoConfigFile = $false,
		[switch]$FetchWarranty = $false,
		[switch]$GetNetwork = $true,
		[switch]$SkipUser = $false,
		[switch]$GetSoftware = $true
	)
	if ($ConfigFile -ne "" ){
		write-host "Using config file: "$configfile
		[xml] $settings = Get-Content $ConfigFile
	}

	#Get settings from config file:
	if ($settings.Settings.UserInfo.UseActiveDirectory) {$UseActiveDirectory = $settings.Settings.UserInfo.UseActiveDirectory}
	if ($settings.Settings.UserInfo.UseOutlookProfile) {$UseOutlookProfile = $settings.Settings.UserInfo.UseOutlookProfile}
	if ($settings.Settings.UserInfo.UseIPAddressForLocation) {$UseIPAddressForLocation = $settings.Settings.UserInfo.UseActiveDirectory}
	if ($settings.Settings.Company.CompanyName) {$CompanyName = $settings.Settings.Company.CompanyName}
	if ($settings.Settings.GetNetwork) {$GetNetwork = [System.Convert]::ToBoolean($settings.Settings.GetNetwork)}
	if ($settings.Settings.GetSoftware) {$GetSoftware = [System.Convert]::ToBoolean($settings.Settings.GetSoftware)}
    if ($settings.Settings.GetFanInfo) {$GetFanInfo = [System.Convert]::ToBoolean($settings.Settings.GetFanInfo)}
	if ($settings.Settings.Company.SaveLocation -ne $null) {$SAVELOCATION=$settings.Settings.Company.SaveLocation+"\"+$settings.Settings.Company.OutPutFilePrefix+$env:computername+".csv"}
	if ($settings.Settings.DUMPEDID -ne $null) {$DUMPEDID=$settings.Settings.DUMPEDID}else{$DUMPEDID = "dumpedid.exe"}
	if ($export){$SAVELOCATION = $export}
	 
	### User Information
	$UserName = $env:UserName
	#Find User details through ADSI
	if ($UseActiveDirectory -eq $true){
		if ($username -eq $null){
			$username = (get-itemproperty -path registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Authentication\LogonUI  | select-object -expandproperty lastloggedonuser) -creplace '^[^\\]*\\', ''
		}
		$User = [adsisearcher]"(samaccountname=$USERNAME)"
		$Name = [String] $User.FindOne().Properties.cn
		$Email = [String] $User.FindOne().Properties.mail
	}

	#Find User Details through Default Outlook Profile
	if ($UseOutlookProfile -eq $true){
		write-host "Using Outlook for User information"
		try{
			$Outlook = (New-Object -comobject Outlook.Application)
			$OutlookInfo =  $Outlook.Session.Accounts | select-object smtpaddress,UserName,Name
			$OutlookInfo.name = $Outlook.GetNameSpace("MAPI").Session.currentuser.name
			$UserName = $OutlookInfo.Username
			$Email = $OutlookInfo.smtpaddress
			$Name = $Outlook.GetNameSpace("MAPI").Session.currentuser.name
		}catch{
			Write-host "Not Outlook object found..."
		}
	}

	#get CimInstance Preferred, if not use get-WmiObject
	$cmdName = "Get-CimInstance"
	if (Get-Command $cmdName -errorAction SilentlyContinue){	
		$getinfo="Get-CimInstance"
	}else{
		$getinfo="Get-WMIObject"
	}
	$DETAILSBIOS=Invoke-Expression "$getinfo  Win32_BIOS | Select-Object SerialNumber,Manufacturer"
    $DETAILSFAN=Invoke-Expression "$getinfo Win32_Fan"
	$DETAILSCOMP=Invoke-Expression "$getinfo win32_computersystem | select-object model,manufacturer,SystemSKUNumber,name"
	$DETAILSOS=Invoke-Expression "$getinfo win32_operatingsystem"
	$DETAILSProcessor = Invoke-Expression "$getinfo win32_Processor" | select-object Name
	$Memory = Get-WmiObject CIM_PhysicalMemory | Measure-Object -Property capacity -Sum | % {[math]::round(($_.sum / 1GB),2)}
	$EDID = invoke-expression $DUMPEDID -ErrorAction SilentlyContinue
	
	$monitors = ($edid | select-string -Pattern "Active                   : Yes" -AllMatches -Context 4 | out-string) -split "`r`n" | where {$_ -match "Mon"}
	
	#$EDID = ($EDID | where-object {$_ -match "Monitor Name"})#.trim("Monitor Name :")
	$EDID = $monitors
	
	$ItemName=$DETAILSCOMP.name
	### Network Information
	if ($GetNetwork -eq $TRUE){
		$DETAILSNET=Invoke-Expression "$getinfo Win32_NetworkAdapterConfiguration"
		$DETAILSNET = $DETAILSNET | where-object {$_.IPaddress -ne $null} | select DHCPEnabled,IPaddress, DNSServerSearchOrder,MACAddress,DefaultIPGateway
		if($detailsnet -isnot [system.array]){$detailsnet = @($detailsnet)}
		if($DETAILSNET -is [system.array]){
			$IPAddress= ($DETAILSNET |foreach {$_.ipaddress}) -join ', '
			$DNSServers= ($DETAILSNET | foreach {$_.DNSServerSearchOrder}) -join ', '
			$MACAddress = ($DETAILSNET | foreach {$_.MACAddress}) -join ', '
			$DefGateway = ($DETAILSNET | foreach {$_.DefaultIPGateway}) -join ', '
			$DHCP = ($DETAILSNET | foreach {$_.DHCPEnabled}) -join ', '
		}else{
			$IPAddress= $DETAILSNET.ipaddress  -join ', '
			$DNSServers= $DETAILSNET.DNSServerSearchOrder  -join ', '
			$MACAddress = $DETAILSNET.MACAddress
			$DefGateway = $DETAILSNET.DefaultIPGateway
			$DHCP = $DETAILSNET.DHCPEnabled
		}
	}
	switch -regex ($DETAILSCOMP.Model){	
		"600|6200|6300|Desk|desk|Veriton|Compaq Elite" 		{$CATEGORY="Desktop"}
		"640|1040|947|Book|book" 							{$CATEGORY="Laptop"}
        "Surface"											{$CATEGORY="2 in 1"}
		"VMware"											{$CATEGORY="Server"}
		default 											{$CATEGORY="Unknown"}
	}
	#manufacturer specific information gathering
	switch -regex ($DETAILSCOMP.Manufacturer){	
		"HP|Hewlett-Packard" 	{ # Get Model Number for HP Computers 
									if ($DETAILSCOMP.SystemSKUNumber -eq $null){
										$ModelNumber = (Invoke-Expression "($getinfo -Class 'MS_SystemInformation' -Namespace 'root\WMI' -ErrorAction Stop).SystemSKU.Trim()")
										#$ModelNumber = get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\BIOS | select-object -expandProperty SystemSKU
									}else{
										$ModelNumber = $DETAILSCOMP.SystemSKUNumber
									}	
								}
		"Acer"					{$ModelNumber = $DETAILSCOMP.Model}
		"Apple"					{$ModelNumber = $DETAILSCOMP.Model}
			
		#Do nothing for these devices
		default 				{}
	}
		
	$OS=$DETAILSOS.caption
	$CreatedTime = Get-ItemProperty 'C:\System Volume Information' | select CreationTime
	$ImageInstallDate = ([datetime] $CreatedTime.CreationTime).ToString('yyyy-MM-dd')
    $LastBoot = ([datetime] $DETAILSOS.lastbootuptime).ToString('yyyy-MM-dd')

	try{	#Powershell 3.0+
		$OSInstallDate = ([datetime] $DETAILSOS.installdate).ToString('yyyy-MM-dd')
		#$DefRoute = Get-NetRoute | where-object {$_.destinationprefix -like "0.0.0.0/0"} |  select-object -expandProperty nexthop
	}catch{	
        #Use Powershell 2 commands instead
		$OSInstallDate = $detailsos.converttodatetime($detailsOS.installdate).ToString('yyyy-MM-dd')
		#$DefRoute= Get-WmiObject -Class Win32_IP4RouteTable | where { $_.destination -eq '0.0.0.0' -and $_.mask -eq '0.0.0.0'} | Sort-Object metric1 |  select-object -expandProperty nexthop
	}
	#Location Based on IP Address
	if ($settings.Settings.UseIPAddressForLocation){
		$Location = ($settings.Settings.Locations.Location | where-object {$DefGateway -match $_.IP}) | select -expandproperty name
		if ($Location -is [system.array]) {$location = $location -join ', '}
	}
	#Trying HP Warranty
	if ($FetchWarranty){
		try {
			Import-Module HPWarranty
			Write-Host "Module HPWarranty exists - checking for Warranty..."
			# Get info to send info off to HP for warranty check
			$obj = new-object PSObject
			$obj | add-member -membertype NoteProperty -name "SerialNumber" -value $DETAILSBIOS.SerialNumber
			$obj | add-member -membertype NoteProperty -name "ProductNumber" -value $ModelNumber
			#send info
			$warrantyDetails = $obj | Get-HPIncWarrantyEntitlement
			#Break out if no warranty could be found
			if ($warrantydetails.OverallEntitlementStartDate -eq $null) {break}
			#get time
			$enddate=[datetime]$warrantyDetails.OverallEntitlementEndDate
			$startdate=[datetime]$warrantyDetails.OverallEntitlementStartDate
			$WarrantyStartDateString = 	([datetime] $startdate).ToString('yyyy-MM-dd')
			$WarrantyEndDateString = 	([datetime] $enddate).ToString('yyyy-MM-dd')
			#Get warranty length in months
			$monthdiff = $enddate.month - $startDate.month + (($enddate.Year - $startDate.year)*12)
			$WarrantyMonths = $monthdiff
			$PurchaseDate =  $WarrantyStartDateString
			
		} catch {
			Write-Host "Module does not exist"
		}
	}
	
	$LastUpdate = get-date -format ('yyyy-MM-dd')
	#Creating array for CSV import
	$obj = new-object PSObject
	$obj | add-member -membertype NoteProperty -name "Name" -value $name
	$obj | add-member -membertype NoteProperty -name "Email" -value $Email
	$obj | add-member -membertype NoteProperty -name "Username" -value $UserName
	$obj | add-member -membertype NoteProperty -name "Item Name" -value $ItemName
	$obj | add-member -membertype NoteProperty -name "Category" -value $CATEGORY
	$obj | add-member -membertype NoteProperty -name "Model Name" -value $DETAILSCOMP.model
	$obj | add-member -membertype NoteProperty -name "Manufacturer" -value $DETAILSCOMP.manufacturer
	$obj | add-member -membertype NoteProperty -name "Model Number" -value $ModelNumber
	$obj | add-member -membertype NoteProperty -name "Serial Number" -value $DETAILSBIOS.SerialNumber
	$obj | add-member -membertype NoteProperty -name "Asset Tag" -value $ItemName
	$obj | add-member -membertype NoteProperty -name "Location" -value $Location
	$obj | add-member -membertype NoteProperty -name "Notes" -value $Notes
	$obj | add-member -membertype NoteProperty -name "Purchase Date" -value $PurchaseDate
	$obj | add-member -membertype NoteProperty -name "Purchase Cost" -value $PurchaseCost
	$obj | add-member -membertype NoteProperty -name "Company" -value $CompanyName
	$obj | add-member -membertype NoteProperty -name "Status" -value $status
	$obj | add-member -membertype NoteProperty -name "Warranty Months" -value $WarrantyMonths
	$obj | add-member -membertype NoteProperty -name "Warranty Start" -value $WarrantyStartDateString
	$obj | add-member -membertype NoteProperty -name "Warranty End" -value $WarrantyEndDateString
	$obj | add-member -membertype NoteProperty -name "Supplier" -value $Supplier
	$obj | add-member -membertype NoteProperty -name "Image Install Date" -value $ImageInstallDate
	$obj | add-member -membertype NoteProperty -name "OS Install Date" -value $OSInstallDate
	$obj | add-member -membertype NoteProperty -name "Operating System" -value $OS
	$obj | add-member -membertype NoteProperty -name "Last Update" -value $LastUpdate
    $obj | add-member -membertype NoteProperty -name "Last Boot" -value $LastBoot
	$obj | add-member -membertype NoteProperty -name "IP Address" -value $IPAddress
	$obj | add-member -membertype NoteProperty -name "DNS Servers" -value $DNSServers
	$obj | add-member -membertype NoteProperty -name "DHCP Enabled" -value $DHCP
	$obj | add-member -membertype NoteProperty -name "MAC Address" -value $MACAddress
	$obj | add-member -membertype NoteProperty -name "Processor" -value $DETAILSProcessor.Name
	$obj | add-member -membertype NoteProperty -name "Memory (GB)" -value $Memory
	$obj | add-member -membertype NoteProperty -name "Powershell Version" -value $PSVersionTable.PSVersion.Major
    For ($i=0; $i -lt 3; $i++) {
		if ($EDID -ne $null -and $EDID[$i] -ne $null){
			$monitor = ([string]$EDID[$i]).trim("Monitor Name :")
		}else{
			$monitor = ""
		}
		$obj | add-member -membertype NoteProperty -name "Monitor $($i+1)" -value $monitor; 
    }	
	if ($GetSoftware){
		$software = get-childitem -path registry::HKEY_LOCAL_MACHINE\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\  | foreach-object {get-itemproperty $_.PsPath} | select displayname,displayversion
		$software += get-childitem -path registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Uninstall\  | foreach-object {get-itemproperty $_.PsPath} | select-object displayname,displayversion
	
		$importantSoftware = @{
			Project = @{
				Name = "Microsoft Project";
				include = "Microsoft.*( Project )";
				exclude = "Update|Service|Visio"
			}
			Visio = @{
				Name = "Microsoft Visio";
				include = "Microsoft.*( Visio )";
				exclude = "Update|Service|Project"
			}
			Office = @{
				Name = "Microsoft Office";
				include = "Microsoft Office.*( Standard | Professional )";
				exclude = "Update|Service|Project|Visio"
			}
		}
		
		foreach ($key in $importantSoftware.keys){
			$sw = $software | where-object {$_.displayname -match $importantsoftware.$key.include -and $_.displayname -notmatch $importantsoftware.$key.exclude} | select-object displayname,displayversion -unique
			$obj | add-member -membertype NoteProperty -name $importantsoftware.$key.Name -value ($sw.DisplayName -join ',')	
		}
	}
    if ($GetFanInfo){
        For ($i=0; $i -lt 3; $i++) {
		    if ($DetailsFan -ne $null -and $DetailsFan[$i] -ne $null){
			    $fan = $DetailsFan[$i].status
		    }else{
			    $fan = ""
		    }
		    $obj | add-member -membertype NoteProperty -name "Fan $($i+1)" -value $fan;
        }
    }
	if ($SAVELOCATION) {$obj | export-csv $SAVELOCATION -notypeinformation}
	return $obj
}

function Get-MonitorDetails{
    param (
		[string]$configFile = "$(Split-Path -parent $PSCommandPath)\config.xml"
    )
    if ($ConfigFile -ne "" ){
		#write-host "Using config file: "$configfile
		[xml] $settings = Get-Content $ConfigFile
	}
    if ((test-path "$PSCommandPath\DumpEDID\DumpEDID.exe") -eq $false -and $settings.Settings.DUMPEDID -ne $null) {$DUMPEDID=$settings.Settings.DUMPEDID}else{$DUMPEDID = "$PSCommandPath\DumpEDID\DumpEDID.exe"}

    $Monitors = @()
    $EDID = invoke-expression $DUMPEDID -ErrorAction SilentlyContinue	
    $MonitorArray = ($EDID | out-string) -split "`r`n`r`n`r`n`r`n"
    foreach ($Monitor in $MonitorArray) {
        if ($Monitor -match "Active"){
            $MonDetails = $Monitor.split("`r`n") | ? {$_ -match "Active|Monitor Name|Serial Number" -and $_ -notmatch "Numeric"}
            $Active = [bool] (([string]($MonDetails | ? {$_ -match "Active"})).trim("Active :") -match "Yes")
            $MonitorObj = new-object PSObject
	        $MonitorObj | add-member -membertype NoteProperty -name "Name" -value ([string]($MonDetails | ? {$_ -match "Monitor Name"})).trim("Monitor Name :")
	        $MonitorObj | add-member -membertype NoteProperty -name "Serial Number" -value ([string]($MonDetails | ? {$_ -match "Serial Number"})).trim("Serial Number :")
            $MonitorObj | add-member -membertype NoteProperty -name "Active" -value $Active
            $Monitors += $MonitorObj
        }else{continue}
    }
    return $Monitors
}

Function Get-Software{
	param (
		[string]$configFile = $null
	)
	$software = (Get-AssetDetails).software -split ','
	$obj = new-object PSObject
	$obj | add-member -membertype NoteProperty -name "Software" -value $software
	return $obj
}

function Get-AllComputers{
# Requires module HPWarranty
# https://www.powershellgallery.com/packages/HPWarranty/2.6.2
# Install-Module -Name HPWarranty
	param (
		[string]$configFile = "$(Split-Path -parent $PSCommandPath)\config.xml",
		[String]$UseActiveDirectory = $true,
		[String]$CompanyName = $null,
		[String]$Directory = $null,
		[String]$SaveName = $null,
		[String]$Filter = $null,
		[String]$Notes = "",
		[String]$Export = $null,
		[String]$Status = $null,
		[String]$Supplier = $null,
		[switch]$NoConfigFile = $false,
		[switch]$FetchWarranty = $false,
		[switch]$GetNetwork = $true
	)
	if ($ConfigFile -ne "" ){
		write-host "Using config file: "$configfile
		[xml] $settings = Get-Content $ConfigFile
	}
	
	if ($settings.Settings.UserInfo.UseActiveDirectory) {$UseActiveDirectory = $settings.Settings.UserInfo.UseActiveDirectory}
	if ($settings.Settings.UserInfo.UseIPAddressForLocation) {$UseIPAddressForLocation = $settings.Settings.UserInfo.UseActiveDirectory}
	if ($settings.Settings.Company.CompanyName) {$CompanyName = $settings.Settings.Company.CompanyName}
	if ($settings.Settings.GetNetwork) {$GetNetwork = [System.Convert]::ToBoolean($settings.Settings.GetNetwork)}
	if ($settings.Settings.SuperListName) {$SaveName = $settings.Settings.SuperListName}
	if ($settings.Settings.Company.SaveLocation -ne $null) {$SAVELOCATION=$settings.Settings.Company.SaveLocation+"\"+$settings.Settings.Company.OutPutFilePrefix+$env:computername+".csv"}
    if ($export){$SAVELOCATION = $export}

	if ($Settings.settings.Company.OutPutFilePrefix){
	    $FilterString = "*"+$settings.Settings.Company.OutPutFilePrefix+"*.csv"
	}elseif(!$Filter){$Filter=Read-Host 'Please enter Filter String'}
	else{
		$FilterString=$Filter
	}
	if ($settings.settings.Company.SaveLocation){
	    $Directory = $settings.Settings.Company.SaveLocation
	}
	elseif(!$Directory){
	    $Directory=Read-Host 'Please enter the directory to find computers'
	}
	if ($settings.settings.Company.SuperListName){
	    $SaveName = $settings.Settings.Company.SuperListName
	}
	elseif(!$SaveName){
	    $Savename = Read-Host 'Please the name of the existing file if it exists'
	}

	$SAVELOCATION = $Directory+"\"+$SaveName
	#FYI - Need to sort by last write time incase you add fields so that the first import-csv uses everything
	write-host "Importing from $directory"
	$Computers = Get-ChildItem $Directory | where-object {$_.fullname -like $FilterString -and $_.fullname -notmatch $SaveName} | 
	        sort-object LastWriteTime -desc | select -expandproperty fullname | import-csv
	$ComputersPreviouslyImported = import-csv $SAVELOCATION
	foreach ($computer in $computers) {
		#Get previous Computer import if it exists already
		$computerAlreadyImported = $ComputersPreviouslyImported | where-object {$_."Asset Tag" -eq $Computer."Asset Tag"}
		# Get full name and email from Username - this always replaces existing info
		if ($computer."Email" -eq "" -or $computer."Name" -eq ""){
			if ($computer.username -like '*$') {break}
			$user = get-aduser -identity $computer.username -properties mail,Name
			if ($user.mail -ne $null){$computer."Email" = $user.mail}else{$computer."Email" = $user.userPrincipalName}
			$computer."Name" = $user.Name
			write-host $computer."Asset Tag" "- Fetched name -" $computer."Name"
		}
		# Get model if unknown
		if ($computer."Category" -eq "Unknown"){
			switch -regex ($computer."Model Name")
			{	
				"600|6200|6300|Desk|desk|Veriton|Compaq Elite" 		{$CATEGORY="Desktop"}
				"640|1040|947|Book|book" 							{$CATEGORY="Laptop"}
				"VMware|Virtual Machine"							{$CATEGORY="Server"}
				default 											{$CATEGORY="Unknown"}
			}
		}
		# Get Warranty Information
		:WarrantyBreak
		switch -regex ($computer.Manufacturer)
		{	
			"HP|Hewlett-Packard" 	{ # Get Warranty for HP Computers 
			
				if ($FetchWarranty -eq $FALSE){
					if ($computerAlreadyImported."Warranty Months" -ne ""){
					#	write-host $computerAlreadyImported."Asset Tag" "`tWarranty already Exists." 
						$computer."Warranty Months" = $computerAlreadyImported."Warranty Months"
						$computer."Purchase Date" = $computerAlreadyImported."Purchase Date"
						if([bool]($computer.psobject.Properties | where { $_.Name -eq "Warranty Start"})){
							$computer."Warranty Start" =  $computerAlreadyImported."Warranty Start"
						}else{
							$computer | add-member -membertype NoteProperty -name "Warranty Start" -value $computerAlreadyImported."Warranty Start"
						}
						if([bool]($computer.psobject.Properties | where { $_.Name -eq "Warranty End"})){
							$computer."Warranty End" =  $computerAlreadyImported."Warranty End"
						}else{
							$computer | add-member -membertype NoteProperty -name "Warranty End" -value $computerAlreadyImported."Warranty End"
						}
				#		$computer."Warranty End" = 
				#		$computer."Warranty Start" = 
						break WarrantyBreak
					}
				}
				If ( ! (Get-module hpwarranty )) {
					break WarrantyBreak
				}
				if ($computer."Model Number" -ne "" -and $computer."Serial Number" -ne "") {
				#	write-host $computer."Asset Tag"" - fetching warranty"
					# Get info to send info off to HP for warranty check
					$obj = new-object PSObject
					$obj | add-member -membertype NoteProperty -name "SerialNumber" -value $computer."Serial Number"
					$obj | add-member -membertype NoteProperty -name "ProductNumber" -value $computer."Model Number"
					#send info
					$warrantyDetails = $obj | Get-HPIncWarrantyEntitlement
					#Break out if no warranty could be found
					if ($warrantydetails.OverallEntitlementStartDate -eq $null) {break}
					#get time
					$enddate=[datetime]$warrantyDetails.OverallEntitlementEndDate
					$startdate=[datetime]$warrantyDetails.OverallEntitlementStartDate
					$WarrantyStartDateString = 	([datetime] $startdate).ToString('yyyy-MM-dd')
					$WarrantyEndDateString = 	([datetime] $enddate).ToString('yyyy-MM-dd')
					#Get warranty length in months
					$monthdiff = $enddate.month - $startDate.month + (($enddate.Year - $startDate.year)*12)
					$computer."Warranty Months" = $monthdiff
					$computer."Purchase Date" =  $WarrantyStartDateString
					if([bool]($computer.psobject.Properties | where { $_.Name -eq "Warranty Start"})){
						$computer."Warranty Start" =  $WarrantyStartDateString
					}else{
						$computer | add-member -membertype NoteProperty -name "Warranty Start" -value $WarrantyStartDateString
					}
					if([bool]($computer.psobject.Properties | where { $_.Name -eq "Warranty End"})){
						$computer."Warranty End" =  $WarrantyEndDateString
					}else{
							$computer | add-member -membertype NoteProperty -name "Warranty End" -value $WarrantyEndDateString
					}
					write-host $computer."Asset Tag"  "`tWarranty end date:`t"$enddate
					$computers | export-csv $SAVELOCATION -notypeinformation					
				}
			}
			#Do nothing for these devices
			default 								{}
		}
		#Copy remaining info if missing
		foreach ($item in Get-Member -in $computer){
			if ($computer.($item.name) -eq "" -and $computerAlreadyImported -ne $null -and $computerAlreadyImported.($item.name) -ne ""){
				write-host $computer."Asset Tag" "- writing item" $item.name "-"$computerAlreadyImported.($item.name)
			}
		}
	}
	$computers | export-csv $SAVELOCATION -notypeinformation
	return $computers
	#$computers | where-object {$_.email -ne ""} | export-csv $SAVELOCATION -notypeinformation
}
function Get-AllSoftware{
	param (
		[string]$configFile = $null,
		[String]$UseActiveDirectory = $true,
		[String]$CompanyName = $null,
		$computers = $null,
		[String]$Directory = $null,
		[String]$SaveName = $null,
		[String]$Notes = "",
		[String]$Export = $null,
		[String]$Status = $null,
		[String]$Supplier = $null,
		[switch]$NoConfigFile = $false,
		[switch]$count = $false
	)
	if ($ConfigFile -ne "" ){
		write-host "Using config file: "$configfile
		[xml] $settings = Get-Content $ConfigFile
	}
	
	if ($settings.Settings.UserInfo.UseActiveDirectory) {$UseActiveDirectory = $settings.Settings.UserInfo.UseActiveDirectory}
	if ($settings.Settings.UserInfo.UseIPAddressForLocation) {$UseIPAddressForLocation = $settings.Settings.UserInfo.UseActiveDirectory}
	if ($settings.Settings.Company.CompanyName) {$CompanyName = $settings.Settings.Company.CompanyName}
	if ($settings.Settings.GetNetwork) {$GetNetwork = [System.Convert]::ToBoolean($settings.Settings.GetNetwork)}
	if ($settings.Settings.SuperListName) {$SaveName = $settings.Settings.SuperListName}
	if ($settings.Settings.Company.SaveLocation -ne $null) {$SAVELOCATION=$settings.Settings.Company.SaveLocation+"\"+$settings.Settings.Company.OutPutFilePrefix+$env:computername+".csv"}
	if ($export){$SAVELOCATION = $export}

	if ($Settings.settings.Company.OutPutFilePrefix){$FilterString = "*"+$settings.Settings.Company.OutPutFilePrefix+"*.csv"
	}elseif(!$Filter){$Filter=Read-Host 'Please enter Filter String'
	}else{$FilterString=$Filter}
	if ($settings.settings.Company.SaveLocation){$Directory = $settings.Settings.Company.SaveLocation
	}elseif(!$Directory){$Directory=Read-Host 'Please enter the directory to find computers'}
	if ($settings.settings.Company.SuperListName){$SaveName = $settings.Settings.Company.SuperListName
	}elseif(!$SaveName){$Savename = Read-Host 'Please the name of the existing file if it exists'}

	$SAVELOCATION = $Directory+"\"+$SaveName
	#FYI - Need to sort by last write time incase you add fields so that the first import-csv uses everything
	write-host "Importing from $directory"
	if($computers -eq $null){$Computers = Get-ChildItem $Directory | where-object {$_.fullname -like $FilterString -and $_.fullname -notmatch $SaveName} | sort-object LastWriteTime -desc | select -expandproperty fullname | import-csv | where-object {$_.software -ne $null -and $_.software -ne ""}}
	$ComputersPreviouslyImported = import-csv $SAVELOCATION
	
	$softwareList = @(new-object PSObject)
	foreach ($computer in $computers){
			$softwareString = $computer.software -split ','
			foreach ($string in $softwareString) {
				if ($string -eq "") {break}
				$software = $softwareList | where-object {$_.name -eq $String}
				if($software -and $software.Computers -notmatch $Computer."Asset Tag"){
					$software.Computers += $Computer."Asset Tag"
					$software.Total++
				}elseif(!$software){
					$software = new-object PSObject
					$software | add-member -membertype NoteProperty -name "Name" -value $String
					$software | add-member -membertype NoteProperty -name "Total" -value 1
					$software | add-member -membertype NoteProperty -name "Computers" -value @($Computer."Asset Tag")		
					$softwarelist += $software
				}else{
					#Do Nothing
				}
			}	
		#}
	}
	return $softwareList
}

function Get-AllSoftwareFormatted{
	param (
		$sofwareList = $null,
		$computers = $null,
		$format = "csv"
	)
	if ($ConfigFile -ne "" ){
		write-host "Using config file: "$configfile
		[xml] $settings = Get-Content $ConfigFile
	}
	if ($settings.Settings.UserInfo.UseActiveDirectory) {$UseActiveDirectory = $settings.Settings.UserInfo.UseActiveDirectory}
	if ($settings.Settings.UserInfo.UseIPAddressForLocation) {$UseIPAddressForLocation = $settings.Settings.UserInfo.UseActiveDirectory}
	if ($settings.Settings.Company.CompanyName) {$CompanyName = $settings.Settings.Company.CompanyName}
	if ($settings.Settings.GetNetwork) {$GetNetwork = [System.Convert]::ToBoolean($settings.Settings.GetNetwork)}
	if ($settings.Settings.SuperListName) {$SaveName = $settings.Settings.SuperListName}
	if ($settings.Settings.Company.SaveLocation -ne $null) {$SAVELOCATION=$settings.Settings.Company.SaveLocation+"\"+$settings.Settings.Company.OutPutFilePrefix+$env:computername+".csv"}
	if ($export){$SAVELOCATION = $export}

	if ($Settings.settings.Company.OutPutFilePrefix){$FilterString = "*"+$settings.Settings.Company.OutPutFilePrefix+"*.csv"}
		elseif(!$Filter){$Filter=Read-Host 'Please enter Filter String'}
		else{$FilterString=$Filter}
	if ($settings.settings.Company.SaveLocation){$Directory = $settings.Settings.Company.SaveLocation}
		elseif(!$Directory){$Directory=Read-Host 'Please enter the directory to find computers'}
	if ($settings.settings.Company.SuperListName){$SaveName = $settings.Settings.Company.SuperListName}
		elseif(!$SaveName){$Savename = Read-Host 'Please the name of the existing file if it exists'}	
		
	if($computers -eq $null){$Computers = Get-ChildItem $Directory | where-object {$_.fullname -like $FilterString -and $_.fullname -notmatch $SaveName} | sort-object LastWriteTime -desc | select -expandproperty fullname | import-csv | where-object {$_.software -ne $null -and $_.software -ne ""}}
	if($softwareList -eq $null){$softwareList = Get-AllSoftware -configfile $configfile}
	$Licenses = @()

	foreach ($software in $softwareList){
		write-host "Software: "$software.name
		switch ($format){
			"csv" 		{
				write-host "Software: "$software.name
				if (!$License) {$License = new-object PSObject}
				$computers = $software.computers | select -expandproperty computers
				$License | add-member -membertype NoteProperty -name $software.name -value $computers
				$Licenses = $License
			}
			"snipeit"	{
				$Seats = ($settings.Settings.Licenses.software | where-object {$_.name -match $software.name}).seats
				if (!$seats){$seats=99}
				foreach($computer in $software.Computers){
					$License = new-object PSObject
					write-host "Computer is:"$computer
					$comp = $computers | Where-object {$_."Asset Tag" -eq $computer.tostring()}
					$Licenses += new-object PsObject -Property @{
						"User Name" = $comp.name
						"User Email" = $comp.email
						"Username" = $comp.username
						"Software Name" = $Software.name
						"Serial" = "-"
						"Licensed to Name" = $null
						"Licensed to Email" = $null
						"Seats" = $seats
						"Reassignable" = "Yes"
						"Supplier" = $null
						"Maintained" = $null
						"Notes" = $null
						"Purchase Date" = $null
						"Computer" = $comp."Asset Tag"
					}
				}		
			
			}
		
		}

	}
	return $Licenses
}
Function Get-IPOfUser{
	param (
			[string]$user = $null,
			[string]$configFile = $null
	)
		
	if ($ConfigFile -ne "" ){
		write-host "Using config file: "$configfile
		[xml] $settings = Get-Content $ConfigFile
	}
	if ($Settings.settings.Company.OutPutFilePrefix){
		$FilterString = "*"+$settings.Settings.Company.OutPutFilePrefix+"*.csv"
	}elseif(!$Filter){
		$Filter=Read-Host 'Please enter Filter String'
	}else{
		$FilterString=$Filter
	}
	if ($settings.settings.Company.SaveLocation){
		$Directory = $settings.Settings.Company.SaveLocation
	}elseif(!$Directory){
		$Directory=Read-Host 'Please enter the directory to find computers'
	}
	if ($settings.settings.Company.SuperListName){
		$SaveName = $settings.Settings.Company.SuperListName
	}elseif(!$SaveName){
		$Savename = Read-Host 'Please the name of the existing file if it exists'
	}
	$Computers = Get-ChildItem $Directory | where-object {$_.fullname -like $FilterString -and $_.fullname -notmatch $SaveName} | sort-object LastWriteTime -desc | select -expandproperty fullname | import-csv
	$ComputersOfUser = $computers | where-object {$_.Username -eq $user}
	write-host "Computers to search for IPs:"$computersofuser."Asset Tag"
	$separator=","," "
	#$IPs = ($computersOfUser."ip address") | foreach {if($_ -ne $null){$_.split($separator)}} | where-object {$_ -ne " " -and $_ -ne ""}
	$IPs = ($computersOfUser) | foreach {if($_."ip address" -ne $null){$_."ip address".split($separator)}} | where-object {$_ -ne " " -and $_ -ne ""}
	return $IPs
	
}