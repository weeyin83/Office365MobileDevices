<#
.SYNOPSIS
MobileDevices.ps1 - Collect Office 365 Mobile Device Information. 

.DESCRIPTION 
This script will connect to Office 365 and export a list of users that have a mobile device attached to their account. 

.OUTPUTS
Creates and CSV file with the results, the file will be called mobiledevices.csv and stored on the C drive. 

.NOTES
Written by: Sarah Lean

Find me on:

* My Blog:	http://www.techielass.com
* Twitter:	https://twitter.com/techielass
* LinkedIn:	http://uk.linkedin.com/in/sazlean


.EXAMPLE
.\MobileDevices.ps1 
Runs and exports the information

Change Log
V1.00, 10/04/2017 - Initial version

License:

The MIT License (MIT)

Copyright (c) 2017 Sarah Lean

Permission is hereby granted, free of charge, to any person obtaining a copy 
of this software and associated documentation files (the "Software"), to deal 
in the Software without restriction, including without limitation the rights 
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell 
copies of the Software, and to permit persons to whom the Software is 
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all 
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING 
FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER 
DEALINGS IN THE SOFTWARE.

#>

##Connect to Exchange Online
$UserCredential = Get-Credential

##Create session into the Office 365 admin console
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection

##Import the Powershell Session
Import-PSSession $Session -DisableNameChecking > $null

#Output file
$FileOutput = "C:\MobileDevices.csv"

##Set Headers for Output file
Out-File -FilePath $FileOutput -InputObject "DisplayName,UPN,DeviceMobileOperator,DeviceType,DeviceUserAgent,DeviceModel,DeviceFriendlyName,DeviceOS,DeviceOSLanguage,IsRemoteWipeSupported,DevicePolicyApplied,DevicePolicyApplicationStatus,FirstSyncTime,LastPolicyUpdateTime,LastSyncAttemptTime,LastSuccessSync" -Encoding UTF8

##Collect a list of mailboxes within the Office 365 tenant
$mailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited -Filter {HiddenFromAddressListsEnabled -eq $false}

##Loop through each mailbox 
foreach ($mailbox in $mailboxes) { 
	$devices = Get-MobileDeviceStatistics -Mailbox $mailbox.samaccountname 
	 
	##If the mailbox has a device loop through it and collect information
	if ($devices) { 
		foreach ($device in $devices){ 
		  
			##Create a new object and output the information to CSV
			$deviceinf = New-Object -TypeName psobject 
			$deviceinf | Add-Member -Name DisplayName -Value $mailbox.DisplayName -MemberType NoteProperty 
			$deviceinf | Add-Member -Name UPN -Value $mailbox.UserPrincipalName -MemberType NoteProperty 
			$deviceinf | Add-Member -Name DeviceMobileOperator -Value $device.DeviceMobileOperator -MemberType NoteProperty 
			$deviceinf | Add-Member -Name DeviceType -Value $device.DeviceType -MemberType NoteProperty 
			$deviceinf | Add-Member -Name DeviceUserAgent -Value $device.DeviceUserAgent -MemberType NoteProperty 
			$deviceinf | Add-Member -Name DeviceModel -Value $device.DeviceModel -MemberType NoteProperty 
			$deviceinf | Add-Member -Name DeviceFriendlyName -Value $device.DeviceFriendlyName -MemberType NoteProperty 
			$deviceinf | Add-Member -Name DeviceOS -Value $device.DeviceOS -MemberType NoteProperty 
			$deviceinf | Add-Member -Name DeviceOSLanguage -Value $device.DeviceOSLanguage -MemberType NoteProperty 
			$deviceinf | Add-Member -Name IsRemoteWipeSupported -Value $device.IsRemoteWipeSupported -MemberType NoteProperty 
			$deviceinf | Add-Member -Name DeviceWipeSentTime -Value $device.DeviceWipeSentTime -MemberType NoteProperty 
			$deviceinf | Add-Member -Name DeviceWipeRequestTime -Value $device.DeviceWipeRequestTime -MemberType NoteProperty 
			$deviceinf | Add-Member -Name DeviceWipeAckTime -Value $device.DeviceWipeAckTime -MemberType NoteProperty 
			$deviceinf | Add-Member -Name LastDeviceWipeRequestor -Value $device.LastDeviceWipeRequestor -MemberType NoteProperty 
			$deviceinf | Add-Member -Name DevicePolicyApplied -Value $device.DevicePolicyApplied -MemberType NoteProperty 
			$deviceinf | Add-Member -Name DevicePolicyApplicationStatus -Value $device.DevicePolicyApplicationStatus -MemberType NoteProperty 
			$deviceinf | Add-Member -Name FirstSyncTime -Value ($device.FirstSyncTime) -MemberType NoteProperty 
			$deviceinf | Add-Member -Name LastPolicyUpdateTime -Value ($device.LastPolicyUpdateTime) -MemberType NoteProperty 
			$deviceinf | Add-Member -Name LastSyncAttemptTime -Value ($device.LastSyncAttemptTime) -MemberType NoteProperty 
			$deviceinf | Add-Member -Name LastSuccessSync -Value ($device.LastSuccessSync)-MemberType NoteProperty 

            ##Output the above results to a CSV file, appending each time
			Out-File -FilePath $FileOutput -InputObject "$($deviceinf.DisplayName),$($deviceinf.UPN),$($deviceinf.DeviceMobileOperator),$($deviceinf.DeviceType),$($deviceinf.DeviceUserAgent),$($deviceinf.DeviceModel),$($deviceinf.DeviceFriendlyName),$($deviceinf.DeviceOS),$($deviceinf.DeviceOSLanguage),$($deviceinf.IsRemoteWipeSupported),$($deviceinf.DevicePolicyApplied),$($deviceinf.DevicePolicyApplicationStatus),$($deviceinf.FirstSyncTime),$($deviceinf.LastPolicyUpdateTime),$($deviceinf.LastSyncAttemptTime),$($deviceinf.LastSuccessSync)" -Encoding UTF8 -append
			
		} 
	 
	} 
 
}
##Close Powershell session
Remove-PSSession $session

##Remove Variables used
Remove-Variable -Name mailboxes
Remove-Variable -Name devices
Remove-Variable -Name fileoutput
Remove-Variable -Name session
Remove-Variable -Name deviceinf
Remove-Variable -Name mailbox
Remove-Variable -Name device

##Display completion result
Write-Host "The results of this PowerShell script have been stored in "C:\MobileDevices.csv"" -BackgroundColor Black -ForegroundColor Yellow
