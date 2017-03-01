# *******************************************************************************************************************
# *******************************************************************************************************************
# Purpose of Script:
#    Helps with the migration of VMware Hosts from one vCenter to another by
#    pulling the Annotations for a VMware Host and any VMs running on the host 
#    and writing them to a file.
#
#    Prompted inputs:  vCenterName, VMHostName
#    Outputs:          $USERPROFILE$\Documents\VMAnnotations-<VHostName>.csv
#                      $USERPROFILE$\Documents\VMHost-<VHostName>.csv
# *******************************************************************************************************************  
# Authored Date:    Oct 2014
# Original Author:  Graham Jensen
# *************************
# ===================================================================================================================
# Update Log:   Please use this section to document changes made to this script
# ===================================================================================================================
# -----------------------------------------------------------------------------
# Update <Date>
#    Author:    <Name>
#    Description of Change:
#       <Description>
# -----------------------------------------------------------------------------
# *******************************************************************************************************************

# +------------------------------------------------------+
# |        Load VMware modules if not loaded             |
# +------------------------------------------------------+
 
if ( !(Get-Module -Name VMware.VimAutomation.Core -ErrorAction SilentlyContinue) ) {
    if (Test-Path -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\VMware, Inc.\VMware vSphere PowerCLI' ) {
        $Regkey = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\VMware, Inc.\VMware vSphere PowerCLI'
       
    } else {
        $Regkey = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\VMware, Inc.\VMware vSphere PowerCLI'
    }
    . (join-path -path (Get-ItemProperty  $Regkey).InstallPath -childpath 'Scripts\Initialize-PowerCLIEnvironment.ps1')
}
#if ( !(Get-Module -Name VMware.VimAutomation.Core -ErrorAction SilentlyContinue) ) {
#    Write-Host "VMware modules not loaded/unable to load"
#    Exit 99
#}


# **** Path Variable, modify for your use ****
$CSVPath = "c:\HostUpgradeInfo\"

#Prompt User for vCenter
$VCName = Read-Host "Enter the FQHN of the vCenter that the host currently resides in"

#Prompt User for ESXi Host
$VHostName = Read-Host "Enter the FQHN of the ESXi host you are upgrading"

#Connect to VC using input from above
Connect-VIServer $VCName

# Extract Host Annotations and write them to CSV file
Get-VMHost -Name $VHostName | ForEach-Object {
$VM = $_
$VM | Get-Annotation |`
ForEach-Object {
$Report = "" | Select-Object VM,Name,Value
$Report.VM = $VM.Name
$Report.Name = $_.Name
$Report.Value = $_.Value
$Report
}
} | Export-Csv -Path $CSVPath\VMHost-$VHostName.csv -NoTypeInformation

#Extract VM Annotations and write them to CSV File
Get-VMHost -Name $VHostName | Get-VM | ForEach-Object {
$VM = $_
$VM | Get-Annotation |`
ForEach-Object {
$Report = "" | Select-Object VM,Name,Value
$Report.VM = $VM.Name
$Report.Name = $_.Name
$Report.Value = $_.Value
$Report
}
} | Export-Csv -Path $CSVPath\VMAnnotations-$VHostName.csv -NoTypeInformation

#Disconnect from VC
Disconnect-VIServer -Server $VCName -Confirm:$false

#Create Server.txt file for input to next script
Import-csv -Path $CSVPath\VMAnnotations-$VHostName.csv | Select VM | Format-Table -HideTableHeaders | Out-File $CSVPath\VMList-temp.txt 
(Get-Content $CSVPath\VMList-temp.txt)| Foreach {$_.TrimEnd()} | ? {$_.trim() -ne "" } | Sort-Object | Get-Unique |Out-File $CSVPath\server.txt
Remove-Item $CSVPath\VMList-temp.txt