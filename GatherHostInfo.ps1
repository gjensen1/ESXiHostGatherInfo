#param (
#    [switch]$WithLocalCreds = $false
#)
<# 
*******************************************************************************************************************
Authored Date:    Sept 2016
Original Author:  Graham Jensen
*******************************************************************************************************************
Purpose of Script:

   Gathers and documents important information that may be required
   during the upgrade of an ESXi Host.  This includes:
    - VMs running on the host
    - VM IPConfig info
    - Share Info
    - Print Queue Info
    - VM Configuration Info     
    - Host Annotations
    - VM Annotations

   Prompted inputs:  vCenterName, VMHostName

   Outputs:          
            $USERPROFILE$\Documents\HostUpgradeInfo\$VMHost\server.txt
            $USERPROFILE$\Documents\HostUpgradeInfo\$VMHost\$HostName.docx
            $USERPROFILE$\Documents\HostUpgradeInfo\$VMHost\IPConfig\$VMName-ipconfig.txt [Multiple Files]
            $USERPROFILE$\Documents\HostUpgradeInfo\$VMHost\PrinterInfo\$VMName-PrinterInfo.txt [Multiple Files]
            $USERPROFILE$\Documents\HostUpgradeInfo\$VMHost\ShareInfo\$VMName-sharelist.txt [as well as .reg files]
            $USERPROFILE$\Documents\HostUpgradeInfo\$VMHost\VMInfo\$VMName-VMInfo.txt [Multiple Files]
            $USERPROFILE$\Documents\HostUpgradeInfo\$VMHost\Annotations\VMAnnotations-<HostName>.csv
            $USERPROFILE$\Documents\HostUpgradeInfo\$VMHost\Annotations\VMHost-<HostName>.csv
            $USERPROFILE$\Documents\HostUpgradeInfo\GatherHostInfoLog.txt
*******************************************************************************************************************  
Prerequisites:

    #1  This script uses the VMware modules installed by the installation of VMware PowerCLI
        ENSURE that VMware PowerCLI has been installed.  
    
        Installation media can be found here: 
        \\cihs.ad.gov.on.ca\tbs\Groups\ITS\DCO\RHS\RHS\Software\VMware

    #2  To complete necessary tasks this script will require C3 account priviledges
        you will be prompted for your C3 account and password.  The Get-Credential method
        is used for this, so credentials are maintained securely.

===================================================================================================================
Update Log:   Please use this section to document changes made to this script
===================================================================================================================
-----------------------------------------------------------------------------
Update Feb 7th, 2017
   Author:    Graham Jensen
   Description of Change:
      Add ability to prompt for a second set of credentials for connecting
      to VMs to gather IPConfig info, when that VM requires local credentials
      to attain administrator rights.

      This is implemented as a switch to the script.  To invoke run the sciprt
      as follows:
      
      GatherHostInfo.ps1 -WithLocalCreds
-----------------------------------------------------------------------------
-----------------------------------------------------------------------------
Update March 7th, 2017
   Author:    Graham Jensen
   Description of Change:
      Add additional information gathering
      - If File Server, collects Share Listing and exports LanmanServer
        Registry entries
      - Detailed VM Virtual Hardware Configuration
      - If Print Server, collects PrintQueue and Driver Info
      - Code is present to enable backup of PrintQueues with PrintBRM
        however still having some issues with PSRemoting enablement so 
        currently this functionality is disabled.
      - WithLocalCreds switch no longer required for Courts vm enumeration
        script now automatically detects failed attempt and then prompts 
        for additional creds if required.
-----------------------------------------------------------------------------
-----------------------------------------------------------------------------
Update March 21st, 2017
   Author:    Graham Jensen
   Description of Change:
    - Modified the GetVMInfo function to remove dependency on depricated
      Get-VM functionality for Disk Info and NetworkAdapter Info.  Changed
      code to use new Get-Harddisk, and Get-NetworkAdapter PowerCLI cmdlets
    - Add functionality to check for VM PowerState and GuestID.  If VM is not
      powered on logs state and discontinues additional checks for that VM.  
      If GuestID is not 'Like' "Win*" then log and discontinue additional
      checks for that VM
    - Cleaned up some the Host and log output
-----------------------------------------------------------------------------
-----------------------------------------------------------------------------
Update <Date>
   Author:    <Name>
   Description of Change:
      <Description>
-----------------------------------------------------------------------------
*******************************************************************************************************************
#>

# +------------------------------------------------------+
# |        Load VMware modules if not loaded             |
# +------------------------------------------------------+
"Loading VMWare Modules"
$ErrorActionPreference="SilentlyContinue" 
if ( !(Get-Module -Name VMware.VimAutomation.Core -ErrorAction SilentlyContinue) ) {
    if (Test-Path -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\VMware, Inc.\VMware vSphere PowerCLI' ) {
        $Regkey = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\VMware, Inc.\VMware vSphere PowerCLI'
       
    } else {
        $Regkey = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\VMware, Inc.\VMware vSphere PowerCLI'
    }
    . (join-path -path (Get-ItemProperty  $Regkey).InstallPath -childpath 'Scripts\Initialize-PowerCLIEnvironment.ps1')
}
$ErrorActionPreference="Continue"

#*************************
# Load BitsTransfer Module
#*************************
Import-Module BitsTransfer


# -----------------------
# Define Global Variables
# -----------------------
$Global:Folder = $env:USERPROFILE+"\Documents\HostUpgradeInfo" 
$Global:WorkFolder = $null
$Global:VCName = $null
$Global:HostName = $null
$Global:RunDate = Get-Date
$Global:RunAgain = $null
$Global:Creds = $null
$Global:CredsLocal = $null
$Global:FileServer = $null
$Global:PrintServer = $null

#*****************
# Get VC from User
#*****************
Function Get-VCenter {
    #Prompt User for vCenter
    Write-Host "Enter the FQHN of the vCenter that the host currently resides in: " -ForegroundColor "Yellow" -NoNewline
    $Global:VCName = Read-Host 
}
#*******************
# EndFunction Get-VC
#*******************

#*************
# Get HostName
#*************
Function Get-HostName {
    #Prompt User for ESXi Host
    Write-Host "Enter the FQHN of the ESXi Host you want to collect data from: " -ForegroundColor "Yellow" -NoNewLine
    $Global:HostName = Read-Host
}
#*************************
# EndFunction Get-HostName
#*************************


#*************************************************
# Check for Folder Structure if not present create
#*************************************************
Function Verify-Folders {
    "Building Local folder structure" 
    If (!(Test-Path $Global:WorkFolder)) {
        New-Item $Global:WorkFolder -type Directory
        New-Item "$Global:WorkFolder\Annotations" -type Directory
        New-Item "$Global:WorkFolder\IPConfig" -type Directory
        New-Item "$Global:WorkFolder\ShareInfo" -type Directory
        New-Item "$Global:WorkFolder\VMInfo" -type Directory
        New-Item "$Global:WorkFolder\PrinterInfo" -type Directory
        }
    "Folder Structure built" 
}
#***************************
# EndFunction Verify-Folders
#***************************

#*******************
# Connect to vCenter
#*******************
Function Connect-VC {
    "Connecting to $Global:VCName"
    Connect-VIServer $Global:VCName -Credential $Global:Creds > $null
}
#***********************
# EndFunction Connect-VC
#***********************

#*******************
# Disconnect vCenter
#*******************
Function Disconnect-VC {
    "Disconnecting $Global:VCName"
    Disconnect-VIServer -Server $Global:VCName -Confirm:$false
}
#**************************
# EndFunction Disconnect-VC
#**************************


#*********************
# Get Host Information
#*********************
Function Get-HostInfo {

    # Extract Host Annotations and write them to CSV file
    "Writing Host Annotations to VMHost-$Global:HostName.csv" 
    Get-VMHost -Name $Global:HostName | ForEach-Object {
        $VM = $_
        $VM | Get-Annotation |`
        ForEach-Object {
            $Report = "" | Select-Object VM,Name,Value
            $Report.VM = $VM.Name
            $Report.Name = $_.Name
            $Report.Value = $_.Value
            $Report
            }
    } | Export-Csv -Path $Global:WorkFolder\Annotations\VMHost-$Global:HostName.csv -NoTypeInformation

    #Extract VM Annotations and write them to CSV File
    "Writing VM Annotations to VMAnnotations-$Global:HostName.csv"
    Get-VMHost -Name $Global:HostName | Get-VM | ForEach-Object {
    $VM = $_
    $VM | Get-Annotation |`
    ForEach-Object {
        $Report = "" | Select-Object VM,Name,Value
        $Report.VM = $VM.Name
        $Report.Name = $_.Name
        $Report.Value = $_.Value
        $Report
        }
    } | Export-Csv -Path $Global:WorkFolder\Annotations\VMAnnotations-$Global:HostName.csv -NoTypeInformation

    #Create Server.txt file for input to next script
    Import-csv -Path $Global:WorkFolder\Annotations\VMAnnotations-$Global:HostName.csv | Select VM | Format-Table -HideTableHeaders | Out-File $Global:WorkFolder\VMList-temp.txt 
    (Get-Content $Global:WorkFolder\VMList-temp.txt)| Foreach {$_.TrimEnd()} | ? {$_.trim() -ne "" } | Sort-Object | Get-Unique |Out-File $Global:WorkFolder\server.txt
        Remove-Item $Global:WorkFolder\VMList-temp.txt
}
#*************************
# EndFunction Get-HostInfo
#*************************

#***********************
# Determine Server Roles
#***********************
Function DetermineServerRoles($s) {
    $Global:FileServer = $null
    $Global:PrintServer = $null
    "Checking for server roles on $s"
    $shrQuery ="select Name, Path, Description from Win32_share where NAME != 'print$' AND Name != 'prnproc$' AND Type =0"
    $shr=get-wmiobject  -computer $s -query $shrQuery -Credential $Global:CredsLocal  -ErrorAction SilentlyContinue
    $numshr = $shr.length

    if($numshr -gt 0){
        $Global:FileServer = $True
        }
        
    $prn = @(get-wmiobject win32_printer -computer $s -Credential $Global:CredsLocal  -ErrorAction SilentlyContinue | where {$_.Shared -eq $TRUE})
    $numprn = $prn.length

    if($numprn -gt 0){
        $Global:PrintServer = $True    
        }

}
#*********************************
# EndFunction DetermineServerRoles
#*********************************

#*****************
# Get VM IPConfigs
#*****************
Function Get-IPConfigs($s) {
    $SavePath = "$Global:WorkFolder\IPConfig"
    $ErrorActionPreference="SilentlyContinue"
    "Running IPConfig on $s"
    $result = invoke-wmimethod -computer $s -path Win32_process -name Create -ArgumentList "cmd /c ipconfig /all > c:\temp\$s-ipconfig.txt" -Credential $Global:CredsLocal 
    switch ($result.returnvalue) {
        0 {"$s Successful Completion."}
        2 {"$s Access Denied."}
        3 {"$s Insufficient Privilege."}
        8 {"$s Unknown failure."}
        9 {"$s Path Not Found."}
        21 {"$s Invalid Parameter."}
        default {"$s Could not be determined."}
        }
    sleep 2
    if ($result.returnvalue -ne 0){
        Write-Host "Was not able to connect to $s" -ForegroundColor Red
        Write-Host "Please supply alternate credentials!!!" -ForegroundColor Red
        $Global:CredsLocal = Get-Credential -Credential $null

        #Retry with new credentials
        $result = invoke-wmimethod -computer $s -path Win32_process -name Create -ArgumentList "cmd /c ipconfig /all > c:\temp\$s-ipconfig.txt" -Credential $Global:CredsLocal
        switch ($result.returnvalue) {
            0 {"$s Successful Completion."}
            2 {"$s Access Denied."}
            3 {"$s Insufficient Privilege."}
            8 {"$s Unknown failure."}
            9 {"$s Path Not Found."}
            21 {"$s Invalid Parameter."}
            default {"$s Could not be determined."}
            }
    
        }
    if ($result.returnvalue -eq 0) {
        "Connecting to C$ on $s"
        New-PSDrive REMOTE -PSProvider FileSystem -Root \\$s\c$\temp -Credential $Global:CredsLocal > $null 
        "Moving $s-ipconfig.txt to local machine"
        move-item -path REMOTE:\$s-ipconfig.txt $SavePath -force
        "Disconnect $s" 
        Remove-PSDrive REMOTE
        net use \\$s\c$\temp /d > $null
        }
        Else {
            "Not able to connect to $s to retrieve IPConfig"
            "You will need to manually gather this information"
            Read-Host -Prompt "Press <Enter> to continue" 
        }
    $ErrorActionPreference="Continue"
}
#**************************
# EndFunction Get-IPConfigs
#**************************

#*************************
# Get VM Share Information
#*************************
Function Get-ShareInfo($s) {
    $SavePath = "$Global:WorkFolder\ShareInfo"
 
    "Gathering Share Info on $s"
    $result = Invoke-WmiMethod -computer $s -path Win32_Process -name Create -ArgumentList "cmd /c net share > c:\temp\$s-sharelist.txt" -Credential $Global:CredsLocal
    switch ($result.returnvalue) {
        0 {"$s Successful Completion."}
        2 {"$s Access Denied."}
        3 {"$s Insufficient Privilege."}
        8 {"$s Unknown failure."}
        9 {"$s Path Not Found."}
        21 {"$s Invalid Parameter."}
        default {"$s Could not be determined."}
        }
    "Extract Shares from registry on $s"
    $result = Invoke-WmiMethod -computer $s -path Win32_Process -name Create -ArgumentList "cmd /c reg export HKLM\SYSTEM\CurrentControlSet\Services\LanmanServer\Shares c:\temp\$s-shares.reg" -Credential $Global:CredsLocal
    switch ($result.returnvalue) {
        0 {"$s Successful Completion."}
        2 {"$s Access Denied."}
        3 {"$s Insufficient Privilege."}
        8 {"$s Unknown failure."}
        9 {"$s Path Not Found."}
        21 {"$s Invalid Parameter."}
        default {"$s Could not be determined."}
        }
    "Extract Share Security from registry on $s"
    $result = Invoke-WmiMethod -computer $s -path Win32_Process -name Create -ArgumentList "cmd /c reg export HKLM\SYSTEM\CurrentControlSet\Services\LanmanServer\Shares\Security c:\temp\$s-security.reg" -Credential $Global:CredsLocal
    switch ($result.returnvalue) {
        0 {"$s Successful Completion."}
        2 {"$s Access Denied."}
        3 {"$s Insufficient Privilege."}
        8 {"$s Unknown failure."}
        9 {"$s Path Not Found."}
        21 {"$s Invalid Parameter."}
        default {"$s Could not be determined."}
        }
    sleep 3
    if ($result.returnvalue -eq 0) {
        "Connecting to C$ on $s"
        New-PSDrive REMOTE -PSProvider FileSystem -Root \\$s\c$\temp -Credential $Global:CredsLocal > $null 
        "Moving $s-sharelist.txt to local machine"
        move-item -path REMOTE:\$s-sharelist.txt $SavePath -force
        "Moving $s Registry files to local machine"
        move-item -path REMOTE:\$s-*.reg $SavePath -force
        "Disconnect $s" 
        Remove-PSDrive REMOTE
        net use \\$s\c$\temp /d > $null
        }
}
#**************************
# EndFunction Get-ShareInfo
#**************************

#****************************
# Check for PSRemoting Status
#****************************
Function PSRemotingStatus ($s) {
    $PSRemotingEnabled = [bool](Test-WSMan -Computer $s -ErrorAction SilentlyContinue)

    If ($PSRemotingEnabled -eq $false){
        "PSRemoting being enabled on $s"
        $username = $Global:CredsLocal.Username
        #unencrypting $credsLocal.Password so that we can send it to PSExec
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Global:CredsLocal.Password)
        $pass = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
        .\PSExec.exe \\$s -u $username -p $pass powershell enable-psremoting -force
        #Clear unencrypted password from memory
        $pass = $null
        $PSRemotingEnabled = [bool](Test-WSMan -Computer $s -ErrorAction SilentlyContinue)
        If ($PSRemotingEnabled -eq $false){
            "PSRemoting was not successfully enabled on $s"
            $returnValue = $False
            }
            Else {
               "PSRemoting is now enabled on $s"
               $returnValue = $True
               } 
        }
    Else {
        "PSRemoting is enabled on $s"
        $returnValue = $True
    }
    return ,$returnValue
}
#*****************************
# EndFunction PSRemotingStatus
#*****************************

#*********************************
# Backup PrintQueues with Printbrm
#*********************************
Function Get-PrintQueues($s){
    $SavePath = "$Global:WorkFolder\PrinterInfo"

    "Gathering Print Queue Info on $s"
    $driverQuery ="select Name, Version, SupportedPlatform, DriverPath, OEMUrl from Win32_printerdriver"
    $prnQuery ="select * from Win32_printer where NAME != 'Microsoft XPS Document Writer'";
    $prn = get-wmiobject -computer $s  -query $prnQuery -Credential $Global:CredsLocal

    "$s PrinterInfo" | Out-File -FilePath $SavePath\$s-PrinterInfo.txt
    "===========================" | Out-File -FilePath $SavePath\$s-PrinterInfo.txt -append

    foreach($printer in $prn){
        "ShareName:`t" + $printer.ShareName | Out-File -FilePath $SavePath\$s-PrinterInfo.txt -append
        "DriverName:`t" + $printer.driverName | Out-File -FilePath $SavePath\$s-PrinterInfo.txt -append
        "Location:`t" + $printer.location | Out-File -FilePath $SavePath\$s-PrinterInfo.txt -append
        "Description:`t" + $printer.description | Out-File -FilePath $SavePath\$s-PrinterInfo.txt -append
        #$printer.SystemName
        $port = $printer.PortName
        $ipTCPport =get-wmiobject -class win32_tcpIPprinterPort -computer $server -Credential $Global:CredsLocal | where-object{$_.Name -eq $port}
        "PrinterIP:`t" + $ipTCPport.HostAddress | Out-File -FilePath $SavePath\$s-PrinterInfo.txt -append
        "-----------------------------" | Out-File -FilePath $SavePath\$s-PrinterInfo.txt -append
        }
    
    " " | Out-File -FilePath $SavePath\$s-PrinterInfo.txt -append
    "Print Drivers" | Out-File -FilePath $SavePath\$s-PrinterInfo.txt -append
    "=============" | Out-File -FilePath $SavePath\$s-PrinterInfo.txt -append    
    $result = get-wmiobject  -computer $s  -query $DriverQuery -Credential $Global:CredsLocal
    foreach($config in $result){
        $Dname = $config.Name
        $Durl  = $config.OEMURL
        $s+","+ $Dname+","+$Durl | Out-File -FilePath $SavePath\$s-PrinterInfo.txt -append   
        }

 <#   "Backing up Print Queues on $s"
    "-----------------------------"
    If ((PSRemotingStatus $s) -eq $True){
        $command = { C:\Windows\System32\spool\tools\printbrm -B -F c:\temp\prnbackup.printerexport -o force  }
        "Running printbrm on $s"
        Invoke-Command -Computer $s -ScriptBlock $command -Credential $Global:CredsLocal
        "Getting Printer backup file from $s to $SavePath\$s-prnbackup.printerexport"
        New-PSDrive REMOTE -PSProvider FileSystem -Root \\$s\c$\temp -Credential $Global:CredsLocal > $null
        Start-BitsTransfer -Source REMOTE:\prnbackup.printerexport -Destination $SavePath\$s-prnbackup.printerexport -Description "Transfer PrinterBackup file from $s" -DisplayName "PrinterBackup" -Credential $Global:CredsLocal
        "Cleaning up Temp drive on $s"
        Remove-item -path REMOTE:\prnbackup.printerexport -force
        Remove-PSDrive REMOTE
        net use \\$s\c$\temp /d        
        }
        Else {
            "Unable to run PrintBRM on $server"  
            "If Print Queue backup is required" 
            "it will have to be done manually"
            Read-Host -Prompt "Press <Enter> to continue" 
            } #>
}
#****************************
# EndFunction Get-PrintQueues
#****************************

#*********************
# Gather VM Guest Info
#*********************
Function GetVMInfo($s) {
    #"Gathering VM Configuration Info"
    $SavePath = "$Global:WorkFolder\VMInfo"
    #Reset Global Switches
    $Global:PoweredOn = $Null
    $Global:IsWindows = $Null

    "Getting VM Config Info for $s"
    $VMInfo = Get-VM $s
    $VMGuest = Get-Vmguest $s
    $VMScsi = Get-ScsiController $s
    $VMDisk = Get-HardDisk $s
    $VMNetwork = Get-NetworkAdapter $s

    If ($VMinfo.PowerState -eq "PoweredOn"){
        $Global:PoweredOn = $True
        }

    If ($VMInfo.GuestId -like "Win*"){
        $Global:IsWindows = $True
        }

    #Write Info to Text Files
    $VMInfo.Name | Out-File -FilePath $Global:WorkFolder\VMInfo\$s-VMInfo.txt
    "`n===============`n" | Out-File -FilePath $Global:WorkFolder\VMInfo\$s-VMInfo.txt -append
    "`nVM Id:`t" + $VMInfo.ID | Out-File -FilePath $Global:WorkFolder\VMInfo\$s-VMInfo.txt -append
    "Folder: `t" + $VMInfo.Folder | Out-File -FilePath $Global:WorkFolder\VMInfo\$s-VMInfo.txt -append
    "PowerState:`t" + $VMInfo.PowerState | Out-File -FilePath $Global:WorkFolder\VMInfo\$s-VMInfo.txt -append
    "Guest OS:`t" + $VMGuest.OSFullName | Out-File -FilePath $Global:WorkFolder\VMInfo\$s-VMInfo.txt -append
    "Guest IP:`t" + $VMGuest.IPAddress | Out-File -FilePath $Global:WorkFolder\VMInfo\$s-VMInfo.txt -append
    "CPU(s):`t" + $VMInfo.NumCPU | Out-File -FilePath $Global:WorkFolder\VMInfo\$s-VMInfo.txt -append
    "Ram (MB):`t" + $VMInfo.MemoryMB | Out-File -FilePath $Global:WorkFolder\VMInfo\$s-VMInfo.txt -append
    " " | Out-File -FilePath $Global:WorkFolder\VMInfo\$s-VMInfo.txt -append
    "SCSI Adapters" | Out-File -FilePath $Global:WorkFolder\VMInfo\$s-VMInfo.txt -append
    "-------------" | Out-File -FilePath $Global:WorkFolder\VMInfo\$s-VMInfo.txt -append
    $VMScsi | Format-Table Name,Type,UnitNumber -AutoSize |Out-File -FilePath $Global:WorkFolder\VMInfo\$s-VMInfo.txt -append
    " " | Out-File -FilePath $Global:WorkFolder\VMInfo\$s-VMInfo.txt -append
    "Disk Configuration" | Out-File -FilePath $Global:WorkFolder\VMInfo\$s-VMInfo.txt -append
    "------------------" | Out-File -FilePath $Global:WorkFolder\VMInfo\$s-VMInfo.txt -append
    $VMDisk | Format-Table Name,CapacityGB,CapacityKB,Filename -autosize -wrap | Out-File -FilePath $Global:WorkFolder\VMInfo\$s-VMInfo.txt -append
    " " | Out-File -FilePath $Global:WorkFolder\VMInfo\$s-VMInfo.txt -append
    "Network Configuration" | Out-File -FilePath $Global:WorkFolder\VMInfo\$s-VMInfo.txt -append
    "---------------------" | Out-File -FilePath $Global:WorkFolder\VMInfo\$s-VMInfo.txt -append
    $VMNetwork | Format-Table Name,Type,NetworkName,MacAddress -AutoSize | Out-File -FilePath $Global:WorkFolder\VMInfo\$s-VMInfo.txt -append
}
#**********************
# EndFunction GetVMInfo
#**********************

#********************
# Build Word Document
#********************
Function Build-Word {
    "Building Word Document"
    $wdOrientLandscape = 1
    $wdOrientPortrait = 0

    $Word = New-Object -ComObject Word.Application
    $Word.Visible = $True
    $Document = $Word.Documents.Add()
    $Selection = $Word.Selection
    $Range = $Document.Range()
    #$Document.PageSetup.Orientation = $wdOrientLandscape

    #Get HostName and SiteAddress from VMAnnotations File and add to Word Doc
    $SiteInfo = Import-Csv $Global:WorkFolder\Annotations\VMHost-$Global:HostName.csv
    $SiteAddress = $SiteInfo | where {$_.Name -eq "Site Address"} | Select Value | Format-Table -HideTableHeaders | Out-String
    $Global:HostName = $SiteInfo | where {$_.Name -eq "Site Address"} | Select VM | Format-Table -HideTableHeaders | Out-String
    $SiteAddress = $SiteAddress.Trim()
    $Global:HostName = $Global:HostName.Trim()

    #Build Word Doc
    #Title Page
    $Selection.Style="Intense Quote"
    $Selection.TypeText("Gather Information for ESXi Host Upgrades ")
    $Selection.TypeText("$Global:RunDate")
    $Selection.TypeParagraph()
    $Selection.Style="Intense Reference"
    $Selection.TypeText("Site: $SiteAddress")
    $Selection.TypeParagraph()
    $Selection.Style="Intense Reference"
    $Selection.TypeText("Host: $Global:HostName")
    $Selection.TypeParagraph()
    $Selection.Style="Strong"
    $Selection.TypeText("VMs Present on Host:")
    $Selection.TypeParagraph()
    $Selection.InsertFile("$Global:WorkFolder\server.txt")
    $Selection.InsertNewPage()
    
    #Add VMInfo Files
    $Selection.Style="Strong"
    $Selection.TypeText("VM Configurations")
    $Selection.TypeParagraph()
    Get-ChildItem "$Global:WorkFolder\VMInfo" |
    ForEach-Object{
        $Selection.InsertFile("$Global:WorkFolder\VMInfo\$_")
        $Selection.TypeParagraph()
        $Selection.TypeText("-------------------")
        $Selection.InsertNewPage()  
        }
    
    #Add IPConfig Files
    $Selection.Style="Strong"
    $Selection.TypeText("IP Configurations")
    $Selection.TypeParagraph()
    Get-ChildItem "$Global:WorkFolder\IPConfig" |
    ForEach-Object{
        $Selection.InsertFile("$Global:WorkFolder\IPConfig\$_")
        $Selection.TypeParagraph()
        $Selection.TypeText("-------------------")
        $Selection.InsertNewPage()  
        }

    #Add ShareListings
    $Selection.Style="Strong"
    $Selection.TypeText("Share Listings")
    $Selection.TypeParagraph()
    Get-ChildItem "$Global:WorkFolder\ShareInfo" |
    ForEach-Object{
        If ($_ -like "*.txt"){
        $Selection.TypeText("$_")
        $Selection.InsertFile("$Global:WorkFolder\ShareInfo\$_")
        $Selection.TypeParagraph()
        $Selection.TypeText("-------------------")
        $Selection.InsertNewPage()       
        }
    }

    #Add PrinterListings
    $Selection.Style="Strong"
    $Selection.TypeText("Printer Info")
    $Selection.TypeParagraph()
    Get-ChildItem "$Global:WorkFolder\PrinterInfo" |
    ForEach-Object{
        If ($_ -like "*.txt"){
        $Selection.TypeText("$_")
        $Selection.TypeParagraph()
        $Selection.InsertFile("$Global:WorkFolder\PrinterInfo\$_")
        $Selection.TypeParagraph()
        $Selection.TypeText("-------------------")
        $Selection.InsertNewPage()       
        }
    }

    #Add Annotation Files
    $Selection.Style="Strong"
    $Selection.TypeText("Host Annotations")
    $Selection.TypeParagraph()
    $Selection.InsertFile("$Global:WorkFolder\Annotations\VMHost-$Global:HostName.csv")
    $Selection.InsertNewPage()
    $Selection.Style="Strong"
    $Selection.TypeText("VM Annotations")
    $Selection.TypeParagraph()
    $Selection.InsertFile("$Global:WorkFolder\Annotations\VMAnnotations-$Global:HostName.csv")

    #Return to top
    $FileName = "$Global:WorkFolder\$Global:HostName.docx"
    $Document.SaveAs([ref] $FileName)
    $Document.Close()
    $Word.Quit()
    "Word Doc built and saved to $FileName"
    "********************************"
}
#**********************
# EndFuntion Build-Word
#**********************


#*************************
# Prompt for Running Again
#*************************
Function Run-Again {
    Write-Host "Do you have another Host to collect " -ForeGroundColor "Yellow" -NoNewLine
    Write-Host "(y/n)" -ForeGroundColor "Red" -NoNewLine
    Write-Host ": " -ForeGroundColor "Yellow" -NoNewLine
    $Global:RunAgain = Read-Host 
    $Global:RunAgain = $Global:RunAgain.Substring(0,1).ToLower()

}

#*********************
# Clean Up after Run
#*********************
Function Clean-Up {
    $Global:Folder = $null
    $Global:WorkFolder = $null
    $Global:VCName = $null
    $Global:HostName = $null
    $Global:RunDate = $null
    $Global:Creds = $null
    $Global:CredsLocal = $null
}

#***************
# Execute Script
#***************
CLS
$ErrorActionPreference="SilentlyContinue"
Stop-Transcript | out-null
$ErrorActionPreference="Continue"
Start-Transcript -path $Global:Folder\GatherHostInfoLog.txt
"================================================="
" "
Write-Host "Get CIHS credentials" -ForegroundColor Yellow

$Global:Creds = Get-Credential -Credential $null
$Global:CredsLocal = $Global:Creds
CLS
Get-VCenter
Connect-VC
"-------------------------------------------------"
$Global:RunAgain = "y"
Do {
    Get-HostName
    CLS
    $Global:WorkFolder = "$Global:Folder\$Global:HostName"
    Verify-Folders
    "-------------------------------------------------"  
    Get-HostInfo
    "-------------------------------------------------"
    $servers = Get-Content "$Global:WorkFolder\server.txt"
    forEach ($server in $servers) {
        GetVMInfo $server

        If ($Global:PoweredOn -eq $True){
            If ($Global:IsWindows -eq $True){
                Get-IPConfigs $server
                DetermineServerRoles $server
                If ($Global:FileServer -eq $True){
                    "$server is a File Server - Gathering Share Info"
                    Get-ShareInfo $server
                    }
                If ($Global:PrintServer -eq $True){
                    "$Server is a Print Server - Backing up Print Queues"
                    Get-PrintQueues $server
                    }
                }
                Else {
                    Write-Host "VM $Server is not Windows, skipping additional steps" -ForegroundColor Yellow
                    }

            }
            Else {
                Write-Host "VM $Server is not Powered On skipping additional steps " -ForegroundColor Yellow
            }
        "-------------------------------------------------"
        }
    Build-Word
    Run-Again
    } While ($Global:RunAgain -eq "y")
Disconnect-VC
"Open Explorer to $Global:Folder"
Invoke-Item $Global:Folder
Clean-Up
Stop-Transcript
