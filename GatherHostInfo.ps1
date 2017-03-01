param (
    [switch]$WithLocalCreds = $false
)
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
    - Host Annotations
    - VM Annotations

   Prompted inputs:  CRQ, vCenterName, VMHostName

   Outputs:          
            $USERPROFILE$\Documents\HostUpgradeInfo\$VMHost\server.txt
            $USERPROFILE$\Documents\HostUpgradeInfo\$VMHost\$HostName.docx
            $USERPROFILE$\Documents\HostUpgradeInfo\$VMHost\IPConfig\$VMName-ipconfig.txt [Multiple Files]
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
        }
    " "
    "Folder Structure built" 
}
#***************************
# EndFunction Verify-Folders
#***************************

#*********************
# Get Host Information
#*********************
Function Get-HostInfo {
    #Connect to VC using input from above
    "-----------------------------------------"
    "Connecting to $Global:VCName"
    Connect-VIServer $Global:VCName -Credential $Global:Creds > $null

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

    #Disconnect from VC
    "Disconnecting $Global:VCName"
    "-----------------------------------------"
    Disconnect-VIServer -Server $Global:VCName -Confirm:$false

    #Create Server.txt file for input to next script
    Import-csv -Path $Global:WorkFolder\Annotations\VMAnnotations-$Global:HostName.csv | Select VM | Format-Table -HideTableHeaders | Out-File $Global:WorkFolder\VMList-temp.txt 
    (Get-Content $Global:WorkFolder\VMList-temp.txt)| Foreach {$_.TrimEnd()} | ? {$_.trim() -ne "" } | Sort-Object | Get-Unique |Out-File $Global:WorkFolder\server.txt
        Remove-Item $Global:WorkFolder\VMList-temp.txt
}
#*************************
# EndFunction Get-HostInfo
#*************************

#*****************
# Get VM IPConfigs
#*****************
Function Get-IPConfigs {
    $SavePath = "$Global:WorkFolder\IPConfig"
    $servers = Get-Content "$Global:WorkFolder\server.txt"

    forEach ($s in $servers) {
        "Running IPConfig on $s"
        $result = invoke-wmimethod -computer $s -path Win32_process -name Create -ArgumentList "cmd /c ipconfig /all > c:\temp\$s-ipconfig.txt" -Credential $Global:Creds
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
        if ($result.returnvalue -eq 0) {
            "Connecting to C$ on $s"
            New-PSDrive REMOTE -PSProvider FileSystem -Root \\$s\c$\temp -Credential $Global:Creds > $null 
            "Moving $s-ipconfig.txt to local machine"
            move-item -path REMOTE:\$s-ipconfig.txt $SavePath -force
            "Disconnect $s" 
            Remove-PSDrive REMOTE
            }
    "-----------------------------------------"
    }
}
#**************************
# EndFunction Get-IPConfigs
#**************************

#**************************
# Get VM IPConfigsWithLocal
#**************************
Function Get-IPConfigsWithLocal {
    $SavePath = "$Global:WorkFolder\IPConfig"
    $servers = Get-Content "$Global:WorkFolder\server.txt"

    forEach ($s in $servers) {
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
        if ($result.returnvalue -eq 0) {
            "Connecting to C$ on $s"
            New-PSDrive REMOTE -PSProvider FileSystem -Root \\$s\c$\temp -Credential $Global:CredsLocal > $null 
            "Moving $s-ipconfig.txt to local machine"
            move-item -path REMOTE:\$s-ipconfig.txt $SavePath -force
            "Disconnect $s" 
            Remove-PSDrive REMOTE
            }
    "-----------------------------------------"
    }
}
#**************************
# EndFunction Get-IPConfigs
#**************************


#********************
# Build Word Document
#********************
Function Build-Word {
    "Building Word Document"
    $wdOrientLandscape = 1

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
    $Selection.Style="Strong"
    $Selection.TypeText("IP Configurations")
    $Selection.TypeParagraph()

    #Add IPConfig Files
    Get-ChildItem "$Global:WorkFolder\IPConfig" |
    ForEach-Object{
        $Selection.InsertFile("$Global:WorkFolder\IPConfig\$_")
        $Selection.TypeParagraph()
        $Selection.TypeText("-------------------")
        $Selection.InsertNewPage()  
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
$Global:Creds = Get-Credential
if ($WithLocalCreds) {
        "Enter local credentials to pull IP Info"
        "I.E:  localhost\rhsadmin"
        $Global:CredsLocal = Get-Credential
        }
CLS
Get-VCenter
$Global:RunAgain = "y"
Do {
    Get-HostName
    CLS
    $Global:WorkFolder = "$Global:Folder\$Global:HostName"
    Verify-Folders
    Get-HostInfo
    
    if ($WithLocalCreds) {
        Get-IPConfigsWithLocal
        }
    Else {
        Get-IPConfigs
        }

    Build-Word
    "Data Collection for $Global:HostName is complete!"
    "================================================="
    " "
    Run-Again
    } While ($Global:RunAgain -eq "y")

"Open Explorer to $Global:Folder"
Invoke-Item $Global:Folder
Clean-Up
Stop-Transcript

