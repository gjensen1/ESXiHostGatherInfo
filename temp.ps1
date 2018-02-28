

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
$Global:FileServer = $null
$Global:PrintServer = $null
$Global:PoweredOn = $null
$Global:IsWindows = $null

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
    " "
    "Folder Structure built" 
}
#***************************
# EndFunction Verify-Folders
#***************************

#*******************
# Connect to vCenter
#*******************
Function Connect-VC {
    "-----------------------------------------"
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
    "-----------------------------------------"
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

    "-----------------------------------------"
}
#**********************
# EndFunction GetVMInfo
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

$Global:Creds = Get-Credential
$Global:CredsLocal = $Global:Creds
Get-VCenter
Connect-VC
$Global:RunAgain = "y"
Do {
    Get-HostName
    CLS
    $Global:WorkFolder = "$Global:Folder\$Global:HostName"
    Verify-Folders
    Get-HostInfo
    
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
                    "VM $Server is not Windows, skipping additional steps"
                    }

            }
            Else {
                "VM $Server is not Powered On skipping additional steps "
            }
        }
    Build-Word
    Run-Again
    } While ($Global:RunAgain -eq "y")
Disconnect-VC
"Open Explorer to $Global:Folder"
Invoke-Item $Global:Folder
Clean-Up