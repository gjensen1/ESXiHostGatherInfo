
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
-----------------------------------------------------------------------------

******************************************************************************************
******************************************************************************************
** This script uses the VMware modules installed by the installation of VMware PowerCLI **
**               ENSURE that VMware PowerCLI has been installed.                        **
**                                                                                      **    
**    Installation media can be found here:                                             **
**    \\cihs.ad.gov.on.ca\tbs\Groups\ITS\DCO\RHS\RHS\Software\VMware\VMware PowerCLI    **
**                                                                                      **
**          ======================================================                      **
**	    Version 5.8 is the version used to develop this script                      **
**          ======================================================                      **
**                                                                                      **
******************************************************************************************
******************************************************************************************


Usage:

Simply run the script, you will be prompted for Credentials, and then the vCenter and Host that you want to work with.

All results will stored in your Documents folder under HostUpgradeInfo.


