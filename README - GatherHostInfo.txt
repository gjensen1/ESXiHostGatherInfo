Purpose of Script:

   Gathers and documents important information that may be required
   during the upgrade of an ESXi Host.  This includes:
    - VMs running on the host
    - VM IPConfig info     
    - Host Annotations
    - VM Annotations

   Prompted inputs:  Credentials, vCenterName, VMHostName

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

Usage:

Simply run the script, you will be prompted for Credentials, and then the vCenter and Host that you want to work with.

All results will stored in your Documents folder under HostUpgradeInfo.


