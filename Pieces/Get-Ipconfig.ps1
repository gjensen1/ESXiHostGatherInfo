    $SavePath = "$env:USERPROFILE\Documents\IPConfigInfo"
    $servers = Get-Content "$env:USERPROFILE\Documents\IPConfigInfo\servers.txt"
    $Creds = Get-Credential


    forEach ($s in $servers) {
        "Running IPConfig on $s"
        $result = invoke-wmimethod -computer $s -path Win32_process -name Create -ArgumentList "cmd /c ipconfig /all > c:\temp\$s-ipconfig.txt" -Credential $Creds
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
            New-PSDrive REMOTE -PSProvider FileSystem -Root \\$s\c$\temp -Credential $Creds > $null 
            "Moving $s-ipconfig.txt to local machine"
            move-item -path REMOTE:\$s-ipconfig.txt $SavePath -force
            "Disconnect $s" 
            Remove-PSDrive REMOTE
            }
         "-----------------------------------------"
        }