$wdOrientLandscape = 1
$strComputer = “.”
$CRQ = "CRQ12345"
$WorkFolder = "c:\HostUpgradeInfo\$CRQ"
$VHostName = "itspkeseesx101.cihs.gov.on.ca"



$Word = New-Object -ComObject Word.Application
$Word.Visible = $True
$Document = $Word.Documents.Add()
$Selection = $Word.Selection
$Range = $Document.Range()
$Document.PageSetup.Orientation = $wdOrientLandscape

#Get HostName and SiteAddress from VMAnnotations File and add to Word Doc
$SiteInfo = Import-Csv $WorkFolder\Annotations\VMHost-$VHostName.csv
$SiteAddress = $SiteInfo | where {$_.Name -eq "Site Address"} | Select Value | Format-Table -HideTableHeaders | Out-String
$HostName = $SiteInfo | where {$_.Name -eq "Site Address"} | Select VM | Format-Table -HideTableHeaders | Out-String
$SiteAddress = $SiteAddress.Trim()
$HostName = $HostName.Trim()

#Build Word Doc
#Title Page
$Selection.Style="Intense Quote"
$Selection.TypeText("Gather Information for ESXi Host Upgrades ")
$Selection.TypeText((Get-Date))
$Selection.TypeParagraph()
$Selection.Style="Intense Reference"
$Selection.TypeText("Site: $SiteAddress")
$Selection.TypeParagraph()
$Selection.Style="Intense Reference"
$Selection.TypeText("Host: $HostName")
$Selection.TypeParagraph()
$Selection.Style="Strong"
$Selection.TypeText("VMs Present on Host:")
$Selection.TypeParagraph()
$Selection.InsertFile("$WorkFolder\server.txt")
$Selection.InsertNewPage()
$Selection.Style="Strong"
$Selection.TypeText("IP Configurations")
$Selection.TypeParagraph()

#Add IPConfig Files
Get-ChildItem "$WorkFolder\IPConfig" |
ForEach-Object{
    $Selection.InsertFile("$WorkFolder\IPConfig\$_")
    $Selection.TypeParagraph()
    $Selection.TypeText("-------------------")
    $Selection.InsertNewPage()  
    }

#Add Annotation Files
$Selection.Style="Strong"
$Selection.TypeText("Host Annotations")
$Selection.TypeParagraph()
$Selection.InsertFile("$WorkFolder\Annotations\VMHost-$VHostName.csv")
$Selection.InsertNewPage()
$Selection.Style="Strong"
$Selection.TypeText("VM Annotations")
$Selection.TypeParagraph()
$Selection.InsertFile("$WorkFolder\Annotations\VMAnnotations-$VHostName.csv")

#Return to top
$FileName = "$WorkFolder\$VHostName.docx"
$Document.SaveAs([ref] $FileName)
$Document.Close()
$Word.Quit()