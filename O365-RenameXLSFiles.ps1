<#

 Microsoft provides programming examples for illustration only, without warranty either expressed or
 implied, including, but not limited to, the implied warranties of merchantability and/or fitness 
 for a particular purpose. 
 
 This sample assumes that you are familiar with the programming language being demonstrated and the 
 tools used to create and debug procedures. Microsoft support professionals can help explain the 
 functionality of a particular procedure, but they will not modify these examples to provide added 
 functionality or construct procedures to meet your specific needs. if you have limited programming 
 experience, you may want to contact a Microsoft Certified Partner or the Microsoft fee-based consulting 
 line at (800) 936-5200. 


==============================================================#>

#################### Parameters ###########################################
$siteUrl = "<TenantURL>"
$webUrl = "$siteUrl/sites/<SiteCollection>";
$listUrl = "<DocumentLibrary";


$User = "<SiteCollectionAdminUPN>"
$PWord = ConvertTo-SecureString -String "<Password>" -AsPlainText -Force
$credentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $User, $PWord

Add-Type -AssemblyName Microsoft.Office.Interop.Excel
$xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook

Connect-PnPOnline -Url $webUrl -Credentials $credentials
$web = Get-PnPWeb
$list = Get-PNPList -Identity $listUrl

#################### /Parameters ###########################################

function ProcessFolder($folderUrl, $destinationFolder) {

    $folder = Get-PnPFolder -RelativeUrl $folderUrl
    $tempfiles = Get-PnPProperty -ClientObject $folder -Property Files
   
    if (!(Test-Path -path $destinationfolder)) {
        $dest = New-Item $destinationfolder -type directory 
    }

    $total = $folder.Files.Count
    For ($i = 0; $i -lt $total; $i++) {
        $file = $folder.Files[$i]
        
        
        Get-PnPFile -ServerRelativeUrl $file.ServerRelativeUrl -Path $destinationfolder -FileName $file.Name -AsFile
        if($file.Name.EndsWith(".xls"))
        {
            $excel = New-Object -ComObject Excel.Application
            $excel.visible = $false
            $excel.DisplayAlerts = $false
            $excel.WarnOnFunctionNameConflict = $False
            $workbook = $excel.workbooks.open($siteUrl +  $file.ServerRelativeUrl,2,$True) 
            $fileName = $file.ServerRelativeUrl
            $fileNameNew = $fileName.Replace(".xls",".xlsx")
            $path += $siteUrl +  $fileNameNew
            
            $workbook.saveas($path, 51, [Type]::Missing, [Type]::Missing, $false, $false, 1, 2)
            $workbook.close()
            $excel.Quit()
            $excel = $null
            [gc]::collect()
            [gc]::WaitForPendingFinalizers()
        }
    }
}


ProcessFolder $listUrl $destination + "\" 
