
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

 
 ----------------------------------------------------------
 Purpose:
 ----------------------------------------------------------

 The purpose of this script is to itterate through all the results specified in the search query and export the Path name to the specified text file..

 ==============================================================#>

Connect-PnPOnline https://pucelikdemo.sharepoint.com/sites/FlowTraining -UseWebLogin

$SearchResults = Submit-PnPSearchQuery -Query "-path:https://<Tenant>.sharepoint.com/sites/<SiteName>/lists/LargeList*" -All -SelectProperties Path -TrimDuplicates $false 

$allshareditems = $SearchResults | ? {$_.TableType -like "RelevantResults"}

write-host "found " $allshareditems.TotalRows " rows"

$counter = 1

"Path" | Out-File Output-filename.txt -Force

foreach($item in $allshareditems.ResultRows){

    $item.Path | Out-File c:\temp\SITEshareditems.txt -Append

    Write-Progress  -Activity "Exporting results" -PercentComplete ($counter/$allshareditems.TotalRows *100) -Status $counter.ToString() 

    $counter++

    }