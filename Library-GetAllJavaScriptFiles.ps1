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

<#
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
Import-Module 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll' 
#>

#Site collection Variable
$SiteURL="http://<WebApp>"
$ReportOutput="C:\temp\SiteInventory.csv"
 
#Get the site collection
$Site = Get-SPSite $SiteURL
 
$ResultData = @()
#Ge All Sites of the Site collection
Foreach($web in $Site.AllWebs)
{
    Write-host -f Yellow "Processing Site: "$Web.URL
  
    #Get all lists - Uncomment to Exclude Hidden System lists
    $ListCollection = $web.lists #| Where-Object  { ($_.hidden -eq $false) -and ($_.IsSiteAssetsLibrary -eq $false)}
 
    #Iterate through All lists and Libraries
    ForEach ($List in $ListCollection)
    {
        if($List.BaseTemplate -eq "DocumentLibrary")
        {
            Write-host -f Cyan "Processing Document Library: '$($List.Title)' with $($List.ItemCount) Item(s)"
  
            Do {
               
                #Filter Files to retrieve only JavaScript files.
                $Files = $List.Items | Where-Object {$_.Name -like "*.js*"}
 
                $DocumentInventory = @()
                Foreach($Item in $Files)
                {
                    $File = $Item.File

                    $DocumentData = New-Object PSObject
                    $DocumentData | Add-Member NoteProperty SiteURL($SiteURL)
                    $DocumentData | Add-Member NoteProperty DocLibraryName($List.Title)
                    $DocumentData | Add-Member NoteProperty FileName($File.Name)
                    $DocumentData | Add-Member NoteProperty FileURL($File.ServerRelativeUrl)
                    $DocumentData | Add-Member NoteProperty CreatedBy($File["Author"].Email)
                    $DocumentData | Add-Member NoteProperty CreatedOn($File.TimeCreated)
                    $DocumentData | Add-Member NoteProperty ModifiedBy($File["Editor"].Email)
                    $DocumentData | Add-Member NoteProperty LastModifiedOn($File.TimeLastModified)
                    $DocumentData | Add-Member NoteProperty Size-KB([math]::Round($File.Length/1KB))
                        
                    #Add the result to an Array
                    $DocumentInventory += $DocumentData
                }
                #Export the result to CSV file
                $DocumentInventory | Export-CSV $ReportOutput -NoTypeInformation -Append
                $Query.ListItemCollectionPosition = $ListItems.ListItemCollectionPosition
            } While($Query.ListItemCollectionPosition -ne $null)

        } 
    } 
}
Write-Output $DocumentInventory | Format-Table
 
Write-host -f Green "Report Generated Successfully at : "$ReportOutput