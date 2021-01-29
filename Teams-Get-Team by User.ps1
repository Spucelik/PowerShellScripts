#Install-Module MicrosoftTeams -Force -AllowClobber
$connectTeams = Connect-MicrosoftTeams

$Teams = Get-Team -User <UserPrincipalName> 
$TeamDetails=@()

foreach ($team in $Teams) {
    
    $teamLinkTemplate = "https://teams.microsoft.com/l/team/<ThreadId>/conversations?groupId=<GroupId>&tenantId=<TenantId>"
    $channel = Get-TeamChannel -GroupId $team.GroupId | Where-Object {$_.DisplayName -eq "General"} | Select-Object -First 1
    $teamLink = $teamLinkTemplate.Replace("<ThreadId>",$channel.Id).Replace("<GroupId>",$team.GroupId).Replace("<TenantId>",$connectTeams.TenantId)
    $teamLink

    $TeamDetailsItem = New-Object PSObject
    $TeamDetailsItem | Add-Member -MemberType NoteProperty -Name "TeamName" -Value $team.DisplayName
    $TeamDetailsItem | Add-Member -MemberType NoteProperty -Name "TeamURL" -Value $teamLink
    $TeamDetails += $TeamDetailsItem

   
}

$TeamDetails | export-csv -Path c:\temp\TeamDetails.csv -NoTypeInformation