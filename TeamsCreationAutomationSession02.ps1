###
# Danny Davis
# Session Example
# Teams Creation Automation
# Session Part 02
# Create a Team from a SharePoint list with PowerShell
# Created: 09/25/19
# Modified: 05/05/20
###
 
Import-Module MicrosoftTeams
 
# Build Credentials
$credentials = Get-Credential

# Tenant ID
$tenant = ""

# URL to the SharePoint to pull information from
$url = "https://yourtenantname.sharepoint.com/sites/yoursite"
$listTitle = "TeamsCreation"

# Connect to SharePoint Online Service
Connect-PnPOnline -Url $url -Credentials $credentials
$item = Get-PNPListItem -List Lists/$listTitle -Id 5

# Connect to Microsoft Teams
Connect-MicrosoftTeams #-TenantId $tenant -Credential $credentials

# Every type of team will have a different configuration
# We have to get the type of team we are going to create
$type = $item.FieldValues.TeamType

# We will always create a new Team
$team = New-Team -DisplayName $item.FieldValues.Title -Visibility $item.FieldValues.Visibility -Owner $item.FieldValues.Owner.Email -MailNickName $item.FieldValues.MailNickName

# Make a decision depending on the type
# We will use a simple if-clause
if($type -eq "Project")
    {
        # We can set the fun settings. For Projects we disable Giphy, because external people might work on it
        Set-Team -GroupId $team.GroupId -AllowGiphy $false -GiphyContentRating Strict -AllowStickersAndMemes $false
        New-TeamChannel -GroupId $team.GroupId -DisplayName "Phase 1 Planning"
        New-TeamChannel -GroupId $team.GroupId -DisplayName "Phase 2 Customer Meetings"
        # Add creator of the request as an owner
        Add-TeamUser -GroupId $team.GroupId -User $item.FieldValues.Author.Email -Role Owner
    }
if($type -eq "Department")
    {
        # We can set the fun settings. For Departments we allow Giphy, a little fun can't hurt
        Set-Team -GroupId $team.GroupId -AllowGiphy $true -GiphyContentRating Moderate -AllowStickersAndMemes $true
        New-TeamChannel -GroupId $team.GroupId -DisplayName "Team Meetings"
        # Add creator of the request as an owner
        Add-TeamUser -GroupId $team.GroupId -User $item.FieldValues.Author.Email -Role Owner
    }
else
    {
        # We don't know what this team is for, so we are adding an admin to the team
        Add-TeamUser -GroupId $team.GroupId -User "fh@pko3.com" -Role Owner

        # There is no reason to create a Team if we don't know what it is used for
        # But just run with it ;)
    }