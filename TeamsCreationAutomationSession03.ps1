###
# Danny Davis
# Session Example
# Teams Creation Automation
# Session Part 03
# Create a Team from a SharePoint list with PowerShell
# This is for a pre-release PowerShell Azure Function
# Created: 09/25/19
# Modified: 05/05/20
###

# Azure Function header start
# POST method: $req
$requestBody = Get-Content $req -Raw | ConvertFrom-Json
$name = $requestBody.name

# GET method: each querystring parameter is its own variable
if ($req_query_name) 
{
    $name = $req_query_name 
}
if ($req_query_ItemID) 
{
    #$itemID = $req_query_ItemID 
    $itemID = 1
}
if ($req_query_URL) 
{
    #$url = $req_query_Url
    $url = "https://yourtenantname.sharepoint.com/sites/yoursite"
}
if ($req_query_ListTitle) 
{
    #$listTitle = $req_query_ListTitle
    $listTitle = "TeamsCreation"
}

Out-File -Encoding Ascii -FilePath $res -inputObject "Hello $name"
# Azure Function header end

# Create Context for PowerShell Modules and User Credentials (connection to O365, O365 Admin)
$FunctionName = 'HttpTriggerPowerShell1'

# Define Modules
$TeamsModuleName = 'MicrosoftTeams'
$TeamsVersion = '1.0.2'

$username = $Env:user
$pw = $Env:password

# Import PS modules
$TeamsModulePath = "D:\home\site\wwwroot\$FunctionName\bin\$TeamsModuleName\$TeamsVersion\$TeamsModuleName.psd1"
$res = "D:\home\site\wwwroot\$FunctionName\bin"
 
Import-Module $TeamsModulePath
 
# Build Credentials
$keypath = "D:\home\site\wwwroot\$FunctionName\bin\keys\PassEncryptKey.key"
$pwfile = @(Get-Content $keypath)[0]
$credentials= New-Object System.Management.Automation.PSCredential ($username, $secpassword)

# Tenant ID
$tenant = ""

# Connect to SharePoint Online Service
Connect-PnPOnline -Url $url -Credentials $credentials
$item = Get-PNPListItem -List Lists/$listTitle -Id $itemID

# Connect to Microsoft Teams
Connect-MicrosoftTeams -TenantId $tenant -Credential $credentials

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
        Set-TeamFunSettings -GroupId $team.GroupId -AllowGiphy $false -GiphyContentRating Strict -AllowStickersAndMemes $false
        New-TeamChannel -GroupId $team.GroupId -DisplayName "Phase 1: Planning"
        New-TeamChannel -GroupId $team.GroupId -DisplayName "Phase 2: Customer Meetings"
        # Add creator of the request as an owner
        Add-TeamUser -GroupId $team.GroupId -User $item.FieldValues.Author.Email -Role Owner
    }
if($type -eq "Department")
    {
        # We can set the fun settings. For Departments we allow Giphy, a little fun can't hurt
        Set-TeamFunSettings -GroupId $team.GroupId -AllowGiphy $true -GiphyContentRating Moderate -AllowStickersAndMemes $true
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