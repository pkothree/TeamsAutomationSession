###
# Danny Davis
# Session Example
# Teams Creation Automation
# Session Part 01
# Create a Team from a SharePoint list with PowerShell
# Created: 09/25/19
# Modified: 05/05/20
###

# First we need to find the Teams module
Find-Module *teams*
# Now we can get the Teams module name and install it
Install-Module MicrosoftTeams

# If we want to use the module in an Azure Function we have to save the module
Save-Module MicrosoftTeams -Repository PSGallery -Path "C:\temp"

# To use the cmdlets from the module, we have to import the module
Import-Module MicrosoftTeams

# Azure Tenant Id
# the tenant id is needed to identify the tenant you want to use
# the tenant id can be found within the Azure AD settings
$tenant = ""

# We will have to provide credentials to access the cloud
# there are multiple ways, in this case we simply provide them manually
$credentials = Get-Credential

# Connect to Teams service
# -TenantId is optional
Connect-MicrosoftTeams -Credential $credentials #-TenantId $tenant

# Display all exisitng Teams
Get-Team

# Choose one DisplayName or one GroupId to get a specific Team
# Get Team by Displayname
# Get-Team -Displayname Customer1
Get-Team -Displayname ""

# Get Team by GroupId
Get-Team -GroupId ""

# Create a new Team
New-Team -DisplayName "Test Team Meetup" -MailNickname "TestTeamMeetup" -Visibility "Public"

# Get the Team and store the object
$team = Get-Team -Displayname "Test Team Meetup"

# Display the Team GroupId
$team.GroupId

# You can use the GroupId like this and add new users
Add-TeamUser -GroupId $team.GroupId -User "user@yourdomain.com"

# Add a new channel to your Team
New-TeamChannel -GroupId $team.GroupId -DisplayName "Project007"