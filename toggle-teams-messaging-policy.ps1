# script to toggle the messaging and metting policy for a group of users.
# the point is for an entire school to be able to change the policy.

# function for logging
function LogToFile {
  param (
    [string]$Message
  )
  Add-Content -Path 'C:\temp\toggle-policy.log' -Value "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") - $($message)"
}

# create credentials for the user
# this users needs access to Azure AD and Teams to work
$username = ''
$password = ''
$credentials = New-Object System.Management.Automation.PSCredential $username, $password

# get the distribution lists which are going to be changed.
# This should be a list of ObjectIDs
$groups = Get-Content -Path 'C:\temp\distributionlists.txt'

# connect to the azure ad tenant
Import-Module AzureAD
try {
  Connect-AzureAD -TenantId $tenant -Credential $credentials
}
catch {
  LogToFile -Message "$($Error.Message)"
  Exit
}

# get the users for which to set the policy
LogToFile -Message "Fetching users to add"
$users = @()

$groups | ForEach-Object {
  try {
    $members = Get-AzureADGroupMember -ObjectId $_ -ErrorAction Stop
    if ($members.count -ne 0) {
      $users += $members
    }
  }
  catch {
    LogToFile -Message "$($Error.Message)"
    $Error.Clear()
  }
}

LogToFile -Message "Found $($users.count) users"

Import-Module "C:\Program Files\Common Files\Skype for Business Online\Modules\SkypeOnlineConnector\SkypeOnlineConnector.psd1"
$session = New-CsOnlineSession -OverrideAdminDomain "digirom.onmicrosoft.com" -Credential $credentials
try {
  Import-PSSession $session
}
catch {
  LogToFile -Message "$($Error.Message)"
  Exit
}

LogToFile -Message "Granting policy to users..."
$users | ForEach-Object {
  try {
    LogToFile -Message "User: $($_.UserPrincipalName)"
    Grant-CsExternalAccessPolicy -Identity $_ -PolicyName ""
  }
  catch {
    LogToFile -Message "$($Error.Message)"
    $Error.Clear()
  }
}