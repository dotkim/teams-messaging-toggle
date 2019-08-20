# script to toggle the messaging and metting policy for a group of users.
# the point is for an entire school to be able to change the policy.

# function for logging
function LogToFile {
  param (
    [string]$Message
  )
  Add-Content -Path "C:\\Github\\teams-messaging-toggle\\toggle-teams-messaging-policy.ps1.log" -Value "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") - $($message)"
}

function Get-ConfigFile {
  try {
    $config = Get-Content -Path ".\toggle-teams-messaging-policy.ps1.config" -ErrorAction Stop
    $config = ($config | ConvertFrom-Json)
    return $config
  }
  catch {
    LogToFile -Message "Caught error when fetching config file"
    LogToFile -Message "$($Error.Exception.Message)"
    $Error.Clear()
    Exit
  }
}

$startTime = get-date
LogToFile -Message "####### Script started #######"

# get the config file
$config = Get-ConfigFile

# create credentials for the user
# this users needs access to Azure AD and Teams to work
try {
  $username = $config.Username
  $password = ConvertTo-SecureString $config.Password -AsPlainText -Force -ErrorAction Stop
  $credentials = New-Object System.Management.Automation.PSCredential $username, $password -ErrorAction Stop
}
catch {
  LogToFile -Message "Error while creating credential"
  LogToFile -Message "$($Error.Exception.Message)"
  $Error.Clear()
  Exit
}

# get the distribution lists which are going to be changed.
# This should be a list of ObjectIDs
try {
  $groups = Get-Content -Path $config.GroupListPath -ErrorAction Stop
}
catch {
  LogToFile -Message "Caught error when fetching groups"
  LogToFile -Message "$($Error.Exception.Message)"
  $Error.Clear()
  Exit
}

# connect to the azure ad tenant
try {
  Import-Module AzureAD
  Connect-AzureAD -TenantId $config.TenantId -Credential $credentials
}
catch {
  LogToFile -Message "Caught error when connecting to AzureAD"
  LogToFile -Message "$($Error.Exception.Message)"
  $Error.Clear()
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
    LogToFile -Message "$($Error.Exception.Message)"
    $Error.Clear()
  }
}

LogToFile -Message "Found $($users.count) users"

try {
  Import-Module "C:\Program Files\Common Files\Skype for Business Online\Modules\SkypeOnlineConnector\SkypeOnlineConnector.psd1"
  $session = New-CsOnlineSession -OverrideAdminDomain "digirom.onmicrosoft.com" -Credential $credentials
  Import-PSSession $session
}
catch {
  LogToFile -Message "$($Error.Exception.Message)"
  Exit
}

LogToFile -Message "Granting policy to users..."
$users | ForEach-Object {
  try {
    LogToFile -Message "User: $($_.UserPrincipalName)"
    Grant-CsTeamsMessagingPolicy -Identity $_ -PolicyName $config.CsTeamsMessagingPolicyAllow
    Grant-CsTeamsMeetingPolicy -identity $_ -PolicyName $config.CsTeamsMeetingPolicyAllow
  }
  catch {
    LogToFile -Message "$($Error.Exception.Message)"
    $Error.Clear()
  }
}

$endTime = get-date
LogToFile -Message "Done. Time for the full run was: $(New-TimeSpan $startTime $endTime). Number of users affected: $($users.count)"