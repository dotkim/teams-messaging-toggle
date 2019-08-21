# script to toggle the messaging and metting policy for a group of users.
# the point is for an entire school to be able to change the policy.

# function for logging
function LogToFile {
  <#
    .LogToFile
    Parameters: Message
    Logs the message provided to the log file.
    INFO, WARN and ERROR tags are added for the CM Trace Log Tool. Which is not baked into the function.
  #>
  param (
    [string]$Message
  )
  if (!$Message) { return }
  Add-Content -Path ".\toggle-teams-messaging-policy.ps1.log" -Value "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") - $($message)"
}

function Get-ConfigFile {
  <#
    .Get-ConfigFile
    Fetches the config for the script from disk.
    Returns PSCustomObject (Converted JSON)
  #>
  try {
    $config = Get-Content -Path ".\toggle-teams-messaging-policy.ps1.config" -ErrorAction Stop
    $config = ($config | ConvertFrom-Json)
    return $config
  }
  catch {
    LogToFile -Message "ERROR: Caught error when fetching config file"
    LogToFile -Message "ERROR: $($Error.Exception.Message)"
    $Error.Clear()
    Exit
  }
}

$startTime = Get-Date
LogToFile -Message "INFO: Script started"

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
  LogToFile -Message "ERROR: While creating credential"
  LogToFile -Message "ERROR: $($Error.Exception.Message)"
  $Error.Clear()
  Exit
}

# get the distribution lists which are going to be changed.
# This should be a list of ObjectIDs
try {
  $groups = Get-Content -Path $config.GroupListPath -ErrorAction Stop
  if ($groups.count -lt 1) {
    LogToFile -Message "WARN: No groups are listed, exiting"
    Exit
  }
}
catch {
  LogToFile -Message "ERROR: Caught error when fetching groups"
  LogToFile -Message "ERROR: $($Error.Exception.Message)"
  $Error.Clear()
  Exit
}

# connect to the azure ad tenant
try {
  Import-Module AzureAD
  Connect-AzureAD -TenantId $config.TenantId -Credential $credentials
}
catch {
  LogToFile -Message "ERROR: Caught error when connecting to AzureAD"
  LogToFile -Message "ERROR: $($Error.Exception.Message)"
  $Error.Clear()
  Exit
}


# get the users for which to set the policy
LogToFile -Message "INFO: Fetching users to add"
$users = @()

$groups | ForEach-Object {
  try {
    $members = Get-AzureADGroupMember -ObjectId $_ -ErrorAction Stop
    if ($members.count -ne 0) {
      $users += $members
    }
  }
  catch {
    LogToFile -Message "ERROR: $($Error.Exception.Message)"
    $Error.Clear()
  }
}

LogToFile -Message "INFO: Found $($users.count) users"

try {
  # seems like this is the best way to get the module, as it is wierdly enough not accessable by its module name sometimes
  Import-Module "C:\Program Files\Common Files\Skype for Business Online\Modules\SkypeOnlineConnector\SkypeOnlineConnector.psd1"
  $session = New-CsOnlineSession -OverrideAdminDomain "digirom.onmicrosoft.com" -Credential $credentials -ErrorAction Stop
  Import-PSSession $session
}
catch {
  LogToFile -Message "ERROR: $($Error.Exception.Message)"
  Remove-PSSession $session
  Exit
}

# sets the policies for allowing messaging and meetings for each user
LogToFile -Message "INFO: Granting policy to users..."
$users | ForEach-Object {
  try {
    LogToFile -Message "INFO: User: $($_.UserPrincipalName)"
    Grant-CsTeamsMessagingPolicy -Identity $_.UserPrincipalName -PolicyName $config.CsTeamsMessagingPolicyAllow -ErrorAction Stop
    Grant-CsTeamsMeetingPolicy -identity $_.UserPrincipalName -PolicyName $config.CsTeamsMeetingPolicyAllow -ErrorAction Stop
  }
  catch {
    LogToFile -Message "ERROR: $($Error.Exception.Message)"
    $Error.Clear()
  }
}

# remove the skype online session for cleanup purposes.
# if the script is ran within 1 hour of the session starting the script will fail because of user policies.
# this is ONLY if the session is not removed porperly.
Remove-PSSession $session

$endTime = Get-Date
LogToFile -Message "INFO: Done. Time for the full run was: $(New-TimeSpan $startTime $endTime). Number of users affected: $($users.count)"