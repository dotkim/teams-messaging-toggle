# script to toggle the messaging and metting policy for a group of users.
# the point is for an entire school to be able to change the policy.

# function for logging
function LogToFile {
  <#
    .SYNOPSIS
      Logs the message provided to the log file.
    .DESCRIPTION
      Logs the message provided to the log file.
    .PARAMETER Message
      The string value you want to log
    .EXAMPLE
      LogToFile -Message "INFO: Script started"
  #>
  param (
    [parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Message
  )
  if (!$Message) { return }
  Add-Content -Path ".\toggle-teams-messaging-policy.ps1.log" -Value "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") - $($message)"
}

function Get-ConfigFile {
  <#
    .SYNOPSIS
      Fetches the config for the script from disk.
    .DESCRIPTION
      Fetches the config for the script from disk.
    .EXAMPLE
      Get-ConfigFile
    .NOTES
      Returns PSCustomObject (Converted from JSON)
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
$usersUnfiltered = @()
$users = @()

$groups | ForEach-Object {
  try {
    $members = Get-AzureADGroupMember -All $true -ObjectId $_ -ErrorAction Stop
    if ($members.count -ne 0) {
      $usersUnfiltered += $members
    }
  }
  catch {
    LogToFile -Message "ERROR: $($Error.Exception.Message)"
    $Error.Clear()
  }
}

LogToFile -Message "INFO: Found $($usersUnfiltered.count) users to filter"

try {
  LogToFile -Message "INFO: Checking if users already have the policies."
  $AllowedUsers = Get-Content -Path $config.UsersEnabled -ErrorAction Stop
  $AllowedUsers = ($AllowedUsers | ConvertFrom-Json)
  if (!$AllowedUsers) {
    $AllowedUsers = @{}
  }
  $usersUnfiltered | ForEach-Object {
    $name = $_.UserPrincipalName.split("@")[0]
    if (!($AllowedUsers | Get-Member $name)) {
      $users += $_
    }
  }
  if ($users.count -lt 1) {
    LogToFile -Message "WARN: Found $($users.count) users to set, exiting script"
    Exit
  }
  LogToFile -Message "INFO: Found $($users.count) users to set"
}
catch {
  LogToFile -Message "ERROR: $($Error.Exception.Message)"
  $Error.Clear()
}

try {
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
    $name = $_.UserPrincipalName.split("@")[0]
    Grant-CsTeamsMessagingPolicy -Identity $_.UserPrincipalName -PolicyName $config.CsTeamsMessagingPolicyAllow -ErrorAction Stop
    Grant-CsTeamsMeetingPolicy -identity $_.UserPrincipalName -PolicyName $config.CsTeamsMeetingPolicyAllow -ErrorAction Stop
    $AllowedUsers | Add-Member -MemberType "NoteProperty" -Name $($name) -Value $true -Force
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


($AllowedUsers|ConvertTo-Json) | Out-File -FilePath $config.UsersEnabled -Encoding UTF8 -Force
$endTime = Get-Date
LogToFile -Message "INFO: Done. Time for the full run was: $(New-TimeSpan $startTime $endTime). Number of users affected: $($users.count)"