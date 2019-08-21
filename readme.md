# teams-messaging-toggle #

Script for currently: setting the messaging and meeting policy for an entire distributiongroup.

A few files must be made before use:

- toggle-teams-messaging-policy.ps1.txt
  - A list of objectId's, these are the groups to apply the policies to.
- toggle-teams-messaging-policy.ps1.log
  - Will not be made from the commands used, has to be made manually.
- toggle-teams-messaging-policy.ps1.config
  - see example below.
  
## Example config file ##

```json
{
  "Username": "user",
  "Password": "pw",
  "TenantId": "id",
  "CsTeamsMessagingPolicyAllow": "allowmsg",
  "CsTeamsMeetingPolicyAllow": "allowmeet",
  "GroupListPath": "PATH TO toggle-teams-messaging-policy.ps1.txt"
}
```
