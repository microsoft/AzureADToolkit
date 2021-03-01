<# 
 .Synopsis
  Gets a report of all the applications and service principals in this tenant that have either a password or client secret

 .Description
  This functions returns a list of all applications and service principals that have a credential

 .Example
  Get-AADToolkitApplicationCredentials | Export-Csv -Path '.\AppPermissions.csv'  -NoTypeInformation
  Generates a CSV report of all applications and service principals with credentials.
#>

function Get-AADToolkitApplicationCredentials {
    function Get-CredentialInfo ($objectType, $item, $cred, $credentialType)
    {
        [pscustomobject]@{
            ObjectId = $item.id
            AppDisplayName = $item.displayName
            ObjectType = $objectType
            AppId = $item.appId
            Credentialtype = $credentialType
            KeyId = $cred.keyId
            Hint = $cred.hint
            CredDisplayName = $cred.displayName
            StartDateTime = $cred.startDateTime
            EndDateTime = $cred.endDateTime
            KeyType = $cred.type
            Usage = $cred.usage
        }
    }

    function Get-CredentialReport ($objectType)
    {
        $reportJson = Invoke-AADTGraph -Uri "/$objectType"        
        do
        {
            foreach($item in $reportJson.value)
            {
                foreach($cred in $item.passwordCredentials)
                {
                    Get-CredentialInfo $objectType $item $cred "PasswordCredential"
                }
                foreach($cred in $item.keyCredentials)
                {
                    Get-CredentialInfo $objectType $item $cred "KeyCredential"
                }
            }
            if($null -ne $reportJson.'@odata.nextLink') { $reportJson = Invoke-GraphRequest -Uri $reportJson.'@odata.nextLink' }
        } while ($null -ne $reportJson.'@odata.nextLink')     
    }

    Get-CredentialReport "applications"
    Get-CredentialReport "servicePrincipals"
}