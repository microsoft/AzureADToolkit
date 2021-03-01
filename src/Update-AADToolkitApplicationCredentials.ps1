<# 
 .Synopsis
  Helper utility to remove or roll over the client secrets and certificates of a given application or service principals in Azure Active Directory. 
  Run Get-AADToolkitApplicationCredentials to get a summary of all the applications and service principals 

 .Description
  This interactive function allows a user to select an application or service principle and manage it's credentials. 

 .Example
  Update-AADToolkitApplicationCredentials
  
#>

function Update-AADToolkitApplicationCredentials
{
    function Show-AppInfo ($appInfo){
        Write-Host
        $appDisplay = $appInfo | Format-List -Property objectType, objectId, displayName, appId | Out-String
        $appCreds = $appInfo.creds | Format-Table -Property keyId, startDateTime, endDateTime, expired, credentialtype, description | Out-String
        Write-Host $appDisplay
        Write-Host $appCreds
    }
    function Show-Menu {
        Param($appInfo,
            [string[]]$MenuItems,
            [string] $Title
        )

        $header = $null
        if (![string]::IsNullOrWhiteSpace($Title)) {
            $len = [math]::Max(($MenuItems | Measure-Object -Maximum -Property Length).Maximum, $Title.Length)
            $header = '{0}{1}{2}' -f $Title, [Environment]::NewLine, ('-' * $len)
        }

        # possible choices: didits 1 to 9, characters A to Z
        $choices = (49..57) + (65..90) | ForEach-Object { [char]$_ }
        $i = 0
        $items = ($MenuItems | ForEach-Object { '[{0}]  {1}' -f $choices[$i++], $_ }) -join [Environment]::NewLine

        # display the menu and return the chosen option
        while ($true) {
            Clear-Host
            
            if ($header) { Write-Host $header -ForegroundColor Yellow }        
            Write-Host $items
            Show-AppInfo $appInfo
            Write-Host

            $answer = (Read-Host -Prompt 'Please make your choice').ToUpper()
            $index  = $choices.IndexOf($answer[0])

            if ($index -ge 0 -and $index -lt $MenuItems.Count) {
                return $MenuItems[$index]
            }
            else {
                Write-Warning "Invalid choice.. Please try again."
                Start-Sleep -Seconds 2
            }
        }
    }
    function Get-CredentialInfo ($cred, $credentialType)
    {
        $expired = "No"
        if(Get-IsExpired -date $cred.endDateTime){
            $expired = "Yes"
        }
        [pscustomobject]@{
            CredentialType = $credentialType
            KeyId = $cred.keyId
            Hint = $cred.hint
            Description = $cred.displayName
            StartDateTime = $cred.startDateTime
            EndDateTime = $cred.endDateTime
            KeyType = $cred.type
            Usage = $cred.usage
            Expired = $expired
        }
    }
    function Invoke-CredentialRollover ($appInfo, $keyId){
        $rolloverKey = $appInfo.Creds | Where-Object {$_.KeyId -eq $keyId}
        if($rolloverKey)
        {
            switch ($rolloverKey.CredentialType) {
                $credentialTypePassword { 
                    Add-Password -objectId $appInfo.objectId $appInfo.objectType
                    Remove-Password -objectId $appInfo.objectId -objectType $appInfo.objectType -keyId $rolloverKey.keyId                
                }            
                $credentialTypeKey {
                    $certFilePath = Read-Host -Prompt 'Enter the path to the certificate file'
                    $appInfo.keyCredentials = $appInfo.keyCredentials | Where-Object {$_.keyId -ne $keyId}
                    Add-Key -objectId $appInfo.objectId $appInfo.objectType -certFilePath $certFilePath
                }
            }
        }
        else {
            Write-Error "Secret or certificate with this KeyId was not found." -ErrorAction Stop
        }
    }

    function Get-IsExpired($date){
        return (Get-Date).Subtract($date) -gt 0
    }

    function Get-Passwords($appInfo) {
        return $appInfo.Creds | Where-Object {$_.CredentialType -eq $credentialTypePassword}
    }
    function Get-Certificates($appInfo) {
        return $appInfo.Creds | Where-Object {$_.CredentialType -eq $credentialTypeKey}
    }
    function Get-ExpiredCredentials($appInfo) {
        return $appInfo.Creds | Where-Object {(Get-IsExpired -date $_.endDateTime)}
    }
    function Get-ExpiredPasswords($appInfo) {
        return $appInfo.Creds | Where-Object {(Get-IsExpired -date $_.endDateTime) -and $_.CredentialType -eq $credentialTypePassword}
    }
    function Get-ExpiredCertificates($appInfo) {
        return $appInfo.Creds | Where-Object {(Get-IsExpired -date $_.endDateTime) -and $_.CredentialType -eq $credentialTypeKey}
    }
    function Remove-AppCredentials($appInfo, $selection)
    {    
        switch ($selection) {
            $menuRemoveAll          { $credsToRemove = $appInfo.Creds }
            $menuRemoveSecrets      { $credsToRemove = Get-Passwords -appInfo $appInfo }
            $menuRemoveCerts        { $credsToRemove = Get-Certificates -appInfo $appInfo }
            $menuRemoveExpiredAll   { $credsToRemove = Get-ExpiredCredentials -appInfo $appInfo }
            $menuRemoveExpiredSecrets{ $credsToRemove = Get-ExpiredPasswords -appInfo $appInfo }
            $menuRemoveExpiredCerts { $credsToRemove = Get-ExpiredCertificates -appInfo $appInfo }
        }

        # Remove passwords
        foreach($cred in $credsToRemove){
            if($cred.CredentialType -eq $credentialTypePassword){
                Remove-Password -objectId $appInfo.objectId -objectType $appInfo.objectType -keyId $cred.keyId
            }
        }
        # Remove keys
        Remove-Keys -appInfo $appInfo -credsToRemove $credsToRemove
    }
    function Remove-Keys ($appInfo, $credsToRemove) {
        $keyCredsToRemove = $credsToRemove | Where-Object {$_.CredentialType -eq $credentialTypeKey}
        if($keyCredsToRemove.Length -gt 0){
            foreach($cred in $keyCredsToRemove)
            {
                Write-Host ("Removing certificate ({0}) from {1} ({2})" -f $cred.keyId, $appInfo.ObjectType, $appInfo.ObjectId)
                $appInfo.keyCredentials = $appInfo.keyCredentials | Where-Object {$_.keyId -ne $cred.keyId}
            }
            if($appInfo.keyCredentials.Length -eq 0){ # Convert null to an empty array to generate a Graph compatible json
                $appInfo.KeyCredentials = @()
            }
            $body = @{keyCredentials = $appInfo.KeyCredentials} | ConvertTo-Json
            
            $uri = '/{0}/{1}' -f $appInfo.GraphObjectType, $appInfo.objectId
            Invoke-AADTGraph -uri $uri -body $body -method PATCH
        }    
    }
    function Remove-Password($objectId, $objectType, $keyId){
        Write-Host ("Removing client secret ({0}) from {1} ({2})" -f $keyId, $objectType, $objectId)
        switch ($objectType) {
            $objectTypeApplication { $graphObjectType = 'applications' }
            $objectTypeServicePrincipal { $graphObjectType = 'servicePrincipals' }
        }
        $uri = "/$graphObjectType/$objectId/removePassword"

        $body = @{keyId=$keyId} | ConvertTo-Json
        
        Invoke-AADTGraph -uri $uri -method POST -body $body
    }

    function Add-Password($objectId, $objectType){
        Write-Host ("Rolling over client secret for {0} ({1})" -f $objectType, $objectId)
        switch ($objectType) {
            $objectTypeApplication { $graphObjectType = 'applications' }
            $objectTypeServicePrincipal { $graphObjectType = 'servicePrincipals' }
        }
        $uri = "/$graphObjectType/$objectId/addPassword"

        $body = @{passwordCredential = @{displayName="Rollover"; endDateTime=((Get-Date).AddYears(1).ToString('s'))} } | ConvertTo-Json
        Write-Host $body
        Invoke-AADTGraph -uri $uri -method POST -body $body
    }

    function Add-Key($objectId, $objectType, $certFilePath){
        Write-Host ("Rolling over certificate for {0} ({1})" -f $objectType, $objectId)
        switch ($objectType) {
            $objectTypeApplication { $graphObjectType = 'applications' }
            $objectTypeServicePrincipal { $graphObjectType = 'servicePrincipals' }
        }
        $uri = "/$graphObjectType/$objectId/addKey"
        
        $cer = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certFilePath)
        $bin = $cer.GetRawCertData()
        $base64Value = [System.Convert]::ToBase64String($bin)
        $bin = $cer.GetCertHash()
        $base64Thumbprint = [System.Convert]::ToBase64String($bin)
        $keyId = [System.Guid]::NewGuid().ToString() 
        $newKeyCredential = @{
            customKeyIdentifier = $base64Thumbprint
            displayName = 'Rollover'
            keyId = $keyId
            type = 'AsymmetricX509Cert'
            usage = 'Verify'        
            key = $base64Value
        }
        $appInfo.KeyCredentials += $newKeyCredential

        $body = @{keyCredentials = $appInfo.KeyCredentials} | ConvertTo-Json
        $uri = '/{0}/{1}' -f $appInfo.GraphObjectType, $appInfo.objectId
        Invoke-AADTGraph -uri $uri -body $body -method PATCH
    }

    function Get-GraphSearchResults($type, $searchObjectId){
        $searchValue = "'$searchObjectId'"
        $uri = '/{0}?$filter=id eq {1} or id eq {1}' -f $type, $searchValue #Repear with or to avoid PowerShell Graph SDK throwing errors when there are no entries
        return Invoke-AADTGraph -uri $uri
    }
    function Get-GraphApp($searchObjectId){
        $result = Get-GraphSearchResults -type $graphObjectTypeApplications -searchObjectId $searchObjectId
        
        if($result.value.length -eq 1){
            $ObjectType = $objectTypeApplication
            $graphObjectType = $graphObjectTypeApplications
            $app = $result.value
        }
        else {
            $result = Get-GraphSearchResults -type $graphObjectTypeServicePrincipals -searchObjectId $searchObjectId
            if($result.value.length -eq 1){
                $ObjectType = $objectTypeServicePrincipal
                $graphObjectType = $graphObjectTypeServicePrincipals
                $app = $result.value
            }
            else {
                Write-Error "Object Id not found." -ErrorAction Stop
            }
        }
        $creds = @()
        foreach($cred in $app.passwordCredentials)
        {
            $creds += Get-CredentialInfo -cred $cred -credentialType $credentialTypePassword
        }
        foreach($cred in $app.keyCredentials)
        {
            $creds += Get-CredentialInfo -cred $cred -credentialType $credentialTypeKey
        }
        [pscustomobject]@{
            ObjectType = $objectType
            GraphObjectType = $graphObjectType
            Creds = $creds
            KeyCredentials = $app.keyCredentials #Until GraphAPI supports removing KeyCredentials, need to cache the KeyCredentials and use them with a PATCH 
            ObjectId = $app.id
            DisplayName = $app.displayName
            AppId = $app.appId
        }
    }


    $searchObjectId = Read-Host -Prompt 'Enter the ObjectId of the Application or Service Principal'

    $objectTypeApplication = 'Application'
    $objectTypeServicePrincipal = 'Service Principal'
    $graphObjectTypeApplications = 'applications'
    $graphObjectTypeServicePrincipals = 'servicePrincipals'
    $credentialTypePassword = 'Client secret'
    $credentialTypeKey = 'Certificate'
    $menuRemoveAll = 'Remove all certificates and secrets'
    $menuRemoveCerts = 'Remove all certificates'
    $menuRemoveSecrets = 'Remove all secrets'
    $menuRemoveExpiredAll = 'Remove expired certificates and secrets'
    $menuRemoveExpiredCerts = 'Remove expired certificates'
    $menuRemoveExpiredSecrets = 'Remove expired secrets'
    $menuRolloverCred = 'Rollover a certificate or secret'
    $menuQuit = 'Quit'


    if(![guid]::TryParse($searchObjectId, $([ref][guid]::Empty)))
    {
        Write-Error "Invalid object identifier format" -ErrorAction Stop
    }
    else
    {
        $appInfo = Get-GraphApp -searchObjectId $searchObjectId
        if(!$appInfo){
            Write-Error "Application or ServicePrincipal with this ObjectId was not found"
        }
        
        ##TODO Filter to show options applicable for the selected app
        if($appInfo.Creds -and $appInfo.Creds.Length -gt 0)
        {
            Show-AppInfo $appInfo

            $hasSecrets = ($appInfo.Creds | Where-Object {$_.CredentialType -eq $credentialTypePassword}).Length -gt 0
            $hasCerts = ($appInfo.Creds | Where-Object {$_.CredentialType -eq $credentialTypePassword}).Length -gt 0
            $menu = @($menuRemoveAll)
            if(Get-Passwords -appInfo $appInfo){ $menu += menuRemoveSecrets}
            if(Get-Certificates -appInfo $appInfo){ $menu += $menuRemoveCerts}
            if(Get-ExpiredCredentials -appInfo $appInfo){ $menu += $menuRemoveExpiredAll}
            if(Get-ExpiredPasswords -appInfo $appInfo){ $menu += $menuRemoveExpiredSecrets}
            if(Get-ExpiredCertificates -appInfo $appInfo){ $menu += $menuRemoveExpiredCerts}
            $menu += $menuRolloverCred, $menuQuit
            $title = "Manage credentials for {0} {1} ({2})" -f $appInfo.ObjectType, $appInfo.DisplayName, $appInfo.ObjectId
            $selection = Show-Menu -appInfo $appInfo  -MenuItems $menu -Title $title
            
            switch ($selection) {
                {$_ -in $menuRemoveAll, $menuRemoveCerts, $menuRemoveSecrets, $menuRemoveExpiredAll, $menuRemoveExpiredCerts, $menuRemoveExpiredSecrets}{
                    Remove-AppCredentials -appInfo $appInfo -selection $selection
                }
                $menuRolloverCred {
                    $rolloverKeyId = Read-Host -Prompt 'Enter the KeyId of the client secret or certificate to be rolled over'
                    Invoke-CredentialRollover -appInfo $appInfo -keyId $rolloverKeyId
                }
            }
        }
        else {
            $message = "{0} does have any client secrets or certificates." -f $appInfo.ObjectType
            Write-Error $message
        }
    }
}