<#
.SYNOPSIS
    Find Users with Admin Roles that are not registered for MFA
.DESCRIPTION
    Find Users with Admin Roles that are not registered for MFA by evaluating their authentication methods registered for MFA and their sign-in activity.
.PARAMETER IncludeSignIns
    Include Sign In log activity -  Note this can cause the query to run slower in larger active environments
.EXAMPLE
    Find-UnprotectedUsersWithAdminRoles
    Enumrate users with role assignments including their sign in activity
.EXAMPLE
    Find-UnprotectedUsersWithAdminRoles -includeSignIns:$false
    Enumerate users with role assignments including their sign in activity
.INPUTS
    Inputs to this cmdlet (if any)
.OUTPUTS
    Output from this cmdlet (if any)
.NOTES
     - Eligible users for roles may not have active assignments showing in their directoryrolememberships, but they have the potential to elevate to assigned roles
     - Large amounts of role assignments may take time process.
     - Must be connected to MS Graph with appropriate scopes for reading user, role, an sign in information and selected the beta profile before running.
      --  Connect-MgGraph
      --  Select-MgProfile -name Beta

#>
function Find-UnprotectedUsersWithAdminRoles {
    [CmdletBinding(DefaultParameterSetName = 'Parameter Set 1',
        PositionalBinding = $false,
        HelpUri = 'http://www.microsoft.com/',
        ConfirmImpact = 'Medium')]
    [Alias()]
    [OutputType([String])]

    Param (
        [switch]
        $IncludeSignIns
    )
    
    begin {

        if ($null -eq (Get-MgContext)) {
            Write-Error "Please Connect to MS Graph API with the Connect-MgGraph cmdlet from the Microsoft.Graph.Authentication module first before calling functions!" -ErrorAction Stop
        }
        else {
            

            if ((Get-MgProfile).Name -eq 'v1.0') {
                Write-Error ("Current MGProfile is set to v1.0, and some cmdlets may need to use the beta profile.   Run Select-MgProfile -Name beta to switch to beta API profile") -ErrorAction Stop
            }

        }

    }
    
    process {
       
        $usersWithRoles = Get-UsersWithRoleAssignments
       
        Write-Verbose ("Checking {0} users with roles..." -f $usersWithRoles.count)

        $checkedUsers = @()

        foreach ($user in $usersWithRoles) {

           
            $userObject = $null
            $userObject = get-mguser -userID $user.PrincipalId -Property signInActivity, UserPrincipalName, Id
            Write-Verbose ("Evaluating {0} with role assignments...." -f $userObject.Id)

            if ($Null -ne $userObject) {
                $UserAuthMethodStatus = Get-UserMfaRegisteredStatus -UserId $userObject.UserPrincipalName

                $checkedUser = [ordered] @{}
                $checkedUser.UserID = $userObject.Id
                $checkedUser.UserPrincipalName = $userObject.UserPrincipalName
            
                If ($null -eq $userObject.signInActivity.LastSignInDateTime) {
                    $checkedUser.LastSignInDateTime = "Unknown"
                    $checkedUser.LastSigninDaysAgo = "Unknown"
                }
                else {
                    $checkedUser.LastSignInDateTime = $userObject.signInActivity.LastSignInDateTime
                    $checkedUser.LastSigninDaysAgo = (New-TimeSpan -Start $checkedUser.LastSignInDateTime -End (get-date)).Days
                }
                $checkedUser.DirectoryRoleAssignments = $user.RoleName
                $checkedUser.DirectoryRoleAssignmentType = $user.AssignmentType
                $checkedUser.DirectoryRoleAssignmentCount = $user.RoleName.count
                $checkedUser.IsMfaRegistered = $UserAuthMethodStatus.isMfaRegistered

                if ($includeSignIns -eq $true) {
                    $signInInfo = get-UserSignInSuccessHistoryAuth -userId $checkedUser.UserId

                    $checkedUser.SuccessSignIns = $signInInfo.SuccessSignIns
                    $checkedUser.MultiFactorSignIns = $signInInfo.MultiFactorSignIns
                    $checkedUser.SingleFactorSignIns = $signInInfo.SingleFactorSignIns
                    $checkedUser.RiskySignIns = $signInInfo.RiskySignIns
                }
                else
                {
                    $checkedUser.SuccessSignIns = "Skipped"
                    $checkedUser.MultiFactorSignIns = "Skipped"
                    $checkedUser.SingleFactorSignIns = "Skipped"
                    $checkedUser.RiskySignIns = "Skipped"
                }
                $checkedUsers += ([pscustomobject]$checkedUser)
            }
        }
        

    }
    
    end {
        Write-Output $checkedUsers
    }
}

function Get-UserMfaRegisteredStatus ([string]$UserId) {

    $mfaMethods = @("#microsoft.graph.fido2AuthenticationMethod", "#microsoft.graph.softwareOathAuthenticationMethod", "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod", "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod", "#microsoft.graph.phoneAuthenticationMethod")

    $authMethods = (Get-MgUserAuthenticationMethod -UserId $UserId).AdditionalProperties."@odata.type"

    $isMfaRegistered = $false
    foreach ($mfa in $MfaMethods) { if ($authmethods -contains $mfa) { $isMfaRegistered = $true } }
    
    $results = @{}
    $results.IsMfaRegistered = $isMfaRegistered
    $results.AuthMethodsRegistered = $authMethods

    Write-Output ([pscustomobject]$results)

}

function get-UserSignInSuccessHistoryAuth ([string]$userId) {

    $signinAuth = @{}
    $signinAuth.UserID = $userId
    $signinAuth.SuccessSignIns = 0
    $signinAuth.MultiFactorSignIns = 0
    $signinAuth.SingleFactorSignIns = 0
    $signInAuth.RiskySignIns = 0

    $filter = ("UserId eq '{0}' and status/errorCode eq 0" -f $userId)
    Write-Debug $filter
    $signins = Get-MgAuditLogSignIn -Filter $filter -all:$True
    Write-Debug $signins.count

    if ($signins.count -gt 0) {

        $signinAuth.SuccessSignIns = $signins.count
        $groupedAuth = $signins | Group-Object -Property AuthenticationRequirement

        $MfaSignInsCount = 0
        $MfaSignInsCount = $groupedAuth | Where-Object -FilterScript {$_.Name -eq 'multiFactorAuthentication'} | Select-Object -ExpandProperty count
        if ($null -eq $MfaSignInsCount)
        {
            $MfaSignInsCount = 0
        }
        $signinAuth.MultiFactorSignIns = $MfaSignInsCount

        $singleFactorSignInsCount = 0
        $singleFactorSignInsCount = $groupedAuth | Where-Object -FilterScript {$_.Name -eq 'singleFactorAuthentication'} | Select-Object -ExpandProperty count


        if ($null -eq $singleFactorSignInsCount)
        {
            $singleFactorSignInsCount = 0
        }
        $signinAuth.SingleFactorSignIns = $singleFactorSignInsCount

        $signInAuth.RiskySignIns = ($signins | Where-Object -FilterScript { $_.RiskLevelDuringSignIn -ne 'none' } | Measure-Object | Select-Object -ExpandProperty Count)

    }

    Write-Output ([pscustomobject]$signinAuth)
}

function Get-UsersWithRoleAssignments()
{
    $uniquePrincipals = $null
    $usersWithRoles = $Null
    $groupsWithRoles = $null
    $servicePrincipalsWithRoles = $null
    $roleAssignments = @()
    $activeRoleAssignments = $null
    $eligibleRoleAssignments = $null
    $AssignmentSchedule =@()

    Write-Verbose "Retrieving Active Role Assignments..."
    $activeRoleAssignments = Get-MgRoleManagementDirectoryRoleAssignmentSchedule -All:$true|Add-Member -MemberType NoteProperty -Name AssignmentScope -Value "Active" -Force -PassThru
    Write-Verbose ("{0} Active Role Assignments..." -f $activeRoleAssignments.count)
    $AssignmentSchedule += $activeRoleAssignments
    

    Write-Verbose "Retrieving Eligible Role Assignments..."
    $eligibleRoleAssignments = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All:$true|Add-Member -MemberType NoteProperty -Name AssignmentScope -Value "Eligible" -Force -PassThru
    Write-Verbose ("{0} Eligible Role Assignments..." -f $eligibleRoleAssignments.count)
    $AssignmentSchedule += $eligibleRoleAssignments

    Write-Verbose ("{0} Total Role Assignments to all principals..." -f $AssignmentSchedule.count)
    $uniquePrincipals = $AssignmentSchedule.PrincipalId|Get-Unique
    Write-Verbose ("{0} Total Role Assignments to unique principals..." -f $uniquePrincipals.count)
    
    foreach ($assignment in ($AssignmentSchedule))
    {
        $roleAssignment = @{}
        $roleAssignment.PrincipalId = $assignment.PrincipalId
        $directoryObject = Get-MgDirectoryObject -DirectoryObjectId $assignment.PrincipalId
        $roleAssignment.PrincipalType = $directoryObject.AdditionalProperties."@odata.type".split('.')[2] 
        $roleAssignment.AssignmentType = $assignment.AssignmentScope
        $roleAssignment.RoleDefinitionId = $assignment.RoleDefinitionId
        $roleAssignment.RoleName = Get-MgRoleManagementDirectoryRoleDefinition -UnifiedRoleDefinitionId $assignment.RoleDefinitionId|Select-Object -ExpandProperty displayName
        $roleAssignments += ([pscustomobject]$roleAssignment)

    }
    

    Write-Verbose ("{0} Total Role Assignments" -f $roleAssignments.count)
    $usersWithRoles = $roleAssignments|Where-Object -FilterScript {$_.PrincipalType -eq 'user'}
    $groupsWithRoles = $roleAssignments|Where-Object -FilterScript {$_.PrincipalType -eq 'group'}
    $servicePrincipalsWithRoles = $roleAssignments|Where-Object {$_.PrincipalType -eq 'servicePrincipal'}

    if ($groupsWithRoles.count -gt 0)
    {
        Write-warning ("Groups with Assigned Roles! -  Groups with roles are not currently enumerated for users who are members of the role assignable groups in the process!")
    }


    foreach ($type in ($roleAssignments|Group-Object PrincipalType))
    {
        Write-Verbose ("{0} assignments to {1} type" -f $type.count, $type.name)
    }



    Write-Output $usersWithRoles
}