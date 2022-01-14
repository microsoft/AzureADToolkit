<#
.SYNOPSIS
    Gets user sign-in activity asscoaited with external tenants

.DESCRIPTION
    Gets two types of user sign-in activity associated with external tenants, depending on mode selected:

        * Lists sign-in events of external tenant IDs accessed by local users
        * List sign-in events of external tenant IDs of external users accessing local tenant

    Has a mode to show summary statistics for weach external tenant, in each mode.

.EXAMPLE
    Get-AADtoolkitExternalTenantUserActivity -LocalUsersAccessingExternalTenant

    Gets all available sign-in events for local users accessing resources in an external teanant. 
    Lists by unique external tenant.

.EXAMPLE
    Get-AADtoolkitExternalTenantUserActivity -LocalUsersAccessingExternalTenant -SummaryStats

    Provides a summary of the number of sign-ins, unique users and unique resources per external tenant, for local
    users accessing resources in an external tenant.

.EXAMPLE
    Get-AADtoolkitExternalTenantUserActivity -ExternalUsersAccessingLocalTenant

    Gets all available sign-in events for external users accessing resources in the local teanant. 
    Lists by unique external tenant.

.EXAMPLE
    Get-AADtoolkitExternalTenantUserActivity -ExternalUsersAccessingLocalTenant -SummaryStats

    Provides a summary of the number of sign-ins, unique users and unique resources per external tenant, for external
    users accessing resources in the local tenant.

.NOTES
    THIS CODE-SAMPLE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED 
    OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR 
    FITNESS FOR A PARTICULAR PURPOSE.

    This sample is not supported under any Microsoft standard support program or service. 
    The script is provided AS IS without warranty of any kind. Microsoft further disclaims all
    implied warranties including, without limitation, any implied warranties of merchantability
    or of fitness for a particular purpose. The entire risk arising out of the use or performance
    of the sample and documentation remains with you. In no event shall Microsoft, its authors,
    or anyone else involved in the creation, production, or delivery of the script be liable for 
    any damages whatsoever (including, without limitation, damages for loss of business profits, 
    business interruption, loss of business information, or other pecuniary loss) arising out of 
    the use of or inability to use the sample or documentation, even if Microsoft has been advised 
    of the possibility of such damages, rising out of the use of or inability to use the sample script, 
    even if Microsoft has been advised of the possibility of such damages.   


#>
function Get-AADtoolkitExternalTenantUserActivity {

    [CmdletBinding(DefaultParameterSetName="LocalUser")]
    param(

        #List external tenant IDs accessed by local users
        [Parameter(Position=0,ParameterSetName="LocalUser")]
        [switch]$LocalUsersAccessingExternalTenant,

        #List external tenant IDs of external users accessing local tenant
        [Parameter(Position=1,ParameterSetName="ExternalUser")]
        [switch]$ExternalUsersAccessingLocalTenant,

        #Show summary statistics by tenant
        [switch]$SummaryStats

        )
    
    begin {
        
        #Connection and profile check

        Write-Verbose -Message "$(Get-Date -f T) - Checking connection..."

        if ($null -eq (Get-MgContext)) {

            Write-Error "$(Get-Date -f T) - Please connect to MS Graph API with the Connect-AADToolkit cmdlet!" -ErrorAction Stop
        }
        else {

            Write-Verbose -Message "$(Get-Date -f T) - Checking profile..."

            if ((Get-MgProfile).Name -eq 'v1.0') {

                Write-Error "$(Get-Date -f T) - Current MGProfile is set to v1.0, and some cmdlets may need to use the beta profile. Run 'Select-MgProfile -Name beta' to switch to beta API profile" -ErrorAction Stop
            }

        }

        if (!$PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent) {

            Write-Host "NOTE: $(Get-Date -f T) - This process may take a while to complete, depending on the size of the environment. Please run with -Verbose switch for detailed progress output."
        }

        Write-Verbose -Message "$(Get-Date -f T) - Connection and profile OK"

    }
    
    process {

        #Get filtered sign-in logs

        if ($LocalUsersAccessingExternalTenant) {

            Write-Verbose -Message "$(Get-Date -f T) - Getting external tenant IDs accessed by local users"

            if (!$PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent) {

                Write-Host "NOTE: $(Get-Date -f T) - Getting external tenant IDs accessed by local users."
            }

            $SignIns = Get-MgAuditLogSignIn -Filter ("ResourceTenantId ne '{0}'" -f (Get-MgContext).TenantId) -all:$True | Group-Object ResourceTenantID

        }
        elseif ($ExternalUsersAccessingLocalTenant) {

            Write-Verbose -Message "$(Get-Date -f T) - Getting external tenant IDs for external users accessing local tenant"

            if (!$PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent) {

                Write-Host "NOTE: $(Get-Date -f T) - Getting external tenant IDs for external users accessing local tenant"
            }


            $SignIns = Get-MgAuditLogSignIn -Filter ("HomeTenantId ne '{0}' and TokenIssuerType eq 'AzureAD'" -f (Get-MgContext).TenantId) -all:$True | Group-Object HomeTenantID

        }
        else {

            Write-Error "$(Get-Date -f T) - Please specific one of the following switches: -LocalUsersAccessingExternalTenant or -ExternalUsersAccessingLocalTenant." -ErrorAction Stop

        }

        #Analyse sign-in logs

        Write-Verbose -Message "$(Get-Date -f T) - Checking for sign-ins..."

        if ($SignIns) {
            
            Write-Verbose -Message "$(Get-Date -f T) - Sign-ins obtained"
            Write-Verbose -Message "$(Get-Date -f T) - Iterating Sign-ins..."

            foreach ($TenantID in $SignIns) {

                #Provide summary

                if ($SummaryStats) {

                    Write-Verbose -Message "$(Get-Date -f T) - Creating summary stats for external tenant - $($TenantId.Name)"

                    $Analysis = [pscustomobject]@{

                        ExternalTenantId = $TenantId.Name
                        SignIns = $TenantId.Count
                        UniqueUsers = ($TenantID.Group | select UserId -unique).count
                        #UniqueApps = ($TenantID.Group | select AppId -unique).count
                        UniqueResources = ($TenantID.Group | select ResourceId -unique).count


                    }

                    Write-Verbose -Message "$(Get-Date -f T) - Adding stats for $($TenantId.Name) to total analysis object"

                    [array]$TotalAnalysis += $Analysis

                }
                else {

                    #Get individual events by external tenant

                    Write-Verbose -Message "$(Get-Date -f T) - Getting individual sign-in events for external tenant - $($TenantId.Name)"

                    $TenantID.group | Select @{n='ExternalTenantId';e={$TenantId.name}},UserDisplayName,UserPrincipalName,UserId,UserType,CrossTenantAccessType,AppDisplayName,AppId,`
                                             ResourceDisplayName,ResourceId,@{n='SignInId';e={$_.Id}},CreatedDateTime

                }

            }

        }
        else {

            Write-Host "NOTE: $(Get-Date -f T) - No sign-ins matching the selected criteria found."

        }

        if ($SummaryStats) {

            #Show array of summary objects for each external tenant

            Write-Verbose -Message "$(Get-Date -f T) - Displaying total analysis object"

            $TotalAnalysis | Sort SignIns -Descending

        }


    }
       
}