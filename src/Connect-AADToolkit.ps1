<#
.SYNOPSIS
    Connect the Azure AD Toolkit module to Azure AD tenant.
.DESCRIPTION
    This command will connect Microsoft.Graph to your Azure AD tenant.
.EXAMPLE
    PS C:\>Connect-AADToolkit
    Connect to home tenant of authenticated user.
.EXAMPLE
    PS C:\>Connect-AADToolkit -TenantId 3043-343434-343434
    Connect to a specific Tenant
#>
function Connect-AADToolkit {
    param(
        [Parameter(Mandatory = $false)]
        [string] $TenantId = 'common'
    )    
    Connect-MgGraph -Scopes 'Application.ReadWrite.All' -TenantId $TenantId
}