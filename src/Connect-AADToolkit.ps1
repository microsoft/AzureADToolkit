<#
.SYNOPSIS
    Connect the Azure AD Toolkit module to Azure AD tenant.
.DESCRIPTION
    This command will connect Microsoft.Graph to your Azure AD tenant.
.EXAMPLE
    PS C:\>Connect-AADToolkit
    Connect to home tenant of authenticated user.
#>
function Connect-AADToolkit {
    Connect-MgGraph -Scopes 'Application.ReadWrite.All'
}