<#
.SYNOPSIS
    Disconnects the current session of the Azure AD Toolkit module.
.DESCRIPTION
    This command will disconnect Microsoft.Graph from the current session. Required when switching between different tenants.
.EXAMPLE
    PS C:\>Disconnect-AADToolkit
    Connect to home tenant of authenticated user.
#>
function Disconnect-AADToolkit {
    Disconnect-MgGraph
}