# Azure AD Toolkit

The Azure AD Toolkit is a PowerShell module that providers helper cmdlets to manage the credentials of your application or service principal.

## Installing the module
```powershell
    Install-Module AzureADToolkit
```

## Using the module

### Connecting to your tenant
Connect to the user's default tenant.
```powershell
    Connect-AADToolkit    
```
Specify the Tenant ID if the user signing in has access to multiple Azure Active Directory tenants.
```powershell
    Connect-AADToolkit -TenantId 344b8aab-389c-4e4a-8fa1-4c1ae2c0a60d
```

### Exporting a list of all the Service Principals and Applications having credentials
```powershell
    Get-AADToolkitApplicationCredentials | Export-Csv -Path '.\AppPermissions.csv'  -NoTypeInformation
```

### Interactively removing and rolling over the certificates and secrets of a Service Principal or Application
This command provides a menu drive interface to view the credentials of an application and allows the user to remove or roll them over.
```powershell
    Update-AADToolkitApplicationCredentials
```

### Exporting a list of Service Principals and Applications with privilege scores (requires external module to generate Excel Workbook)
It is recommended that you use `Connect-MgGraph -Scopes Application.Read.All` to connect to Microsoft Graph PowerShell for this report. The minimum administrative role necessary to consent to this permission is Application Administrator.

Connect to Microsoft Graph PowerShell with the appropriate permissions:
```powershell
    Install-Module ImportExcel
    Install-Module Microsoft.Graph
    Connect-MgGraph -Scopes Application.Read.All
```

This example will export the report to an Excel workbook:
```
    Build-AADToolkitAppConsentGrantReport -ReportOutputType ExcelWorkbook -ExcelWorkbookPath C:\temp\export.xlsx
```

This example will retrieve the data and store it in PowerShell objects instead of exporting to Excel:
```
    Build-AADToolkitAppConsentGrantReport -ReportOutputType PowerShellObjects
```

### List all users with admin roles and their strong authentication status

Find Users with Admin Roles that are not registered for MFA by evaluating their authentication methods registered for MFA and their sign-in activity.

```
   Connect-MgGraph -Scopes RoleManagement.Read.Directory,UserAuthenticationMethod.Read.All,AuditLog.Read.All,User.Read.All,Group.Read.All,Application.Read.All
   Select-MgProfile -name Beta
   Find-AADToolkitUnprotectedUsersWithAdminRoles -Verbose -IncludeSignIns | Export-Csv ./admins.csv
```


### Disconnecting from your tenant
```powershell
    Disconnect-AzureADToolkit
```

## Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.opensource.microsoft.com.

When you submit a pull request, a CLA bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft 
trademarks or logos is subject to and must follow 
[Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general).
Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship.
Any use of third-party trademarks or logos are subject to those third-party's policies.
