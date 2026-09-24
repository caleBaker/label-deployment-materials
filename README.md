# Automated Label Training SharePoint Site

Builds a SharePoint Online communication site for sensitivity label training via Microsoft Graph.

## Structure

```
label-training-sharepoint-site/
│
├── New-LabelTrainingSite.ps1
│
├── dependencies/
│   ├── labels-definition.json - Template provided
│   ├── overview-outline.md - Template provided
│   └── troubleshooting-outline.md - Template provided
│
├── backgrounds/
│   ├── home-background.jpg
│   ├── labels-background.jpg
│   └── troubleshooting-background.jpg
│
├── screenshots/
│   └── Place your screenshots here and reference in dependencies/labels-definition.json
│
└── logs/ - New logs directory will be created on initial run
    └── New-LabelTrainingSite_yyyyMMdd-HHmmss.txt
```

## Requirements

### PowerShell and modules

- PowerShell 7.0 or later.
- `Microsoft.Graph.Authentication`, the only module the script uses. It provides `Connect-MgGraph` and `Invoke-MgGraphRequest`, and every Graph call goes through the latter. You don't need PnP.PowerShell, SharePoint CSOM, or any other `Microsoft.Graph.*` module.

The script checks the module on every run:

- If it's missing, the script installs the latest version from the PowerShell Gallery for the current user.
- If several versions are installed, the script warns you, because side-by-side `Microsoft.Graph.*` versions load conflicting assemblies. Keep only the newest:
  ```powershell
  Uninstall-Module -Name Microsoft.Graph.Authentication -RequiredVersion <old version>
  ```
- If a newer version is available in the gallery, the script warns you and suggests `Update-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser`.

The script always imports the highest installed version.

To install the module yourself before the first run:

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

### Microsoft Graph permissions

The script signs in interactively with **delegated** permissions. It stops if the Graph session uses app-only authentication. It requests these scopes:

| Scope | Used for |
| --- | --- |
| `Sites.Create.All` | Creating the communication site (`POST /beta/sites`) |
| `Sites.ReadWrite.All` | Creating and updating pages, lists and the document library |
| `Files.ReadWrite.All` | Uploading background images and screenshots |

An administrator may need to consent to `Sites.Create.All` once for the Microsoft Graph PowerShell application. If a requested scope is missing from the token after sign-in, the script logs a warning, and any operation that needs that scope fails.

### Account

- The signed-in account must be allowed to create SharePoint sites in the tenant.
- The signed-in account becomes the site owner. With delegated permissions, Graph can only reach sites the signed-in user can access, so the script rejects an `-OwnerUpn` that doesn't match the signed-in account. To add other owners, add them in Site permissions after the run.

## Usage

```powershell
./New-LabelTrainingSite.ps1 -SiteTitle 'Sensitivity Label Knowledge Base' -SiteAlias 'sensitivitylabels' -TenantHostName 'companyname.sharepoint.com' -OwnerUpn 'placeholder@companyname.com'
```

The script reads the label definitions and both outlines from `dependencies/` by default. Pass `-LabelDefinitionPath`, `-OverviewContentPath` or `-TroubleshootingContentPath` to use other files; relative paths are resolved against the script folder.

## License

[MIT](LICENSE)
