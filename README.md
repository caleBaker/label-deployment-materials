# Label Deployment Materials

Builds a SharePoint Online communication site for sensitivity label training via Microsoft Graph.

## Structure

```
label-deployment-materials/
│
├── New-LabelTrainingSite.ps1
├── labels-definition.json - Template provided
├── overview-outline.md - Template provided
├── troubleshooting-outline.md - Template provided
│
├── Backgrounds/
│   ├── home-background.jpg
│   ├── labels-background.jpg
│   └── troubleshooting-background.jpg
│
├── Screenshots/
│   └── Place your screenshots here and reference in labels-definition.json
│
└── Logs/ - New Logs directory will be created on initial run
    └── New-LabelTrainingSite_yyyyMMdd-HHmmss.txt
```

## Usage

```powershell
./New-LabelTrainingSite.ps1 -SiteTitle 'Sensitvity Label Knowledge Base' -SiteAlias 'sensitivitylabels' -TenantHostName 'companyname.sharepoint.com' -OwnerUpn 'placeholder@companyname.com' -LabelDefinitionPath ./labels-definition.json
```

Requires PowerShell 7+ and `Microsoft.Graph.Authentication` (installed automatically if missing).

## License

[MIT](LICENSE)
