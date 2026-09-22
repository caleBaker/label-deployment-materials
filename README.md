# Label Deployment Materials

Builds a SharePoint Online communication site for sensitivity label training via Microsoft Graph.

## Contents

- `New-LabelTrainingSite.ps1` – PowerShell 7 script that signs in with delegated Graph permissions, creates the site, uploads images, and generates the Overview, Labels, per-label, and Troubleshooting pages.
- `overview-outline.md` / `troubleshooting-outline.md` – Markdown sources for the Overview and Troubleshooting & FAQ pages.
- `labels-definition.json` – Label names, descriptions, correct/incorrect uses, and screenshot references.
- `Backgrounds/`, `Screenshots/` – Images uploaded to the site.

## Usage

```powershell
pwsh ./New-LabelTrainingSite.ps1
```

Requires PowerShell 7+ and `Microsoft.Graph.Authentication` (installed automatically if missing). Run logs are written to `Logs/`.
