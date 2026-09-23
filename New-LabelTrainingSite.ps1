#Requires -Version 7.0

<#
.SYNOPSIS
    Creates and configures a SharePoint Online communication site for sensitivity
    label deployment training, using raw Microsoft Graph REST calls
    (Invoke-MgGraphRequest) and delegated (interactive) permissions.

.DESCRIPTION
    The script performs six stages:

      1. Module check  - verifies Microsoft.Graph.Authentication, the only module needed,
                         and installs it for the current user if it is missing.
      2. Sign-in       - Connect-MgGraph with delegated scopes (interactive user sign-in).
                         No app registration, no client credentials, no app-only path.
      3. Content       - reads and validates the two markdown outlines and the label JSON
                         before anything in the tenant is touched.
      4. Site          - creates a blank communication site (template
                         'sitepagepublishing') if it does not already exist, then polls
                         the asynchronous provisioning operation until the site answers.
      5. Images        - uploads the page background images to a 'Backgrounds' folder and
                         the label screenshots to an images folder in the site's default
                         document library, reading each file's pixel dimensions so the
                         image web parts are laid out at the right aspect ratio.
      6. Pages         - builds the Overview (home) page from a markdown outline, a
                         Labels parent page, one page per label, and a Troubleshooting &
                         FAQ page from a second markdown outline, all laid out with
                         SharePoint web parts (text, image, divider). Every page uses
                         the plain title layout; the Overview, Labels and Troubleshooting
                         pages open with a full-width image web part showing an uploaded
                         background image, while the label pages have no banner at all.

    The run ends with a printed summary: site URL, whether the site was created or only
    reconciled, the image upload count, every page written, the top-navigation links to
    add by hand, and the log file path.

    KNOWN PLATFORM LIMITATIONS (surfaced, not silently worked around)

      * Site creation is beta-only. Microsoft Graph v1.0 is explicitly read-only for
        site resources ("no ability to create new sites"), so the script POSTs to
        /beta/sites for that one operation and uses v1.0 for everything else.

      * Graph's page API supports a fixed list of web parts. Hero, Quick Links and
        Document Embed carry undocumented 'data' schemas, so this script builds pages
        from text, image and divider web parts, whose payload shapes are documented,
        rather than emitting web part JSON that the service would reject.

      * Graph cannot change a site's welcome page. The Overview content is therefore
        written to the communication site's existing home page (Home.aspx) by default.

      * Graph cannot edit a communication site's top navigation. The script prints the
        three links to add (Overview, Sensitivity labels, Troubleshooting & FAQ) in its
        summary so they can be added once by hand in Site settings > Navigation.

      * SharePoint pages are flat. The per-label pages sit beside the Labels page in the
        Pages library, not underneath it; they read as subpages because the Labels page
        links to each one and because the top navigation nests them under Labels.

.PARAMETER SiteTitle
    Display title of the site, e.g. 'Sensitivity Label Training'.

.PARAMETER SiteAlias
    URL alias of the site, i.e. the <alias> in https://<tenant>/sites/<alias>.

.PARAMETER TenantHostName
    SharePoint host name, e.g. 'contoso.sharepoint.com'.

.PARAMETER OwnerUpn
    UPN (email) of the site owner, resolved by Graph at creation time. Defaults to the
    account that signs in. It must be the signed-in user: with delegated permissions,
    Graph can only resolve sites the signed-in user can access, so a site created for
    somebody else cannot be polled or configured by this script and the provisioning
    wait times out. A different value is rejected before anything is created. To hand
    the site to another owner, add them in Site permissions after the run.

.PARAMETER SiteDescription
    Descriptive text stored on the site.

.PARAMETER Locale
    Site language, default 'en-US'.

.PARAMETER OverviewContentPath
    Markdown file whose '##' sections become the Overview page sections.
    Defaults to 'dependencies/overview-outline.md' beside this script.

.PARAMETER TroubleshootingContentPath
    Markdown file whose '##' sections become the Troubleshooting & FAQ page sections.
    Lines starting with 'Q:' and 'A:' are rendered as question/answer pairs.
    Defaults to 'dependencies/troubleshooting-outline.md' beside this script.

.PARAMETER HomeBackgroundPath
    Image shown in a full-width image web part at the top of the Overview (home) page.
    Defaults to 'backgrounds/home-background.jpg' beside this script. Pass an empty
    string to build the page without a background image.

.PARAMETER LabelsBackgroundPath
    Image shown in a full-width image web part at the top of the Labels page.
    Defaults to 'backgrounds/labels-background.jpg' beside this script. Pass an empty
    string to build the page without a background image.

.PARAMETER TroubleshootingBackgroundPath
    Image shown in a full-width image web part at the top of the Troubleshooting & FAQ
    page. Defaults to 'backgrounds/troubleshooting-background.jpg' beside this script.
    Pass an empty string to build the page without a background image.

.PARAMETER LabelDefinitionPath
    JSON file describing the labels. Defaults to 'dependencies/labels-definition.json'
    beside this script; a relative path is resolved against the script folder. Shape:
        [
          {
            "Name":          "Highly Confidential",
            "Group":         "Confidential",      // omit, null or "" for an individual label
            "VisibleTo":     "Finance",           // who can see and apply the label
            "Description":   "Applies to ...",
            "CorrectUses":   ["Draft earnings statements", "M&A working papers"],
            "IncorrectUses": ["Published annual report", "Cafeteria menu"],
            "QuickNotes":    ["Encrypts the file; external recipients must sign in"],
            "Screenshots":   [
              "./screenshots/hc-word.png",
              {
                "Path":        "./screenshots/hc-outlook.png",
                "Caption":     "The Sensitivity button on the Outlook ribbon",
                "Description": "Pick **Highly Confidential** from the menu. Outlook applies it to the message and every attachment."
              }
            ]
          }
        ]
    Name, VisibleTo and Description are required; the list properties may be omitted.
    A screenshot entry is either a path string or an object with a required Path plus
    optional Caption (plain text shown directly beneath the image) and Description
    (a paragraph, or an array of paragraphs, shown below the caption; accepts the same
    inline markdown as the outlines: links, **bold**, *italic*, `code`).
    Screenshot paths are resolved against the script folder, not the JSON file's folder.
    The file must contain at least one label.

.PARAMETER OverviewPageName
    File name of the Overview page. Default 'Home.aspx' (the communication site home
    page). Graph cannot set a different page as the site's welcome page.

.PARAMETER LabelsPageName
    File name of the Labels parent page. Default 'Sensitivitylabels.aspx'.

.PARAMETER TroubleshootingPageName
    File name of the Troubleshooting & FAQ page. Default 'Troubleshooting.aspx'.

.PARAMETER LogDirectory
    Folder for the run log. Defaults to a 'Logs' folder beside this script.

.PARAMETER ProvisioningTimeoutSeconds
    How long to wait for asynchronous site provisioning. Default 600.

.PARAMETER Force
    Delete and recreate pages that already exist instead of updating them in place.

.EXAMPLE
    ./New-LabelTrainingSite.ps1 -SiteTitle 'Sensitivity Label Training' `
        -SiteAlias 'label-training' -TenantHostName 'contoso.sharepoint.com'

    Creates the site, owned by the account that signs in, and builds every page from
    dependencies/labels-definition.json, dependencies/overview-outline.md and
    dependencies/troubleshooting-outline.md.

.EXAMPLE
    ./New-LabelTrainingSite.ps1 -SiteTitle 'Sensitivity Label Knowledge Base' `
        -SiteAlias 'sensitivitylabels' -TenantHostName 'contoso.sharepoint.com' `
        -OwnerUpn 'admin@contoso.com' -Force

    Rebuilds existing pages from the same label JSON. -OwnerUpn is only consulted when
    the site has to be created, and then it must match the signed-in account.

.EXAMPLE
    ./New-LabelTrainingSite.ps1 -SiteTitle 'Sensitivity Label Training' `
        -SiteAlias 'label-training' -TenantHostName 'contoso.sharepoint.com' `
        -WhatIf -Verbose

    Shows every site, page and upload operation that would run, without changing anything.

.NOTES
    Requires PowerShell 7+ and Microsoft.Graph.Authentication only (no PnP.PowerShell,
    no SharePoint CSOM, no other Microsoft.Graph.* modules). Every Graph call is a raw
    Invoke-MgGraphRequest and is logged at VERBOSE level.
    Module handling on every run: if Microsoft.Graph.Authentication is not installed,
    the script installs the latest version from the PowerShell Gallery for the current
    user (no switch needed). If more than one version is installed, the script warns
    and lists them, because side-by-side Microsoft.Graph.* versions cause assembly
    conflicts; keep only the newest (Uninstall-Module -Name
    Microsoft.Graph.Authentication -RequiredVersion <old>). If the newest installed
    version is behind the gallery, the script warns and suggests Update-Module. The
    highest installed version is always the one imported.
    Delegated scopes requested: Sites.Create.All (site creation),
    Sites.ReadWrite.All (pages), Files.ReadWrite.All (image upload). Sites.Create.All
    may require an administrator to consent for the Microsoft Graph PowerShell app once.
    Exit code 0 on success, 1 on any unrecoverable error.
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$SiteTitle,

    [Parameter(Mandatory)]
    [ValidatePattern('^[A-Za-z0-9][A-Za-z0-9._-]*$')]
    [string]$SiteAlias,

    [Parameter(Mandatory)]
    [ValidatePattern('^[A-Za-z0-9.-]+\.sharepoint\.(com|us|de|cn)$')]
    [string]$TenantHostName,

    [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
    [string]$OwnerUpn,

    [string]$SiteDescription = 'Sensitivity label deployment: why we label, who labels, what the labels mean, and where to apply them.',

    [string]$Locale = 'en-US',

    [string]$OverviewContentPath = 'dependencies/overview-outline.md',

    [string]$TroubleshootingContentPath = 'dependencies/troubleshooting-outline.md',

    [AllowEmptyString()]
    [string]$HomeBackgroundPath = 'backgrounds/home-background.jpg',

    [AllowEmptyString()]
    [string]$LabelsBackgroundPath = 'backgrounds/labels-background.jpg',

    [AllowEmptyString()]
    [string]$TroubleshootingBackgroundPath = 'backgrounds/troubleshooting-background.jpg',

    [ValidateNotNullOrEmpty()]
    [string]$LabelDefinitionPath = 'dependencies/labels-definition.json',

    [ValidatePattern('\.aspx$')]
    [string]$OverviewPageName = 'Home.aspx',

    [ValidatePattern('\.aspx$')]
    [string]$LabelsPageName = 'Sensitivitylabels.aspx',

    [ValidatePattern('\.aspx$')]
    [string]$TroubleshootingPageName = 'Troubleshooting.aspx',

    [string]$LogDirectory,

    [ValidateRange(30, 3600)]
    [int]$ProvisioningTimeoutSeconds = 600,

    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

# Web part type ids documented as supported by the Graph pages API.
$script:WebPartType = @{
    Image   = 'd1d91016-032f-456d-98a4-721247c305e8'
    Divider = '2161a1c6-db61-4731-b97c-3cdb303f7cbb'
    Spacer  = '8654b779-4886-46d4-8ffb-b5ed960ee986'
}

$script:LogFilePath = $null
$script:ScriptRoot  = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$script:Summary     = [ordered]@{
    SiteUrl        = $null
    SiteCreated    = $false
    Pages          = [System.Collections.Generic.List[object]]::new()
    Uploads        = 0
    Warnings       = [System.Collections.Generic.List[string]]::new()
    NavigationLinks = @()
}

#region Logging -----------------------------------------------------------------

function Initialize-Log {
    [CmdletBinding()]
    param([string]$Directory)

    if (-not $Directory) { $Directory = Join-Path $script:ScriptRoot 'Logs' }

    if (-not (Test-Path -LiteralPath $Directory)) {
        New-Item -Path $Directory -ItemType Directory -Force | Out-Null
    }

    $stamp = (Get-Date).ToString('yyyyMMdd-HHmmss')
    $script:LogFilePath = Join-Path $Directory ("New-LabelTrainingSite_{0}.txt" -f $stamp)

    $header = @(
        '================================================================'
        " New-LabelTrainingSite.ps1"
        " Started   : $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz'))"
        " Host      : $($env:COMPUTERNAME ?? [System.Net.Dns]::GetHostName())"
        " PowerShell: $($PSVersionTable.PSVersion)"
        '================================================================'
        ''
    ) -join [Environment]::NewLine

    Set-Content -LiteralPath $script:LogFilePath -Value $header -Encoding utf8
    return $Directory
}

function Write-Log {
    <#
        Single logging entry point. Every action is timestamped, typed, separated by a
        blank line in the .txt log, and echoed to the matching PowerShell stream.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [AllowEmptyString()]
        [string]$Message,

        [ValidateSet('INFO', 'VERBOSE', 'WARNING', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )

    $line = '[{0}] [{1,-7}] {2}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $Level, $Message

    if ($script:LogFilePath) {
        try {
            Add-Content -LiteralPath $script:LogFilePath -Value ($line + [Environment]::NewLine) -Encoding utf8
        }
        catch {
            Write-Warning "Could not write to log file '$script:LogFilePath': $($_.Exception.Message)"
        }
    }

    switch ($Level) {
        'VERBOSE' { Write-Verbose     $line }
        'WARNING' { Write-Warning     $line; $script:Summary.Warnings.Add($Message) }
        'ERROR'   { Write-Error       $line -ErrorAction Continue }
        default   { Write-Information $line }
    }
}

function Stop-WithError {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    if ($ErrorRecord) {
        Write-Log ("{0} :: {1}" -f $Message, $ErrorRecord.Exception.Message) -Level ERROR
        # Invoke-MgGraphRequest puts Graph's JSON error body (code + message) here; the
        # exception message alone is just the HTTP status, which says nothing about why.
        if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
            Write-Log ("Graph response: {0}" -f $ErrorRecord.ErrorDetails.Message) -Level ERROR
        }
        Write-Log ("Stack: {0}" -f $ErrorRecord.ScriptStackTrace) -Level VERBOSE
    }
    else {
        Write-Log $Message -Level ERROR
    }

    Write-Log "Run failed. Log file: $script:LogFilePath" -Level ERROR
    exit 1
}

function Test-GraphError {
    <#
        True when a failed Graph call matches the given HTTP status and/or Graph error
        code. The status lives on the HttpResponseMessage when one exists; the code is in
        the JSON body Invoke-MgGraphRequest surfaces through ErrorDetails.Message.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.Management.Automation.ErrorRecord]$ErrorRecord,
        [int[]]$StatusCode,
        [string[]]$Code
    )

    $exception = $ErrorRecord.Exception
    $response  = if ($exception.PSObject.Properties['Response']) { $exception.Response } else { $null }
    $status    = if ($response -and $response.PSObject.Properties['StatusCode']) { [int]$response.StatusCode } else { 0 }
    $text      = @($exception.Message, $ErrorRecord.ErrorDetails.Message) -join ' '

    if ($StatusCode -and $status -in $StatusCode) { return $true }
    foreach ($c in $Code) {
        if ($text -match [regex]::Escape($c)) { return $true }
    }
    # Fall back to the status number in the message text when no response object exists.
    if ($StatusCode -and -not $status) {
        foreach ($sc in $StatusCode) { if ($text -match "\b$sc\b") { return $true } }
    }
    return $false
}

#endregion

#region Prerequisites -----------------------------------------------------------

function Assert-RequiredModule {
    <#
        Only Microsoft.Graph.Authentication is needed: it provides Connect-MgGraph and
        Invoke-MgGraphRequest, and every Graph call in this script goes through the
        latter rather than a generated cmdlet.

        Runs on every invocation with no switch:
          * not installed          -> install the latest gallery version for the current user
          * several versions found -> warn and list them; side-by-side Microsoft.Graph.*
                                      versions load conflicting assemblies
          * newest installed version is behind the gallery -> warn, suggest Update-Module
        The highest installed version is always the one imported.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name
    )

    $installed = @(Get-Module -Name $Name -ListAvailable | Sort-Object -Property Version -Descending)

    if ($installed.Count -eq 0) {
        Write-Log "Module '$Name' is not installed; installing the latest version for the current user." -Level WARNING
        try {
            Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber
            $installed = @(Get-Module -Name $Name -ListAvailable | Sort-Object -Property Version -Descending)
            if ($installed.Count -eq 0) {
                throw "Install-Module completed but '$Name' is still not discoverable."
            }
            Write-Log "Installed module '$Name' $($installed[0].Version)." -Level SUCCESS
        }
        catch {
            Stop-WithError "Failed to install module '$Name'. Install it manually with: Install-Module $Name -Scope CurrentUser" $_
        }
    }

    $newest = $installed[0]

    if ($installed.Count -gt 1) {
        $versions = ($installed | ForEach-Object { "$($_.Version) ($($_.ModuleBase))" }) -join '; '
        Write-Log ("Multiple versions of '$Name' are installed: $versions. Side-by-side Microsoft.Graph.* versions " +
                   "cause assembly conflicts; keep only $($newest.Version) and remove the rest with: " +
                   "Uninstall-Module -Name $Name -RequiredVersion <old version>") -Level WARNING
    }

    # Best-effort check against the gallery; skipped silently when offline or the repository is unreachable.
    try {
        $gallery = Find-Module -Name $Name -Repository PSGallery -ErrorAction Stop
        if ($gallery -and [version]$gallery.Version -gt $newest.Version) {
            Write-Log ("Installed '$Name' $($newest.Version) is behind the PowerShell Gallery ($($gallery.Version)). " +
                       "Update with: Update-Module -Name $Name -Scope CurrentUser") -Level WARNING
        }
        else {
            Write-Log "Module '$Name' $($newest.Version) is the latest gallery version." -Level VERBOSE
        }
    }
    catch {
        Write-Log "Could not query the PowerShell Gallery for '$Name': $($_.Exception.Message)" -Level VERBOSE
    }

    try {
        Import-Module -Name $Name -RequiredVersion $newest.Version -ErrorAction Stop
        Write-Log "Imported module '$Name' $($newest.Version)." -Level VERBOSE
    }
    catch {
        Stop-WithError "Failed to import module '$Name' $($newest.Version)." $_
    }
}

function Invoke-Graph {
    <#
        Single wrapper around Invoke-MgGraphRequest. Every Graph call in the script goes
        through here so each request is logged once and responses come back as PSObjects
        carrying Graph's own camelCase property names.

        -Uri is relative to https://graph.microsoft.com/ (e.g. 'v1.0/sites/{id}') unless
        it is already absolute. -All follows @odata.nextLink and returns the combined
        'value' array.
    #>
    [CmdletBinding()]
    param(
        [ValidateSet('GET', 'POST', 'PATCH', 'PUT', 'DELETE')]
        [string]$Method = 'GET',

        [Parameter(Mandatory)][string]$Uri,

        [object]$Body,

        [string]$InputFilePath,

        [switch]$All
    )

    $fullUri = if ($Uri -match '^https?://') { $Uri } else { 'https://graph.microsoft.com/' + $Uri.TrimStart('/') }
    Write-Log "$Method $fullUri" -Level VERBOSE

    $request = @{
        Method      = $Method
        Uri         = $fullUri
        OutputType  = 'PSObject'
        ErrorAction = 'Stop'
    }

    if ($InputFilePath) {
        $request.InputFilePath = $InputFilePath
        $request.ContentType   = 'application/octet-stream'
    }
    elseif ($null -ne $Body) {
        $request.Body        = $Body | ConvertTo-Json -Depth 20 -Compress
        $request.ContentType = 'application/json'
    }

    $response = Invoke-MgGraphRequest @request

    if (-not $All) { return $response }

    $items = [System.Collections.Generic.List[object]]::new()
    while ($true) {
        if ($response.PSObject.Properties['value']) { $items.AddRange([object[]]@($response.value)) }

        $next = $response.PSObject.Properties['@odata.nextLink']
        if (-not $next -or -not $next.Value) { break }

        Write-Log "GET $($next.Value)" -Level VERBOSE
        $response = Invoke-MgGraphRequest -Method GET -Uri $next.Value -OutputType PSObject -ErrorAction Stop
    }
    return $items.ToArray()
}

function Connect-GraphDelegated {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string[]]$Scopes)

    Write-Log "Connecting to Microsoft Graph with delegated scopes: $($Scopes -join ', ')."

    try {
        Connect-MgGraph -Scopes $Scopes -NoWelcome -ErrorAction Stop
    }
    catch {
        Stop-WithError 'Interactive sign-in to Microsoft Graph failed.' $_
    }

    $context = Get-MgContext
    if (-not $context) { Stop-WithError 'Connected, but no Graph context was returned.' }

    if ($context.AuthType -ne 'Delegated') {
        Stop-WithError "This script requires delegated permissions; the current context is '$($context.AuthType)'. Run Disconnect-MgGraph and sign in interactively."
    }

    Write-Log "Signed in as $($context.Account) (tenant $($context.TenantId))." -Level SUCCESS

    $granted = @($context.Scopes)
    foreach ($scope in $Scopes) {
        if ($granted -notcontains $scope) {
            Write-Log "Scope '$scope' was requested but is not present on the token. Operations that need it will fail; an administrator may have to consent to it for the Microsoft Graph PowerShell application." -Level WARNING
        }
    }

    return $context
}

function Resolve-SiteOwner {
    <#
        The owner of a new site has to be the signed-in user. Delegated Sites.* scopes only
        reach sites the signed-in user can access, so a site created for somebody else is
        invisible to the GET /sites/{host}:/sites/{alias} lookup the provisioning wait and
        every later page and upload call depend on: the wait would time out and nothing
        would be configured. Defaulting to the signed-in account and refusing anything
        else fails before the site exists rather than after.
    #>
    [CmdletBinding()]
    param(
        [AllowEmptyString()][string]$RequestedOwner,
        [Parameter(Mandatory)][string]$SignedInAccount
    )

    if ([string]::IsNullOrWhiteSpace($RequestedOwner)) {
        Write-Log "No -OwnerUpn given; the signed-in account $SignedInAccount will own the site."
        return $SignedInAccount
    }

    if ($RequestedOwner -ieq $SignedInAccount) {
        return $SignedInAccount
    }

    Stop-WithError ("-OwnerUpn '$RequestedOwner' does not match the signed-in account '$SignedInAccount'. " +
        'With delegated permissions Graph can only find sites the signed-in user can access, so a site owned by ' +
        'someone else cannot be polled or configured by this script. Omit -OwnerUpn (the signed-in account becomes ' +
        "the owner) or sign in as $RequestedOwner, then grant additional owners in Site permissions afterwards.")
}

#endregion

#region Markdown -> web part content --------------------------------------------

function ConvertTo-HtmlText {
    <#
        Escapes one run of inline text and applies inline markdown: links, bold, italic
        and code. A literal <br> written in the outline survives as a line break; every
        other tag is shown as typed.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)

    $escaped = $Text -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;'

    $escaped = [regex]::Replace($escaped, '&lt;br\s*/?&gt;', '<br>', 'IgnoreCase')
    $escaped = [regex]::Replace($escaped, '\[([^\]]+)\]\(([^)\s]+)\)', '<a href="$2">$1</a>')
    $escaped = [regex]::Replace($escaped, '\*\*([^*]+)\*\*', '<strong>$1</strong>')
    $escaped = [regex]::Replace($escaped, '(?<!\*)\*([^*]+)\*(?!\*)', '<em>$1</em>')
    $escaped = [regex]::Replace($escaped, '`([^`]+)`', '<code>$1</code>')

    return $escaped
}

function Join-MarkdownLine {
    <#
        Joins the lines of one paragraph or list item into inline HTML. Lines join with a
        space, as markdown does; a line ending in two spaces or a backslash forces a
        line break.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyCollection()][string[]]$Line)

    $parts = foreach ($text in $Line) {
        if ($text -match '( {2,}|\\)$') { (ConvertTo-HtmlText $text.TrimEnd(' ', '\').Trim()) + '<br>' }
        else                            { ConvertTo-HtmlText $text.Trim() }
    }

    return (($parts -join ' ') -replace '<br> ', '<br>' -replace '<br>$', '')
}

function ConvertFrom-MarkdownBody {
    <#
        Converts the body of one markdown section into blocks: Text blocks holding the
        HTML a text web part accepts ('###'/'####' subheadings, bullet and numbered
        lists, paragraphs) and a Divider block wherever a '---' rule appears, so the
        caller can place a real Divider web part between the text web parts.

        Lines are read the way markdown intends: consecutive lines form one paragraph, a
        line directly under a list item continues that item, and a blank line between
        items of the same kind keeps the list going instead of starting a new one.

        One addition for FAQ outlines: a line starting with 'Q:' or 'A:' always opens a
        new paragraph and its prefix is shown in bold, so a question and its answer on
        consecutive lines read as a pair rather than merging into one paragraph.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyCollection()][AllowEmptyString()][string[]]$Line)

    $state = @{
        Block = [System.Collections.Generic.List[object]]::new()
        Html  = [System.Text.StringBuilder]::new()                  # text block being built
        Para  = [System.Collections.Generic.List[string]]::new()    # paragraph lines awaiting output
        Item  = [System.Collections.Generic.List[string]]::new()    # current list item's lines
        List  = $null                                               # 'ul' / 'ol' while a list is open
        Blank = $false                                              # blank line seen since the last item
    }

    $flushParagraph = {
        if ($state.Para.Count) {
            [void]$state.Html.Append('<p>' + (Join-MarkdownLine $state.Para) + '</p>')
            $state.Para.Clear()
        }
    }
    $flushItem = {
        if ($state.Item.Count) {
            [void]$state.Html.Append('<li>' + (Join-MarkdownLine $state.Item) + '</li>')
            $state.Item.Clear()
        }
    }
    $closeList = {
        & $flushItem
        if ($state.List) { [void]$state.Html.Append("</$($state.List)>"); $state.List = $null }
    }
    $endTextBlock = {
        if ($state.Html.Length) {
            $state.Block.Add(@{ Type = 'Text'; Html = $state.Html.ToString() })
            [void]$state.Html.Clear()
        }
    }

    foreach ($raw in $Line) {
        $text = $raw.Trim()

        if ($text.Length -eq 0) {
            & $flushParagraph
            $state.Blank = $true
            continue
        }

        # Horizontal rule ('---', '***', '- - -'). Checked before lists, since '* * *'
        # would otherwise read as a bullet.
        if ($text -match '^([-*_])( ?\1){2,}$') {
            & $flushParagraph
            & $closeList
            & $endTextBlock
            $state.Block.Add(@{ Type = 'Divider' })
            continue
        }

        if ($text -match '^(#{3,6})\s+(.+)$') {
            & $flushParagraph
            & $closeList
            $level = [Math]::Min($Matches[1].Length, 4)    # text web parts stop at h4
            [void]$state.Html.Append("<h$level>" + (ConvertTo-HtmlText $Matches[2]) + "</h$level>")
            continue
        }

        if ($text -match '^(?:[-*+]|(\d{1,9})[.)])\s+(.+)$') {
            $number  = $Matches[1]
            $content = $Matches[2]
            $kind    = if ($number) { 'ol' } else { 'ul' }

            & $flushParagraph
            if ($state.List -and $state.List -ne $kind) { & $closeList }
            if ($state.List) {
                & $flushItem
            }
            else {
                $start = if ($kind -eq 'ol' -and [int]$number -ne 1) { " start=`"$([int]$number)`"" } else { '' }
                [void]$state.Html.Append("<$kind$start>")
                $state.List = $kind
            }
            $state.Item.Add($content)
            $state.Blank = $false
            continue
        }

        # Text directly beneath a list item, or indented under it, continues that item.
        if ($state.List -and (-not $state.Blank -or $raw -match '^\s')) {
            $state.Item.Add($raw)
            continue
        }

        & $closeList

        # 'Q:' / 'A:' start their own paragraph with the prefix in bold.
        if ($text -match '^([QA]):\s+(.+)$') {
            & $flushParagraph
            $state.Para.Add(('**{0}:** {1}' -f $Matches[1], $Matches[2]))
            continue
        }

        $state.Para.Add($raw)
    }

    & $flushParagraph
    & $closeList
    & $endTextBlock

    return ,$state.Block.ToArray()
}

function Get-MarkdownSection {
    <#
        Returns one object per '##' heading: Heading plus the Text/Divider blocks of its
        body (see ConvertFrom-MarkdownBody). Content
        before the first '##' (the outline preamble) is ignored, so WHY / WHO / WHAT /
        WHERE each become their own page section. PageLabel and ParameterName only
        shape the messages, so the Overview and Troubleshooting outlines share this.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$PageLabel = 'Overview',
        [string]$ParameterName = 'OverviewContentPath'
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        Stop-WithError "$PageLabel content file '$Path' was not found. Pass -$ParameterName to point at the outline."
    }

    $lines    = Get-Content -LiteralPath $Path
    $sections = [System.Collections.Generic.List[object]]::new()
    $current  = $null
    $buffer   = [System.Collections.Generic.List[string]]::new()

    foreach ($line in $lines) {
        if ($line -match '^##\s+(.+?)\s*$') {
            if ($current) {
                $sections.Add([pscustomobject]@{ Heading = $current; Block = ConvertFrom-MarkdownBody -Line $buffer })
            }
            $current = $Matches[1]
            $buffer  = [System.Collections.Generic.List[string]]::new()
            continue
        }

        if ($line -match '^#\s+' ) {
            # A new top-level heading closes the section that was open.
            if ($current) {
                $sections.Add([pscustomobject]@{ Heading = $current; Block = ConvertFrom-MarkdownBody -Line $buffer })
                $current = $null
                $buffer  = [System.Collections.Generic.List[string]]::new()
            }
            continue
        }

        if ($current) { $buffer.Add($line) }
    }

    if ($current) {
        $sections.Add([pscustomobject]@{ Heading = $current; Block = ConvertFrom-MarkdownBody -Line $buffer })
    }

    if ($sections.Count -eq 0) {
        Write-Log "No '##' sections were found in '$Path'; the $PageLabel page will only carry the introduction." -Level WARNING
    }
    else {
        Write-Log "Parsed $($sections.Count) section(s) from '$Path': $(($sections.Heading) -join ', ')."
    }

    return $sections
}

#endregion

#region Web part and page builders ----------------------------------------------

function New-TextWebPart {
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$InnerHtml)

    return @{
        '@odata.type' = '#microsoft.graph.textWebPart'
        id            = [guid]::NewGuid().ToString()
        innerHtml     = $InnerHtml
    }
}

function New-DividerWebPart {
    [CmdletBinding()]
    param()

    return @{
        '@odata.type' = '#microsoft.graph.standardWebPart'
        id          = [guid]::NewGuid().ToString()
        webPartType = $script:WebPartType.Divider
        data        = @{
            dataVersion = '2.0'
            title       = 'Divider'
            description = 'Display a horizontal line'
            properties  = @{}
        }
    }
}

function New-ImageWebPart {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ServerRelativeUrl,
        [Parameter(Mandatory)][string]$SiteGuid,
        [Parameter(Mandatory)][string]$WebGuid,
        [Parameter(Mandatory)][string]$ListGuid,
        [Parameter(Mandatory)][string]$UniqueGuid,
        [string]$AltText = '',
        [string]$Caption = '',
        [int]$PixelWidth = 0,
        [int]$PixelHeight = 0
    )

    $properties = @{
        imageSourceType = 2
        altText         = $AltText
        overlayText     = ''
        siteid          = $SiteGuid
        webid           = $WebGuid
        listid          = $ListGuid
        uniqueid        = $UniqueGuid
        fixAspectRatio  = $false
        captionText     = $Caption
        alignment       = 'Center'
    }

    $customMetadata = @{
        siteid   = $SiteGuid
        webid    = $WebGuid
        listid   = $ListGuid
        uniqueid = $UniqueGuid
    }

    if ($PixelWidth -gt 0 -and $PixelHeight -gt 0) {
        $properties.imgWidth      = $PixelWidth
        $properties.imgHeight     = $PixelHeight
        $customMetadata.width     = "$PixelWidth"
        $customMetadata.height    = "$PixelHeight"
    }

    return @{
        '@odata.type' = '#microsoft.graph.standardWebPart'
        id          = [guid]::NewGuid().ToString()
        webPartType = $script:WebPartType.Image
        data        = @{
            dataVersion            = '1.9'
            title                  = 'Image'
            description            = 'Show an image on your page'
            properties             = $properties
            serverProcessedContent = @{
                imageSources   = @(@{ key = 'imageSource'; value = $ServerRelativeUrl })
                customMetadata = @(@{ key = 'imageSource'; value = $customMetadata })
            }
        }
    }
}

function New-PageSection {
    <#
        One horizontal section holding the supplied web parts in a single column.
        'fullWidth' stretches the column edge to edge, which is how a banner image is
        laid out; Graph rejects a fullWidth column unless its width is 0 ("mismatch of
        column width, should be 0"), while every other layout uses the 12-grid span.
        Section and column ids must be unique strings within the page.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][int]$Index,
        [Parameter(Mandatory)][object[]]$WebPart,
        [ValidateSet('none', 'neutral', 'soft', 'strong')][string]$Emphasis = 'none',
        [ValidateSet('oneColumn', 'fullWidth')][string]$Layout = 'oneColumn'
    )

    return @{
        id       = "$Index"
        layout   = $Layout
        emphasis = $Emphasis
        columns  = @(
            @{
                id       = '1'
                width    = if ($Layout -eq 'fullWidth') { 0 } else { 12 }
                webparts = @($WebPart)
            }
        )
    }
}

function New-BackgroundSection {
    <#
        Full-width image web part holding an uploaded background image, used as the
        opening section of the Overview, Labels and Troubleshooting pages. Returns $null
        when the image was not uploaded so callers can simply skip it.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][int]$Index,
        [AllowNull()][object]$Background,
        [Parameter(Mandatory)][string]$AltText
    )

    if (-not $Background) { return $null }

    $webPart = New-ImageWebPart -ServerRelativeUrl $Background.ServerRelativeUrl `
                                -SiteGuid    $Background.SiteGuid `
                                -WebGuid     $Background.WebGuid `
                                -ListGuid    $Background.ListGuid `
                                -UniqueGuid  $Background.UniqueGuid `
                                -AltText     $AltText `
                                -PixelWidth  $Background.Width `
                                -PixelHeight $Background.Height

    return New-PageSection -Index $Index -Layout 'fullWidth' -WebPart @($webPart)
}

function New-SitePageBody {
    <#
        TitleLayout is the title area style. Every page uses 'plain': the Overview,
        Labels and Troubleshooting pages carry their own background image in a
        full-width image web part, and the label pages open straight into their content
        with no colour block.

        PageLayout defaults to 'article', which always renders a title area above the
        canvas. The Overview, Labels and Troubleshooting pages pass 'home' instead, which
        is the one pageLayoutType that suppresses the title area entirely, since their own
        background image banner already serves as the page header and a second title
        strip above it is redundant. pageLayout is create-only: an in-place update keeps
        whatever layout the page already has, so it only takes effect when the page is
        created or recreated with -Force.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Section,
        [string]$TextAboveTitle = '',
        [ValidateSet('plain', 'colorBlock', 'imageAndTitle', 'overlap')][string]$TitleLayout = 'plain',
        [ValidateSet('article', 'home')][string]$PageLayout = 'article'
    )

    return @{
        '@odata.type'        = '#microsoft.graph.sitePage'
        name                 = $Name
        title                = $Title
        pageLayout           = $PageLayout
        showComments         = $false
        showRecommendedPages = $false
        titleArea            = @{
            enableGradientEffect    = ($TitleLayout -ne 'plain')
            layout                  = $TitleLayout
            showAuthor              = $false
            showPublishedDate       = $true
            showTextBlockAboveTitle = [bool]$TextAboveTitle
            textAboveTitle          = $TextAboveTitle
            textAlignment           = 'left'
            title                   = $Title
        }
        canvasLayout         = @{
            horizontalSections = @($Section)
        }
    }
}

#endregion

#region Label input -------------------------------------------------------------

function ConvertTo-NormalizedScreenshot {
    <#
        A screenshot entry is either a path string or an object with Path, Caption and
        Description. Both come back as the same three-property object so the upload
        loop and the page builder never have to tell them apart. Description may be one
        paragraph or an array of paragraphs; it is always returned as a string array.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Entry)

    if ($Entry -is [string]) {
        return [pscustomobject]@{ Path = $Entry; Caption = ''; Description = @() }
    }

    $names = $Entry.PSObject.Properties.Name
    $get = {
        param($Property, $Default)
        if ($names -contains $Property -and $null -ne $Entry.$Property) { $Entry.$Property } else { $Default }
    }

    return [pscustomobject]@{
        Path        = [string](& $get 'Path' '')
        Caption     = [string](& $get 'Caption' '')
        Description = @(& $get 'Description' @() | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { [string]$_ })
    }
}

function ConvertTo-NormalizedLabel {
    <#
        Guarantees every label object carries the same eight properties, because optional
        properties may simply be absent in the JSON file. Strict mode makes missing
        properties fatal otherwise.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Label, [Parameter(Mandatory)][int]$Index)

    $names = $Label.PSObject.Properties.Name
    $get = {
        param($Property, $Default)
        if ($names -contains $Property -and $null -ne $Label.$Property) { $Label.$Property } else { $Default }
    }

    return [pscustomobject]@{
        Index         = $Index
        Name          = [string](& $get 'Name' '')
        Group         = [string](& $get 'Group' '')
        VisibleTo     = [string](& $get 'VisibleTo' '')
        Description   = [string](& $get 'Description' '')
        CorrectUses   = @(& $get 'CorrectUses' @())
        IncorrectUses = @(& $get 'IncorrectUses' @())
        QuickNotes    = @(& $get 'QuickNotes' @())
        Screenshots   = @(@(& $get 'Screenshots' @()) | Where-Object { $null -ne $_ } | ForEach-Object { ConvertTo-NormalizedScreenshot -Entry $_ })
    }
}

function Get-LabelDefinition {
    <#
        Reads the labels from the JSON file supplied through -LabelDefinitionPath and
        returns them normalized. The file is required: the script never asks for label
        details interactively.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        Stop-WithError "Label definition file '$Path' was not found."
    }

    try {
        $raw = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
    }
    catch {
        Stop-WithError "Label definition file '$Path' is not valid JSON." $_
    }

    $labels = @($raw)
    if ($labels.Count -eq 0) {
        Stop-WithError "Label definition file '$Path' does not contain any labels."
    }

    foreach ($label in $labels) {
        foreach ($required in 'Name', 'VisibleTo', 'Description') {
            if (-not ($label.PSObject.Properties.Name -contains $required) -or [string]::IsNullOrWhiteSpace($label.$required)) {
                Stop-WithError "Label definition entry is missing required property '$required': $($label | ConvertTo-Json -Compress)"
            }
        }

        # A screenshot entry is a path string or an object; an object must carry a Path.
        if ($label.PSObject.Properties.Name -contains 'Screenshots' -and $label.Screenshots) {
            foreach ($shot in @($label.Screenshots)) {
                if ($null -eq $shot -or $shot -is [string]) { continue }
                if (-not ($shot.PSObject.Properties.Name -contains 'Path') -or [string]::IsNullOrWhiteSpace($shot.Path)) {
                    Stop-WithError "Screenshot entry for label '$($label.Name)' is missing required property 'Path': $($shot | ConvertTo-Json -Compress)"
                }
            }
        }
    }

    Write-Log "Loaded $($labels.Count) label definition(s) from '$Path'."

    $normalized = @()
    for ($i = 0; $i -lt $labels.Count; $i++) {
        $normalized += ConvertTo-NormalizedLabel -Label $labels[$i] -Index $i
    }
    return $normalized
}

#endregion

#region Site creation -----------------------------------------------------------

function Get-ExistingSite {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$HostName,
        [Parameter(Mandatory)][string]$Alias
    )

    try {
        $site = Invoke-Graph -Uri "v1.0/sites/${HostName}:/sites/${Alias}"
        if ($site) {
            Write-Log "Found an existing site at $($site.webUrl) (id $($site.id))."
            return $site
        }
    }
    catch {
        Write-Log "Lookup of https://$HostName/sites/$Alias did not return a site: $($_.Exception.Message)" -Level VERBOSE
    }

    Write-Log "No existing site at https://$HostName/sites/$Alias."
    return $null
}

function New-CommunicationSite {
    <#
        Creates a blank communication site. POST /sites exists only in the beta endpoint:
        v1.0 is read-only for site resources.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$HostName,
        [Parameter(Mandatory)][string]$Alias,
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Owner,
        [string]$Description,
        [string]$SiteLocale
    )

    $webUrl = "https://$HostName/sites/$Alias"

    $body = @{
        name                   = $Title
        webUrl                 = $webUrl
        locale                 = $SiteLocale
        description            = $Description
        shareByEmailEnabled    = $false
        template               = 'sitepagepublishing'   # blank communication site
        ownerIdentityToResolve = @{ email = $Owner }
    }

    if (-not $PSCmdlet.ShouldProcess($webUrl, 'Create blank communication site')) {
        Write-Log "WhatIf: would create communication site '$Title' at $webUrl owned by $Owner."
        return $null
    }

    Write-Log "Creating communication site '$Title' at $webUrl (POST /beta/sites)."
    try {
        Invoke-Graph -Method POST -Uri 'beta/sites' -Body $body | Out-Null
    }
    catch {
        Stop-WithError "Site creation failed for $webUrl." $_
    }

    Write-Log "Site creation accepted for $webUrl; provisioning is asynchronous." -Level SUCCESS
    return $webUrl
}

function Wait-SiteProvisioning {
    <#
        Site creation returns 202 Accepted. Poll until Graph can resolve the site, which
        is the condition the rest of the script depends on.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$HostName,
        [Parameter(Mandatory)][string]$Alias,
        [Parameter(Mandatory)][int]$TimeoutSeconds
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $delay    = 10
    $attempt  = 0

    Write-Log "Waiting up to $TimeoutSeconds seconds for https://$HostName/sites/$Alias to finish provisioning."

    while ((Get-Date) -lt $deadline) {
        $attempt++
        $site = Get-ExistingSite -HostName $HostName -Alias $Alias

        if ($site) {
            try {
                Invoke-Graph -Uri "v1.0/sites/$($site.id)/drive" | Out-Null
                Write-Log "Site is provisioned and reachable after $attempt check(s)." -Level SUCCESS
                return $site
            }
            catch {
                Write-Log "Site answers but its default document library is not ready yet (attempt $attempt)." -Level VERBOSE
            }
        }

        Start-Sleep -Seconds $delay
        if ($delay -lt 30) { $delay += 5 }
    }

    Stop-WithError "Timed out after $TimeoutSeconds seconds waiting for https://$HostName/sites/$Alias to provision. The site may still appear shortly; re-run the script to continue configuration."
}

function Wait-SitePage {
    <#
        The default document library answers well before the site template has stamped
        the Site Pages library and its stock Home.aspx. Poll the pages endpoint until the
        named page is listed so Save-SitePage takes the update path rather than racing
        provisioning to create a page that is about to exist anyway.

        Warns and returns $null on timeout instead of stopping: Save-SitePage tolerates a
        409 on create, so a late page is recoverable.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SiteId,
        [Parameter(Mandatory)][string]$PageName,
        [Parameter(Mandatory)][int]$TimeoutSeconds
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $delay    = 5
    $attempt  = 0

    Write-Log "Waiting up to $TimeoutSeconds seconds for the site template to provision '$PageName'."

    while ((Get-Date) -lt $deadline) {
        $attempt++
        try {
            $page = Get-SitePageByName -SiteId $SiteId -PageName $PageName
            if ($page) {
                Write-Log "'$PageName' is provisioned after $attempt check(s)." -Level SUCCESS
                return $page
            }
            Write-Log "'$PageName' is not listed yet (attempt $attempt)." -Level VERBOSE
        }
        catch {
            # A 404 here means the Site Pages list itself does not exist yet.
            Write-Log "Pages library is not ready yet (attempt $attempt): $($_.Exception.Message)" -Level VERBOSE
        }

        Start-Sleep -Seconds $delay
        if ($delay -lt 20) { $delay += 5 }
    }

    Write-Log "'$PageName' was not listed within $TimeoutSeconds seconds. Continuing; the page step will create it or fall back to an update if provisioning finishes first." -Level WARNING
    return $null
}

#endregion

#region Image upload ------------------------------------------------------------

function Get-ImageDimension {
    <#
        Reads intrinsic pixel size straight from PNG/JPEG/GIF headers so the image web
        part can carry imgWidth / imgHeight without a graphics library.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    try {
        # Widened to int: a PowerShell [byte] shifted left stays a byte and wraps to 0.
        $bytes = [int[]][System.IO.File]::ReadAllBytes($Path)
    }
    catch {
        return $null
    }

    if ($bytes.Length -lt 24) { return $null }

    # PNG
    if ($bytes[0] -eq 0x89 -and $bytes[1] -eq 0x50 -and $bytes[2] -eq 0x4E -and $bytes[3] -eq 0x47) {
        $width  = ($bytes[16] -shl 24) -bor ($bytes[17] -shl 16) -bor ($bytes[18] -shl 8) -bor $bytes[19]
        $height = ($bytes[20] -shl 24) -bor ($bytes[21] -shl 16) -bor ($bytes[22] -shl 8) -bor $bytes[23]
        return [pscustomobject]@{ Width = $width; Height = $height }
    }

    # GIF
    if ($bytes[0] -eq 0x47 -and $bytes[1] -eq 0x49 -and $bytes[2] -eq 0x46) {
        $width  = $bytes[6] -bor ($bytes[7] -shl 8)
        $height = $bytes[8] -bor ($bytes[9] -shl 8)
        return [pscustomobject]@{ Width = $width; Height = $height }
    }

    # JPEG: walk the marker segments to the first start-of-frame.
    if ($bytes[0] -eq 0xFF -and $bytes[1] -eq 0xD8) {
        $i = 2
        while ($i -lt $bytes.Length - 9) {
            if ($bytes[$i] -ne 0xFF) { $i++; continue }
            $marker = $bytes[$i + 1]
            $length = ($bytes[$i + 2] -shl 8) -bor $bytes[$i + 3]

            if ($marker -ge 0xC0 -and $marker -le 0xCF -and $marker -notin 0xC4, 0xC8, 0xCC) {
                $height = ($bytes[$i + 5] -shl 8) -bor $bytes[$i + 6]
                $width  = ($bytes[$i + 7] -shl 8) -bor $bytes[$i + 8]
                return [pscustomobject]@{ Width = $width; Height = $height }
            }
            $i += 2 + $length
        }
    }

    return $null
}

function Add-SiteImage {
    <#
        Uploads one local image (a label screenshot or a page background) into a folder
        of the site's default document library and returns everything the image web part
        needs, including the SharePoint ids Graph exposes on the driveItem's
        sharepointIds facet. Kind only labels the log messages.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$SiteId,
        [Parameter(Mandatory)][string]$DriveId,
        [Parameter(Mandatory)][string]$FolderItemId,
        [Parameter(Mandatory)][string]$Path,
        [string]$Kind = 'screenshot'
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Log "$Kind '$Path' does not exist; skipping it." -Level WARNING
        return $null
    }

    $fileName = Split-Path -Path $Path -Leaf

    if (-not $PSCmdlet.ShouldProcess($fileName, "Upload $Kind to the site document library")) {
        Write-Log "WhatIf: would upload '$Path' to the site document library."
        return $null
    }

    try {
        $encodedName = [uri]::EscapeDataString($fileName)
        $uploaded = Invoke-Graph -Method PUT `
            -Uri "v1.0/drives/$DriveId/items/${FolderItemId}:/${encodedName}:/content?@microsoft.graph.conflictBehavior=replace" `
            -InputFilePath $Path

        $item = Invoke-Graph -Uri "v1.0/drives/$DriveId/items/$($uploaded.id)?`$select=id,name,webUrl,sharepointIds"
    }
    catch {
        Write-Log "Upload of $Kind '$Path' failed: $($_.Exception.Message). The page will be built without it." -Level WARNING
        return $null
    }

    $spIds = if ($item.PSObject.Properties['sharepointIds']) { $item.sharepointIds } else { $null }
    if (-not $spIds -or -not $spIds.PSObject.Properties['listItemUniqueId'] -or -not $spIds.listItemUniqueId) {
        Write-Log "Uploaded $Kind '$fileName' but Graph did not return its SharePoint ids, which the image web part requires. The page will be built without it." -Level WARNING
        return $null
    }

    $script:Summary.Uploads++
    $serverRelativeUrl = ([uri]$item.webUrl).AbsolutePath
    $dimension = Get-ImageDimension -Path $Path

    Write-Log "Uploaded $Kind '$fileName' to $serverRelativeUrl." -Level VERBOSE

    return [pscustomobject]@{
        FileName          = $fileName
        WebUrl            = $item.webUrl
        ServerRelativeUrl = $serverRelativeUrl
        SiteGuid          = $spIds.siteId
        WebGuid           = $spIds.webId
        ListGuid          = $spIds.listId
        UniqueGuid        = $spIds.listItemUniqueId
        Width             = if ($dimension) { $dimension.Width }  else { 0 }
        Height            = if ($dimension) { $dimension.Height } else { 0 }
    }
}

function Get-ImageFolder {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$DriveId,
        [string]$FolderName = 'LabelScreenshots'
    )

    if (-not $PSCmdlet.ShouldProcess($FolderName, 'Create image folder in the document library')) {
        Write-Log "WhatIf: would ensure the '$FolderName' folder exists in the document library."
        return $null
    }

    try {
        $folder = Invoke-Graph -Method POST -Uri "v1.0/drives/$DriveId/root/children" -Body @{
            name                                = $FolderName
            folder                              = @{}
            '@microsoft.graph.conflictBehavior' = 'replace'
        }

        Write-Log "Image folder '$FolderName' is ready (item $($folder.id))." -Level VERBOSE
        return $folder.id
    }
    catch {
        Write-Log "Could not create the '$FolderName' folder: $($_.Exception.Message). Images destined for it will be skipped." -Level WARNING
        return $null
    }
}

#endregion

#region Page publishing ---------------------------------------------------------

function Get-SitePageByName {
    <#
        Lists the site's pages and returns the one whose file name matches, or $null.
        Throws when the listing itself fails so callers can decide whether that is fatal.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SiteId,
        [Parameter(Mandatory)][string]$PageName
    )

    return @(Invoke-Graph -Uri "v1.0/sites/$SiteId/pages/microsoft.graph.sitePage?`$select=id,name,webUrl" -All) |
           Where-Object { $_.name -and $_.name.Equals($PageName, [StringComparison]::OrdinalIgnoreCase) } |
           Select-Object -First 1
}

function Save-SitePage {
    <#
        Creates the page, or updates/recreates it when it already exists, then publishes
        it so readers see it rather than a draft.

        On a freshly created site the stock Home.aspx can appear between the listing and
        the create call, so a 409 on create re-lists and falls through to the update path.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$SiteId,
        [Parameter(Mandatory)][hashtable]$Body,
        [switch]$Recreate
    )

    $pageName = $Body.name

    if (-not $PSCmdlet.ShouldProcess($pageName, 'Create or update site page')) {
        Write-Log "WhatIf: would create or update page '$pageName' with $($Body.canvasLayout.horizontalSections.Count) section(s)."
        return $null
    }

    try {
        $existing = Get-SitePageByName -SiteId $SiteId -PageName $pageName
    }
    catch {
        Stop-WithError "Could not list the pages of site $SiteId." $_
    }

    $page = $null

    if ($existing -and $Recreate) {
        Write-Log "Page '$pageName' exists and -Force was supplied; deleting and recreating it."
        try {
            Invoke-Graph -Method DELETE -Uri "v1.0/sites/$SiteId/pages/$($existing.id)" | Out-Null
            $existing = $null
        }
        catch {
            Write-Log "Could not delete page '$pageName': $($_.Exception.Message). Falling back to an in-place update." -Level WARNING
        }
    }

    if (-not $existing) {
        Write-Log "Creating page '$pageName'."
        try {
            $page = Invoke-Graph -Method POST -Uri "v1.0/sites/$SiteId/pages" -Body $Body
        }
        catch {
            if (-not (Test-GraphError -ErrorRecord $_ -StatusCode 409 -Code 'nameAlreadyExists')) {
                Stop-WithError "Failed to create page '$pageName'." $_
            }

            # Lost the race with site provisioning: the page now exists, so update it instead.
            Write-Log "Page '$pageName' already exists (created by site provisioning after the listing); switching to an in-place update." -Level WARNING
            try {
                $existing = Get-SitePageByName -SiteId $SiteId -PageName $pageName
            }
            catch {
                Stop-WithError "Page '$pageName' reported as existing but the pages of site $SiteId could not be listed." $_
            }
            if (-not $existing) {
                Stop-WithError "Page '$pageName' reported as existing but is not listed yet. Re-run the script once provisioning settles."
            }
        }
    }

    if ($existing) {
        Write-Log "Updating existing page '$pageName' (id $($existing.id))."
        # The update endpoint rejects the immutable create-only properties (name, pageLayout)
        # with 400 "Property 'name' cannot be used in this request", so send only the
        # updatable subset: title, titleArea, canvasLayout, showComments, showRecommendedPages.
        $updateBody = @{}
        foreach ($key in $Body.Keys) {
            if ($key -notin @('name', 'pageLayout')) { $updateBody[$key] = $Body[$key] }
        }
        try {
            Invoke-Graph -Method PATCH -Uri "v1.0/sites/$SiteId/pages/$($existing.id)/microsoft.graph.sitePage" -Body $updateBody | Out-Null
            $page = $existing
        }
        catch {
            Stop-WithError "Failed to update page '$pageName'. Re-run with -Force to delete and recreate it." $_
        }
    }

    Publish-SitePage -SiteId $SiteId -PageId $page.id -PageName $pageName

    $webUrl = if ($page.PSObject.Properties['webUrl'] -and $page.webUrl) { $page.webUrl } else { $pageName }
    $script:Summary.Pages.Add([pscustomobject]@{ Name = $pageName; Title = $Body.title; WebUrl = $webUrl })

    Write-Log "Page '$pageName' saved." -Level SUCCESS
    return $page
}

function Publish-SitePage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SiteId,
        [Parameter(Mandatory)][string]$PageId,
        [Parameter(Mandatory)][string]$PageName
    )

    try {
        Invoke-Graph -Method POST -Uri "v1.0/sites/$SiteId/pages/$PageId/microsoft.graph.sitePage/publish" | Out-Null
        Write-Log "Published page '$PageName'." -Level VERBOSE
    }
    catch {
        Write-Log "Page '$PageName' was saved but NOT published, so it stays a draft: $($_.Exception.Message). Publish it from the page's 'Publish' button." -Level WARNING
    }
}

#endregion

#region Page content ------------------------------------------------------------

function Write-TopNavigationGuidance {
    <#
        Graph has no write surface for a communication site's top navigation, so the
        links are recorded on the run summary and warned about instead of applied. The
        warning is deliberate: it is the one step the operator has to finish by hand.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Link)

    $script:Summary.NavigationLinks = @($Link)

    if (@($Link).Count -eq 0) { return }

    Write-Log ('Microsoft Graph cannot edit a site''s top navigation. Add these {0} link(s) by hand in Site settings > Navigation, then nest the label pages under "Sensitivity labels".' -f @($Link).Count) -Level WARNING

    foreach ($item in $Link) {
        Write-Log ("Navigation link: {0} -> {1}" -f $item.Title, $item.Url) -Level VERBOSE
    }
}

function Get-LabelPageName {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$LabelName)

    $slug = ($LabelName -replace '[^A-Za-z0-9]+', '-').Trim('-')
    if (-not $slug) { $slug = 'label' }
    return "Label-$slug.aspx"
}

function New-OverviewPageBody {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$MarkdownSection,
        [Parameter(Mandatory)][object[]]$NavigationLink,
        [Parameter(Mandatory)][string]$PageName,
        [Parameter(Mandatory)][string]$Title,
        [AllowNull()][object]$Background
    )

    $sections = [System.Collections.Generic.List[object]]::new()
    $index    = 1

    # Background image in a full-width image web part, in place of a colour-block title.
    $banner = New-BackgroundSection -Index $index -Background $Background -AltText "$Title background"
    if ($banner) { $sections.Add($banner); $index++ }

    # The three navigation destinations, with the descriptions modern SharePoint
    # navigation cannot store, kept on the page where they survive.
    $navHtml = [System.Text.StringBuilder]::new()
    [void]$navHtml.Append('<h2>On this site</h2><ul>')
    foreach ($link in $NavigationLink) {
        $item = '<li><a href="{0}"><strong>{1}</strong></a> &mdash; {2}</li>' -f
                    $link.Url, (ConvertTo-HtmlText $link.Title), (ConvertTo-HtmlText $link.Description)
        [void]$navHtml.Append($item)
    }
    [void]$navHtml.Append('</ul>')

    $sections.Add((New-PageSection -Index ($index++) -Emphasis 'soft' -WebPart @(New-TextWebPart -InnerHtml $navHtml.ToString())))
    $sections.Add((New-PageSection -Index ($index++) -WebPart @(New-DividerWebPart)))

    foreach ($section in $MarkdownSection) {
        # The heading rides in the first text web part; each '---' in the body becomes a
        # Divider web part between text web parts, all inside the same page section.
        $webParts = [System.Collections.Generic.List[object]]::new()
        $html     = '<h2>{0}</h2>' -f (ConvertTo-HtmlText $section.Heading)
        foreach ($block in $section.Block) {
            if ($block.Type -eq 'Divider') {
                if ($html) { $webParts.Add((New-TextWebPart -InnerHtml $html)); $html = '' }
                $webParts.Add((New-DividerWebPart))
            }
            else {
                $html += $block.Html
            }
        }
        if ($html) { $webParts.Add((New-TextWebPart -InnerHtml $html)) }
        $sections.Add((New-PageSection -Index ($index++) -WebPart $webParts.ToArray()))
    }

    return New-SitePageBody -Name $PageName -Title $Title -TextAboveTitle 'Overview' -TitleLayout 'plain' -PageLayout 'home' -Section $sections.ToArray()
}

function New-LabelsPageBody {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Label,
        [Parameter(Mandatory)][string]$PageName,
        [Parameter(Mandatory)][string]$SitePagesUrl,
        [AllowNull()][object]$Background
    )

    $sections = [System.Collections.Generic.List[object]]::new()
    $index    = 1

    # Background image in a full-width image web part, in place of a colour-block title.
    $banner = New-BackgroundSection -Index $index -Background $Background -AltText 'Sensitivity labels background'
    if ($banner) { $sections.Add($banner); $index++ }

    $intro = '<p>Every sensitivity label in use is described on its own page: who can see it, what it means, ' +
             'correct and incorrect uses, quick notes, and screenshots of applying it. Start with the label your department uses most.</p>'
    $sections.Add((New-PageSection -Index ($index++) -WebPart @(New-TextWebPart -InnerHtml $intro)))

    $individual = @($Label | Where-Object { -not $_.Group })
    $grouped    = @($Label | Where-Object { $_.Group }) | Group-Object -Property Group

    # One list entry per label, linking to its page; shared by both lists below.
    $entryFor = {
        param($item)
        '<li><a href="{0}/{1}"><strong>{2}</strong></a> &mdash; {3}</li>' -f
            $SitePagesUrl, (Get-LabelPageName -LabelName $item.Name),
            (ConvertTo-HtmlText $item.Name), (ConvertTo-HtmlText $item.VisibleTo)
    }

    if ($individual.Count -gt 0) {
        $html = [System.Text.StringBuilder]::new()
        [void]$html.Append('<h2>Individual labels</h2><ul>')
        foreach ($item in $individual) { [void]$html.Append((& $entryFor $item)) }
        [void]$html.Append('</ul>')
        $sections.Add((New-PageSection -Index ($index++) -WebPart @(New-TextWebPart -InnerHtml $html.ToString())))
    }

    # Every group sits under one 'Label groups' heading, each as a sub-heading with its labels.
    if ($grouped.Count -gt 0) {
        $html = [System.Text.StringBuilder]::new()
        [void]$html.Append('<h2>Label groups</h2>')
        foreach ($group in $grouped) {
            [void]$html.Append('<h3>{0}</h3><ul>' -f (ConvertTo-HtmlText $group.Name))
            foreach ($item in $group.Group) { [void]$html.Append((& $entryFor $item)) }
            [void]$html.Append('</ul>')
        }
        $sections.Add((New-PageSection -Index ($index++) -WebPart @(New-TextWebPart -InnerHtml $html.ToString())))
    }

    if ($Label.Count -eq 0) {
        $sections.Add((New-PageSection -Index ($index++) -WebPart @(
            New-TextWebPart -InnerHtml '<p><em>No labels have been defined yet.</em></p>'
        )))
    }

    return New-SitePageBody -Name $PageName -Title 'Labels' -TextAboveTitle 'Sensitivity labels' -TitleLayout 'plain' -PageLayout 'home' -Section $sections.ToArray()
}

function New-TroubleshootingPageBody {
    <#
        Troubleshooting & FAQ page: the background image banner, then one page section
        per '##' heading of the troubleshooting outline, laid out exactly like the
        Overview sections. The title area is suppressed ('home' layout) because the
        background image carries the page title.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$MarkdownSection,
        [Parameter(Mandatory)][string]$PageName,
        [AllowNull()][object]$Background
    )

    $sections = [System.Collections.Generic.List[object]]::new()
    $index    = 1

    # Background image in a full-width image web part, in place of a colour-block title.
    $banner = New-BackgroundSection -Index $index -Background $Background -AltText 'Troubleshooting and FAQ background'
    if ($banner) { $sections.Add($banner); $index++ }

    foreach ($section in $MarkdownSection) {
        # The heading rides in the first text web part; each '---' in the body becomes a
        # Divider web part between text web parts, all inside the same page section.
        $webParts = [System.Collections.Generic.List[object]]::new()
        $html     = '<h2>{0}</h2>' -f (ConvertTo-HtmlText $section.Heading)
        foreach ($block in $section.Block) {
            if ($block.Type -eq 'Divider') {
                if ($html) { $webParts.Add((New-TextWebPart -InnerHtml $html)); $html = '' }
                $webParts.Add((New-DividerWebPart))
            }
            else {
                $html += $block.Html
            }
        }
        if ($html) { $webParts.Add((New-TextWebPart -InnerHtml $html)) }
        $sections.Add((New-PageSection -Index ($index++) -WebPart $webParts.ToArray()))
    }

    if ($MarkdownSection.Count -eq 0) {
        $sections.Add((New-PageSection -Index ($index++) -WebPart @(
            New-TextWebPart -InnerHtml '<p><em>No troubleshooting content has been written yet.</em></p>'
        )))
    }

    return New-SitePageBody -Name $PageName -Title 'Troubleshooting & FAQ' -TextAboveTitle 'Troubleshooting' -TitleLayout 'plain' -PageLayout 'home' -Section $sections.ToArray()
}

function New-LabelPageBody {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Label,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Screenshot,
        [Parameter(Mandatory)][string]$LabelsPageUrl
    )

    $sections = [System.Collections.Generic.List[object]]::new()
    $index    = 1

    # Name and Group are carried by the page title and the text above it (see the end of
    # this function); the sections below follow in the order of labels-definition.json.
    $visibleToHtml = '<h2>Visible to</h2><p>{0}</p>' -f (ConvertTo-HtmlText $Label.VisibleTo)
    $sections.Add((New-PageSection -Index ($index++) -Emphasis 'soft' -WebPart @(New-TextWebPart -InnerHtml $visibleToHtml)))

    $descriptionHtml = '<h2>📄 Description</h2><p>{0}</p>' -f (ConvertTo-HtmlText $Label.Description)
    $sections.Add((New-PageSection -Index ($index++) -WebPart @(New-TextWebPart -InnerHtml $descriptionHtml)))

    $listSections = @(
        @{ Property = 'CorrectUses';   Heading = '✅ Correct uses';   Empty = 'Correct uses for this label have not been provided yet.' }
        @{ Property = 'IncorrectUses'; Heading = '❌ Incorrect uses'; Empty = 'Incorrect uses for this label have not been provided yet.' }
        @{ Property = 'QuickNotes';    Heading = '📓Quick notes';    Empty = 'No quick notes for this label.' }
    )
    foreach ($list in $listSections) {
        $items = @()
        if ($Label.PSObject.Properties.Name -contains $list.Property -and $Label.($list.Property)) { $items = @($Label.($list.Property)) }

        $listHtml = [System.Text.StringBuilder]::new()
        [void]$listHtml.Append('<h2>{0}</h2>' -f $list.Heading)
        if ($items.Count -gt 0) {
            [void]$listHtml.Append('<ul>')
            foreach ($item in $items) {
                [void]$listHtml.Append('<li>' + (ConvertTo-HtmlText $item) + '</li>')
            }
            [void]$listHtml.Append('</ul>')
        }
        else {
            [void]$listHtml.Append('<p><em>{0}</em></p>' -f $list.Empty)
        }
        $sections.Add((New-PageSection -Index ($index++) -WebPart @(New-TextWebPart -InnerHtml $listHtml.ToString())))
    }

    $sections.Add((New-PageSection -Index ($index++) -WebPart @(New-DividerWebPart)))
    $sections.Add((New-PageSection -Index ($index++) -WebPart @(
        New-TextWebPart -InnerHtml '<h2>📸 Screenshots</h2><p>Screenshots for this label can be found below:</p>'
    )))

    $uploaded = @($Screenshot | Where-Object { $_ })
    if ($uploaded.Count -gt 0) {
        foreach ($image in $uploaded) {
            # The caption is the image web part's own captionText (plain text under the
            # image); a longer description becomes a text web part in the same section,
            # one <p> per paragraph, so it stays attached to its screenshot.
            $webParts = [System.Collections.Generic.List[object]]::new()
            $webParts.Add((New-ImageWebPart -ServerRelativeUrl $image.ServerRelativeUrl `
                                            -SiteGuid   $image.SiteGuid `
                                            -WebGuid    $image.WebGuid `
                                            -ListGuid   $image.ListGuid `
                                            -UniqueGuid $image.UniqueGuid `
                                            -AltText    ("Applying the {0} label" -f $Label.Name) `
                                            -Caption    $image.Caption `
                                            -PixelWidth $image.Width `
                                            -PixelHeight $image.Height))

            $paragraphs = @($image.Description | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            if ($paragraphs.Count -gt 0) {
                $html = ($paragraphs | ForEach-Object { '<p>' + (ConvertTo-HtmlText $_) + '</p>' }) -join ''
                $webParts.Add((New-TextWebPart -InnerHtml $html))
            }

            $sections.Add((New-PageSection -Index ($index++) -WebPart $webParts.ToArray()))
        }
    }
    else {
        $sections.Add((New-PageSection -Index ($index++) -WebPart @(
            New-TextWebPart -InnerHtml ('<p><em>No screenshots were supplied for this label.</em> Microsoft publishes ' +
                'step-by-step instructions with screenshots for applying sensitivity labels across the Microsoft 365 apps: ' +
                '<a href="https://support.microsoft.com/office/apply-sensitivity-labels-to-your-files-and-email-2f96e7cd-d5a4-403b-8bd7-4cc636bae0f9">' +
                'Apply sensitivity labels to your files and email</a>.</p>')
        )))
    }

    $sections.Add((New-PageSection -Index ($index++) -WebPart @(
        New-TextWebPart -InnerHtml ('<p><a href="{0}">Back to all labels</a></p>' -f $LabelsPageUrl)
    )))

    $groupText = if ($Label.Group) { $Label.Group } else { 'Sensitivity label' }

    # Plain title area: no colour block and no background image on the label pages.
    return New-SitePageBody -Name (Get-LabelPageName -LabelName $Label.Name) `
                            -Title $Label.Name `
                            -TextAboveTitle $groupText `
                            -TitleLayout 'plain' `
                            -Section $sections.ToArray()
}

#endregion

#region Main --------------------------------------------------------------------

$logFolder = Initialize-Log -Directory $LogDirectory

Write-Log "Target site: https://$TenantHostName/sites/$SiteAlias ('$SiteTitle')."
Write-Log "Log folder: $logFolder"

try {
    # 1. Modules ---------------------------------------------------------------
    Assert-RequiredModule -Name 'Microsoft.Graph.Authentication'

    # 2. Sign in ---------------------------------------------------------------
    $graphContext = Connect-GraphDelegated -Scopes @(
        'Sites.ReadWrite.All'   # pages, lists, document library
        'Sites.Create.All'      # site creation (beta POST /sites)
        'Files.ReadWrite.All'   # background and screenshot upload
    )

    $OwnerUpn = Resolve-SiteOwner -RequestedOwner $OwnerUpn -SignedInAccount $graphContext.Account

    # 3. Resolve the overview content and the labels before touching the tenant.
    $overviewPath = if ([System.IO.Path]::IsPathRooted($OverviewContentPath)) {
        $OverviewContentPath
    }
    else {
        Join-Path $script:ScriptRoot $OverviewContentPath
    }

    $markdownSections = Get-MarkdownSection -Path $overviewPath

    $troubleshootingPath = if ([System.IO.Path]::IsPathRooted($TroubleshootingContentPath)) {
        $TroubleshootingContentPath
    }
    else {
        Join-Path $script:ScriptRoot $TroubleshootingContentPath
    }

    $troubleshootingSections = Get-MarkdownSection -Path $troubleshootingPath -PageLabel 'Troubleshooting' -ParameterName 'TroubleshootingContentPath'
    $labelDefinitionFile = if ([System.IO.Path]::IsPathRooted($LabelDefinitionPath)) {
        $LabelDefinitionPath
    }
    else {
        Join-Path $script:ScriptRoot $LabelDefinitionPath
    }

    $labels = @(Get-LabelDefinition -Path $labelDefinitionFile)
    Write-Log "Working with $($labels.Count) label(s)."

    # 4. Site ------------------------------------------------------------------
    $site = Get-ExistingSite -HostName $TenantHostName -Alias $SiteAlias

    if ($site) {
        Write-Log 'Site already exists; skipping creation and reconciling pages only.'
    }
    else {
        New-CommunicationSite -HostName $TenantHostName -Alias $SiteAlias -Title $SiteTitle `
                              -Owner $OwnerUpn -Description $SiteDescription -SiteLocale $Locale | Out-Null

        if ($WhatIfPreference) {
            Write-Log 'WhatIf: no site was created, so page and navigation steps are reported only.'
        }
        else {
            $provisionStart = Get-Date
            $site = Wait-SiteProvisioning -HostName $TenantHostName -Alias $SiteAlias -TimeoutSeconds $ProvisioningTimeoutSeconds
            $script:Summary.SiteCreated = $true

            # Only the stock home page is shared with provisioning; every other page is
            # ours to create. Spend whatever is left of the provisioning budget (at least
            # a minute) waiting for it so the page step updates rather than races it.
            if ($OverviewPageName -eq 'Home.aspx') {
                $elapsed   = [int]((Get-Date) - $provisionStart).TotalSeconds
                $remaining = [Math]::Max(60, $ProvisioningTimeoutSeconds - $elapsed)
                Wait-SitePage -SiteId $site.id -PageName $OverviewPageName -TimeoutSeconds $remaining | Out-Null
            }
        }
    }

    $siteUrl = "https://$TenantHostName/sites/$SiteAlias"
    $script:Summary.SiteUrl = $siteUrl

    $navigationLinks = @(
        [pscustomobject]@{ Title = 'Overview';  Url = "$siteUrl/SitePages/$OverviewPageName"; Description = "Introduction and background on this site's purpose." }
        [pscustomobject]@{ Title = 'Sensitivity labels';    Url = "$siteUrl/SitePages/$LabelsPageName";   Description = 'Details on each sensitivity label, including descriptions, examples, and screenshots.' }
        [pscustomobject]@{ Title = 'Troubleshooting & FAQ'; Url = "$siteUrl/SitePages/$TroubleshootingPageName"; Description = 'Answers to common questions and fixes for when labels do not behave as expected.' }
    )

    # Recorded on every run, not just the WhatIf path: the summary prints these links and
    # they are the same manual step whether the site was created now or already existed.
    Write-TopNavigationGuidance -Link $navigationLinks

    if (-not $site) {
        Write-Log 'No site object is available (WhatIf run); reporting the intended pages without calling Graph.'
    }
    else {
        $siteId = $site.id

        # 5. Images: page backgrounds and label screenshots -------------------
        $backgroundPaths = [ordered]@{}
        if ($HomeBackgroundPath)            { $backgroundPaths.Home            = $HomeBackgroundPath }
        if ($LabelsBackgroundPath)          { $backgroundPaths.Labels          = $LabelsBackgroundPath }
        if ($TroubleshootingBackgroundPath) { $backgroundPaths.Troubleshooting = $TroubleshootingBackgroundPath }

        $backgrounds        = @{ Home = $null; Labels = $null; Troubleshooting = $null }
        $screenshotsByLabel = @{}
        $anyScreenshots = @($labels | Where-Object {
            $_.PSObject.Properties.Name -contains 'Screenshots' -and $_.Screenshots -and @($_.Screenshots).Count -gt 0
        }).Count -gt 0

        $drive = $null
        if ($backgroundPaths.Count -gt 0 -or $anyScreenshots) {
            try {
                $drive = Invoke-Graph -Uri "v1.0/sites/$siteId/drive"
                Write-Log "Default document library: $($drive.name) (drive $($drive.id))." -Level VERBOSE
            }
            catch {
                Write-Log "Could not open the site's default document library: $($_.Exception.Message). Background images and screenshots will be skipped." -Level WARNING
            }
        }

        if ($drive -and $backgroundPaths.Count -gt 0) {
            $backgroundFolderId = Get-ImageFolder -DriveId $drive.id -FolderName 'Backgrounds'
            if ($backgroundFolderId) {
                foreach ($page in @($backgroundPaths.Keys)) {
                    $path = $backgroundPaths[$page]
                    $resolved = if ([System.IO.Path]::IsPathRooted($path)) { $path } else { Join-Path $script:ScriptRoot $path }
                    $backgrounds[$page] = Add-SiteImage -SiteId $siteId -DriveId $drive.id -FolderItemId $backgroundFolderId `
                                                        -Path $resolved -Kind "$page page background"
                }
            }
        }

        if ($drive -and $anyScreenshots) {
            $folderId = Get-ImageFolder -DriveId $drive.id
            if ($folderId) {
                foreach ($label in $labels) {
                    $uploads = [System.Collections.Generic.List[object]]::new()
                    foreach ($shot in @($label.Screenshots)) {
                        $resolved = if ([System.IO.Path]::IsPathRooted($shot.Path)) { $shot.Path } else { Join-Path $script:ScriptRoot $shot.Path }
                        $result = Add-SiteImage -SiteId $siteId -DriveId $drive.id -FolderItemId $folderId -Path $resolved
                        if ($result) {
                            # The caption and description ride along with the upload result so
                            # the page builder gets one object per image.
                            $result | Add-Member -NotePropertyName Caption     -NotePropertyValue $shot.Caption
                            $result | Add-Member -NotePropertyName Description -NotePropertyValue @($shot.Description)
                            $uploads.Add($result)
                        }
                    }
                    $screenshotsByLabel[$label.Index] = $uploads.ToArray()
                }
            }
        }

        # 6. Pages ------------------------------------------------------------
        $overviewBody = New-OverviewPageBody -MarkdownSection $markdownSections `
                                           -NavigationLink $navigationLinks `
                                           -PageName $OverviewPageName `
                                           -Title $SiteTitle `
                                           -Background $backgrounds.Home
        Save-SitePage -SiteId $siteId -Body $overviewBody -Recreate:$Force | Out-Null

        if ($OverviewPageName -ne 'Home.aspx') {
            Write-Log "Overview content was written to '$OverviewPageName', but Microsoft Graph cannot change a site's welcome page. Set '$OverviewPageName' as the home page from the Pages library ('Make homepage') if that is what you want." -Level WARNING
        }

        $labelsBody = New-LabelsPageBody -Label $labels -PageName $LabelsPageName -SitePagesUrl "$siteUrl/SitePages" `
                                         -Background $backgrounds.Labels
        Save-SitePage -SiteId $siteId -Body $labelsBody -Recreate:$Force | Out-Null

        foreach ($label in $labels) {
            $shots = @(if ($screenshotsByLabel.ContainsKey($label.Index)) { $screenshotsByLabel[$label.Index] } else { @() })
            $labelBody = New-LabelPageBody -Label $label -Screenshot $shots -LabelsPageUrl "$siteUrl/SitePages/$LabelsPageName"
            Save-SitePage -SiteId $siteId -Body $labelBody -Recreate:$Force | Out-Null
        }

        $troubleshootingBody = New-TroubleshootingPageBody -MarkdownSection $troubleshootingSections `
                                                           -PageName $TroubleshootingPageName `
                                                           -Background $backgrounds.Troubleshooting
        Save-SitePage -SiteId $siteId -Body $troubleshootingBody -Recreate:$Force | Out-Null

        Write-Log 'SharePoint site pages are flat: the label pages live alongside the Labels page in the Pages library rather than underneath it. They read as subpages because the Labels page links to each one, and because the navigation nests them under Labels once the navigation below is applied.' -Level VERBOSE

    }
}
catch {
    Stop-WithError 'Unhandled error.' $_
}

# 8. Summary -----------------------------------------------------------------
$divider = '=' * 72

Write-Information ''
Write-Information $divider
Write-Information ' RESULT'
Write-Information $divider
Write-Information " Site URL : $($script:Summary.SiteUrl)"
Write-Information " Site      : $(if ($script:Summary.SiteCreated) { 'created by this run' } else { 'already existed; pages reconciled' })"
Write-Information " Uploads   : $($script:Summary.Uploads) image(s) (backgrounds and screenshots)"
Write-Information ''
Write-Information ' Pages'
foreach ($page in $script:Summary.Pages) {
    Write-Information ("   - {0,-34} {1}" -f $page.Name, $page.Title)
}
Write-Information ''
Write-Information ' Top navigation (apply by hand - see warnings above)'
foreach ($link in $script:Summary.NavigationLinks) {
    Write-Information ("   - {0,-22} -> {1}" -f $link.Title, $link.Url)
    Write-Information ("     {0}" -f $link.Description)
}

if ($script:Summary.Warnings.Count -gt 0) {
    Write-Information ''
    Write-Information " Warnings : $($script:Summary.Warnings.Count) (see the log)"
}

Write-Information ''
Write-Information " Log file : $script:LogFilePath"
Write-Information $divider

Write-Log 'Run completed.' -Level SUCCESS
exit 0

#endregion
