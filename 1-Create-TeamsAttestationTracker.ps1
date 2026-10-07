$ErrorActionPreference = 'Stop'

# ---------- Settings (DEV ONLY - prod already has this list) ----------
$SiteUrl        = 'https://gocleverpointcom.sharepoint.com/sites/SharePointServices'
$ClientId       = 'e391b4e0-0151-4aa2-8ce7-dccf4b3921fa'
$DevTenantHost  = 'gocleverpointcom.sharepoint.com'      # script refuses to run against any other host
$ListUrl        = 'Lists/TeamsAttestationTracker'
$ListTitle      = 'Teams Attestation Tracker'
$SampleSiteUrls = @(
    'https://gocleverpointcom.sharepoint.com/sites/SharePointServices'
)

# ---------- Script ----------
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host 'PnP.PowerShell is not installed. Run: Install-Module PnP.PowerShell -Scope CurrentUser' -ForegroundColor Red
    exit 1
}
if (([Uri]$SiteUrl).Host -ne $DevTenantHost) {
    Write-Host "This script only runs in DEV ($DevTenantHost). Site '$SiteUrl' is not DEV." -ForegroundColor Red
    exit 1
}

function Get-ListSafe([string]$Identity) {
    $l = $null
    try { $l = Get-PnPList -Identity $Identity -ErrorAction SilentlyContinue } catch { $l = $null }
    return $l
}

function Add-FieldIfMissing {
    param($List, [string]$InternalName, [string]$Xml)
    $f = $null
    try { $f = Get-PnPField -List $List -Identity $InternalName -ErrorAction SilentlyContinue } catch { $f = $null }
    if ($null -ne $f) { Write-Host "  [skip] $InternalName"; return }
    $xmlWithId = $Xml.Replace('#ID#', ('{' + [guid]::NewGuid().ToString() + '}'))
    Add-PnPFieldFromXml -List $List -FieldXml $xmlWithId | Out-Null
    $f = $null
    try { $f = Get-PnPField -List $List -Identity $InternalName -ErrorAction SilentlyContinue } catch { $f = $null }
    if ($null -eq $f) { throw "Column '$InternalName' was not found after it was created." }
    Write-Host "  [add ] $InternalName" -ForegroundColor Green
}

function Set-ViewSafe {
    param($List, [string]$Title, [string[]]$Fields, [string]$Query, [switch]$Default)
    $v = $null
    try { $v = Get-PnPView -List $List -Identity $Title -ErrorAction SilentlyContinue } catch { $v = $null }
    if ($null -eq $v) {
        if ($Default) { Add-PnPView -List $List -Title $Title -Fields $Fields -Query $Query -RowLimit 100 -SetAsDefault | Out-Null }
        else          { Add-PnPView -List $List -Title $Title -Fields $Fields -Query $Query -RowLimit 100 | Out-Null }
        Write-Host "  [add ] view '$Title'" -ForegroundColor Green
    } else {
        Set-PnPView -List $List -Identity $Title -Fields $Fields | Out-Null
        if ($Query) { Set-PnPView -List $List -Identity $Title -Values @{ ViewQuery = $Query } | Out-Null }
        Write-Host "  [upd ] view '$Title'"
    }
}

Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

$list = Get-ListSafe $ListUrl
if ($null -eq $list) {
    New-PnPList -Title $ListTitle -Url $ListUrl -Template GenericList | Out-Null
    $list = Get-ListSafe $ListUrl
    if ($null -eq $list) { throw "List '$ListUrl' was not found after New-PnPList." }
    Write-Host "List created: $ListTitle" -ForegroundColor Green
} else {
    Write-Host "List exists: $($list.Title)"
}
Set-PnPList -Identity $list -EnableVersioning $true -MajorVersions 50 -EnableAttachments $false | Out-Null
Set-PnPField -List $list -Identity 'Title' -Values @{ Title = 'Site Title'; Required = $false } | Out-Null

Write-Host 'Columns:' -ForegroundColor Cyan
Add-FieldIfMissing $list 'SiteUrl'             "<Field Type='Text' Name='SiteUrl' StaticName='SiteUrl' DisplayName='Site Url' ID='#ID#' Required='TRUE' Indexed='TRUE' MaxLength='255' />"
Add-FieldIfMissing $list 'PrimaryOwner'        "<Field Type='Text' Name='PrimaryOwner' StaticName='PrimaryOwner' DisplayName='Primary Owner' ID='#ID#' MaxLength='255' />"
Add-FieldIfMissing $list 'SecondaryOwners'     "<Field Type='Note' Name='SecondaryOwners' StaticName='SecondaryOwners' DisplayName='Secondary Owners' ID='#ID#' NumLines='3' RichText='FALSE' />"
Add-FieldIfMissing $list 'CostCenter'          "<Field Type='Text' Name='CostCenter' StaticName='CostCenter' DisplayName='Cost Center' ID='#ID#' MaxLength='50' Indexed='TRUE' />"
Add-FieldIfMissing $list 'LastAttestationDate' "<Field Type='DateTime' Name='LastAttestationDate' StaticName='LastAttestationDate' DisplayName='Last Attestation Date' ID='#ID#' Format='DateOnly' />"
Add-FieldIfMissing $list 'AttestationResult'   "<Field Type='Choice' Name='AttestationResult' StaticName='AttestationResult' DisplayName='Attestation Result' ID='#ID#' Format='Dropdown' Indexed='TRUE'><Default>Pending</Default><CHOICES><CHOICE>Pending</CHOICE><CHOICE>Confirmed</CHOICE><CHOICE>Not Confirmed</CHOICE><CHOICE>Site Deleted</CHOICE></CHOICES></Field>"

Write-Host 'Views:' -ForegroundColor Cyan
$viewFields = @('ID','LinkTitle','SiteUrl','PrimaryOwner','SecondaryOwners','CostCenter','Modified','LastAttestationDate','Editor','AttestationResult')
Set-ViewSafe -List $list -Title 'All Items' -Fields $viewFields -Query "<OrderBy><FieldRef Name='Title' /></OrderBy>"
Set-ViewSafe -List $list -Title 'Active' -Fields $viewFields -Query "<Where><Neq><FieldRef Name='AttestationResult' /><Value Type='Choice'>Site Deleted</Value></Neq></Where><OrderBy><FieldRef Name='Title' /></OrderBy>"
Set-ViewSafe -List $list -Title 'Attestation Required' -Fields $viewFields -Query "<Where><Eq><FieldRef Name='AttestationResult' /><Value Type='Choice'>Pending</Value></Eq></Where><OrderBy><FieldRef Name='LastAttestationDate' /></OrderBy>"
Set-ViewSafe -List $list -Title 'Cost Center' -Fields $viewFields -Query "<GroupBy Collapse='TRUE'><FieldRef Name='CostCenter' /></GroupBy><OrderBy><FieldRef Name='Title' /></OrderBy>"

Write-Host 'Sample rows:' -ForegroundColor Cyan
$have = @{}
foreach ($it in @(Get-PnPListItem -List $list -PageSize 5000 -Fields 'SiteUrl')) { $have[([string]$it['SiteUrl']).Trim().ToLowerInvariant()] = $true }
for ($i = 0; $i -lt $SampleSiteUrls.Count; $i++) {
    $url = $SampleSiteUrls[$i]
    if ($have.ContainsKey($url.Trim().ToLowerInvariant())) { Write-Host "  [skip] $url"; continue }
    $leaf = ($url.TrimEnd('/') -split '/')[-1]
    Add-PnPListItem -List $list -Values @{ Title = $leaf; SiteUrl = $url; CostCenter = '0000'; AttestationResult = 'Pending' } | Out-Null
    Write-Host "  [add ] $url" -ForegroundColor Green
}

Write-Host "Done. $SiteUrl/$ListUrl" -ForegroundColor Green
exit 0
