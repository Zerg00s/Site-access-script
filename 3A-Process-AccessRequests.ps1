$ErrorActionPreference = 'Stop'

# ---------- Settings ----------
$SiteUrl           = 'https://gocleverpointcom.sharepoint.com/sites/SharePointServices'
$AdminUrl          = 'https://gocleverpointcom-admin.sharepoint.com'
$ClientId          = 'e391b4e0-0151-4aa2-8ce7-dccf4b3921fa'
$RequestsListUrl   = 'Lists/AccessRequests'
$ArchiveListUrl    = 'Lists/AccessRequestsArchive'
$TrackerListUrl    = 'Lists/TeamsAttestationTracker'   # requests are only granted for sites listed here
$TrackerUrlField   = 'SiteUrl'
$TrackerRefreshMinutes = 10
$GroupClaim        = 'c:0t.c|tenant|38f84872-fa14-440f-bed8-72b7e033445d'   # Entra ID group added as site collection admin
$PollSeconds       = 30
$LockMinutes       = 10      # how long one server owns an item while working on it
$MaxBackoffMinutes = 60      # failed items retry after 2, 4, 8 ... minutes, capped here
$HeartbeatEvery    = 10      # print an idle status line every N quiet cycles
$MaxLogLines       = 40      # lines kept in each item's Processing Log
$PageSize          = 2000

# ---------- Script (version A: add / remove the group as site collection admin) ----------
function Get-FieldValue($Fv, [string]$Name) {
    if ($Fv.ContainsKey($Name)) { return $Fv[$Name] }
    return $null
}

function ConvertTo-Utc($Value) {
    if ($null -eq $Value) { return $null }
    return [datetime]::SpecifyKind([datetime]$Value, [DateTimeKind]::Utc)   # PnP returns UTC with Kind=Unspecified
}

function Find-List([string]$Url) {
    $l = $null
    try { $l = Get-PnPList -Identity $Url -Connection $script:Main -ErrorAction SilentlyContinue } catch { $l = $null }
    if ($null -eq $l) {
        $want = '/' + $Url.Trim('/').ToLowerInvariant()
        foreach ($x in @(Get-PnPList -Connection $script:Main)) {
            $rel = Get-PnPProperty -ClientObject $x.RootFolder -Property ServerRelativeUrl -Connection $script:Main
            if ($rel.ToLowerInvariant().EndsWith($want)) { $l = $x; break }
        }
    }
    if ($null -eq $l) { throw "List '$Url' not found." }
    return $l
}

# Tenant object built from the CSOM assembly PnP loaded (another CSOM copy on the machine does not bind)
function Get-Tenant {
    if ($null -eq $script:Tenant) {
        $ctx = $script:AdminConn.Context
        $dir = Split-Path -Parent $ctx.GetType().Assembly.Location
        $asm = $null
        foreach ($x in [AppDomain]::CurrentDomain.GetAssemblies()) {
            if ($x.IsDynamic -or $x.GetName().Name -ne 'Microsoft.Online.SharePoint.Client.Tenant') { continue }
            if ((Split-Path -Parent $x.Location) -eq $dir) { $asm = $x; break }
        }
        if ($null -eq $asm) { throw "Tenant CSOM assembly not found next to $dir." }
        $type = $asm.GetType('Microsoft.Online.SharePoint.TenantAdministration.Tenant')
        $script:Tenant = $type.GetConstructors()[0].Invoke(@($ctx.PSObject.BaseObject))
    }
    return $script:Tenant
}

function Set-SiteAdmin([string]$Url, [bool]$IsAdmin) {
    $tenant = Get-Tenant
    [void]$tenant.SetSiteAdmin($Url, $GroupClaim, $IsAdmin)
    $script:AdminConn.Context.ExecuteQuery()
}

# SystemUpdate: Modified / Modified By keep showing the user's last edit
function Save-Result([int]$Id, [hashtable]$Values) {
    Set-PnPListItem -List $script:ReqList -Identity $Id -Values $Values -UpdateType SystemUpdate -Connection $script:Main | Out-Null
}

function Write-Line([string]$Text, [string]$Color = 'Gray') {
    Write-Host ('[{0}] {1}' -f (Get-Date).ToString('HH:mm:ss'), $Text) -ForegroundColor $Color
}

# ---------- Start ----------
Write-Host ''
Write-Host 'Access request processor (A)' -ForegroundColor Cyan
Write-Host ("  Site: {0}   Every: {1}s   Ctrl+C to stop" -f $SiteUrl, $PollSeconds)
Write-Host ''

$script:Main      = Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId -ReturnConnection
$script:AdminConn = Connect-PnPOnline -Url $AdminUrl -Interactive -ClientId $ClientId -ReturnConnection
$script:Tenant    = $null
$script:ReqList   = Find-List $RequestsListUrl
Write-Line 'Connected.' 'Cyan'

$fields = @('ID','SiteUrl','RequestStatus','AccessStartDate','AccessEndDate','LastResult')
while ($true) {
    try {
        $now = [datetime]::UtcNow
        $items = @(Get-PnPListItem -List $script:ReqList -PageSize $PageSize -Fields $fields -Connection $script:Main)
        foreach ($it in $items) {
            $fv = $it.FieldValues
            $status = [string](Get-FieldValue $fv 'RequestStatus')
            $site = ([string](Get-FieldValue $fv 'SiteUrl')).Trim().TrimEnd('/')
            $start = ConvertTo-Utc (Get-FieldValue $fv 'AccessStartDate')
            $end = ConvertTo-Utc (Get-FieldValue $fv 'AccessEndDate')

            $grant = ($status -eq 'Pending Access Grant' -and ($null -eq $start -or $start -le $now))
            $remove = ($status -eq 'Pending Access Removal' -or ($status -eq 'Access Granted' -and $null -ne $end -and $end -le $now))
            if (-not $grant -and -not $remove) { continue }

            try {
                $stamp = [datetime]::Now   # not Get-Date: its PSObject wrapper is saved 4 hours off
                if ($grant) {
                    Set-SiteAdmin $site $true
                    $msg = "Added $GroupClaim as site collection admin on $site"
                    Save-Result $it.Id @{ RequestStatus = 'Access Granted'; AccessGrantedDate = $stamp; GrantedPrincipal = $GroupClaim; LastResult = $msg }
                    Write-Line ('#{0} GRANTED  {1}' -f $it.Id, $site) 'Green'
                } else {
                    Set-SiteAdmin $site $false
                    $msg = "Removed $GroupClaim as site collection admin on $site"
                    Save-Result $it.Id @{ RequestStatus = 'Access Revoked'; AccessRevokedDate = $stamp; LastResult = $msg }
                    Write-Line ('#{0} REVOKED  {1}' -f $it.Id, $site) 'Yellow'
                }
            }
            catch {
                # retried every cycle; written back (and shown) only when the error changes
                $err = 'ERROR: ' + $_.Exception.Message
                if ($err.Length -gt 255) { $err = $err.Substring(0, 255) }
                if ([string](Get-FieldValue $fv 'LastResult') -ne $err) {
                    try { Save-Result $it.Id @{ LastResult = $err } } catch {}
                    Write-Line ('#{0} {1}' -f $it.Id, $err) 'Red'
                }
            }
        }
    }
    catch { Write-Line ('Cycle failed: {0}' -f $_.Exception.Message) 'Red' }
    Start-Sleep -Seconds $PollSeconds
}
