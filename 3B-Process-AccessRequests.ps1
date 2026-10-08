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

# ---------- Script (version B: adds tracker check, Processing Log and the archive) ----------
$scriptRoot = $PSScriptRoot
if (-not $scriptRoot) { $scriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Path }
if (-not $scriptRoot) { $scriptRoot = (Get-Location).Path }
$LogDir = Join-Path $scriptRoot 'Logs'
$inv = [Globalization.CultureInfo]::InvariantCulture
$siteHost = ([Uri]$SiteUrl).Host
$SnapFields = @('ID','SiteUrl','RequestStatus','AccessStartDate','AccessEndDate','AccessRevokedDate','AccessGrantedDate',
                'GrantedPrincipal','LastResult')

function Write-Line {
    param([string]$Text, [string]$Color = 'Gray')
    Write-Host ('[{0}] {1}' -f (Get-Date).ToString('HH:mm:ss'), $Text) -ForegroundColor $Color
    try {
        $file = Join-Path $LogDir ('AccessRequests_{0}_{1}.log' -f $env:COMPUTERNAME, (Get-Date).ToString('yyyy-MM-dd'))
        Add-Content -LiteralPath $file -Value ('{0} {1}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $Text)
    } catch {}
}

function ConvertTo-Utc {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) {
        if ($Value.Kind -eq [DateTimeKind]::Local) { return $Value.ToUniversalTime() }
        return [datetime]::SpecifyKind($Value, [DateTimeKind]::Utc)   # PnP read-back is UTC with Kind=Unspecified
    }
    $s = [string]$Value
    if ($s -eq '') { return $null }
    return ([DateTimeOffset]::Parse($s, $inv)).UtcDateTime
}

function Get-FieldValue($Fv, [string]$Name) {
    if ($null -eq $Fv) { return $null }
    if ($Fv.ContainsKey($Name)) { return $Fv[$Name] }
    return $null
}

function Get-LookupText($Value) {
    if ($null -eq $Value) { return '' }
    if ($Value -is [Microsoft.SharePoint.Client.FieldLookupValue]) { return [string]$Value.LookupValue }
    return [string]$Value
}

# Reduces any URL inside a site (pages, lists, trailing slash) to the site collection URL.
function Get-SiteRoot([string]$Url) {
    if (-not $Url) { return '' }
    $t = ($Url.Trim() -split '\s+')[0]
    $u = $null
    if (-not [Uri]::TryCreate($t, [UriKind]::Absolute, [ref]$u)) { return $t.TrimEnd('/') }
    $path = [Uri]::UnescapeDataString($u.AbsolutePath)
    $m = [regex]::Match($path, '^/(sites|teams|personal)/[^/]+', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $p = ''
    if ($m.Success) { $p = $m.Value }
    return ('{0}://{1}{2}' -f $u.Scheme, $u.Host, $p)
}

function Get-SiteKey([string]$Url) {
    return (Get-SiteRoot $Url).ToLowerInvariant()
}

# Site URLs from the tracker (a multi-line text column, so it cannot be filtered server side):
# read in pages and kept in memory, refreshed every $TrackerRefreshMinutes.
function Update-TrackerCache([double]$MaxAgeMinutes) {
    if ($null -ne $script:TrackerLoaded -and ((Get-Date) - $script:TrackerLoaded).TotalMinutes -lt $MaxAgeMinutes) { return }
    $set = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $rows = @(Get-PnPListItem -List $script:TrackerList -PageSize 5000 -Fields $TrackerUrlField -Connection $script:Main)
    foreach ($r in $rows) {
        $raw = [string](Get-FieldValue $r.FieldValues $TrackerUrlField)
        if (-not $raw) { continue }
        $raw = [System.Net.WebUtility]::HtmlDecode([regex]::Replace($raw, '<[^>]+>', ' '))
        foreach ($tok in ($raw -split '[\s;,]+')) {
            if ($tok -match '^https?://') { [void]$set.Add((Get-SiteKey $tok)) }
        }
    }
    $changed = ($null -eq $script:TrackerUrls -or $script:TrackerUrls.Count -ne $set.Count)
    $script:TrackerUrls = $set
    $script:TrackerLoaded = Get-Date
    if ($changed) { Write-Line ('Tracker: {0} site URL(s) from {1} item(s).' -f $set.Count, $rows.Count) 'DarkGray' }
}

function Test-InTracker([string]$Url) {
    return ($null -ne $script:TrackerUrls -and $script:TrackerUrls.Contains((Get-SiteKey $Url)))
}

function Get-RequestAction {
    param($Fv, [datetime]$NowUtc)
    $status = [string](Get-FieldValue $Fv 'RequestStatus')
    if ($status -eq 'Pending Access Grant') {
        $start = ConvertTo-Utc (Get-FieldValue $Fv 'AccessStartDate')
        if ($null -eq $start -or $start -le $NowUtc) { return 'Grant' }
        return 'Scheduled'
    }
    if ($status -eq 'Access Granted') {
        $end = ConvertTo-Utc (Get-FieldValue $Fv 'AccessEndDate')
        if ($null -ne $end -and $end -le $NowUtc) { return 'Revoke' }
        return ''
    }
    if ($status -eq 'Pending Access Removal') { return 'Revoke' }
    if ($status -eq 'Access Revoked') {
        if ($null -eq (ConvertTo-Utc (Get-FieldValue $Fv 'AccessRevokedDate'))) { return 'Revoke' }
        return 'Archive'
    }
    return ''
}

function Assert-SiteUrl([string]$Url) {
    if (-not $Url) { throw 'Site URL is empty.' }
    $u = $null
    if (-not [Uri]::TryCreate($Url.Trim(), [UriKind]::Absolute, [ref]$u)) { throw "Site URL '$Url' is not a valid URL." }
    if ($u.Scheme -ne 'https' -or $u.Host -ne $siteHost) { throw "Site URL '$Url' is not in this tenant ($siteHost)." }
}

# ----- connections -----
function Find-List([string]$Url) {
    $l = $null
    try { $l = Get-PnPList -Identity $Url -Connection $script:Main -ErrorAction SilentlyContinue } catch { $l = $null }
    if ($null -eq $l) {
        # URL lookup is not reliable in every PnP version: compare each list's real URL instead
        $want = '/' + $Url.Trim('/').ToLowerInvariant()
        foreach ($x in @(Get-PnPList -Connection $script:Main)) {
            $rel = Get-PnPProperty -ClientObject $x.RootFolder -Property ServerRelativeUrl -Connection $script:Main
            if ($rel.ToLowerInvariant().EndsWith($want)) { $l = $x; break }
        }
    }
    return $l
}

function Connect-Main {
    $script:Main = Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId -ReturnConnection
    $script:AdminConn = Connect-PnPOnline -Url $AdminUrl -Interactive -ClientId $ClientId -ReturnConnection
    $script:Tenant = $null
    $script:ReqList  = Find-List $RequestsListUrl
    $script:ArchList = Find-List $ArchiveListUrl
    if ($null -eq $script:ReqList)  { throw "List '$RequestsListUrl' not found." }
    if ($null -eq $script:ArchList) { throw "List '$ArchiveListUrl' not found." }
    $script:ReqListId  = $script:ReqList.Id
    $script:ArchListId = $script:ArchList.Id
    $script:TrackerList = Find-List $TrackerListUrl
    if ($null -eq $script:TrackerList) { throw "List '$TrackerListUrl' not found." }
    $script:TrackerLoaded = $null
    $skip = @('ContentType','Attachments','Created','Modified','Author','Editor','OriginalItemId','SourceKey','ArchivedDate','ArchivedBy')
    $types = @('Text','Note','Choice','MultiChoice','Number','Currency','DateTime','Boolean','User','UserMulti','Lookup','LookupMulti','URL')
    $script:CopyFields = @(Get-PnPField -List $script:ArchList -Connection $script:Main | Where-Object {
        (-not $_.Hidden) -and (-not $_.ReadOnlyField) -and ($skip -notcontains $_.InternalName) -and
        ($types -contains $_.TypeAsString) -and (-not $_.InternalName.StartsWith('_'))
    })
}

# SystemUpdate: Modified / Modified By keep showing the user's last edit
function Save-Request([int]$Id, [hashtable]$Values) {
    $ctx = $script:Main.Context
    $it = $ctx.Web.Lists.GetById($script:ReqListId).GetItemById($Id)
    foreach ($k in $Values.Keys) { $it[$k] = $Values[$k] }
    $it.SystemUpdate()
    $ctx.ExecuteQuery()
}

function Add-LogLine([string]$Existing, [string]$Message) {
    $line = '{0} UTC [{1}] {2}' -f [datetime]::UtcNow.ToString('yyyy-MM-dd HH:mm', $inv), $env:COMPUTERNAME, $Message
    $old = @(([string]$Existing) -split "`r?`n" | Where-Object { $_ -ne '' })
    $lines = @($line) + $old
    if ($lines.Count -gt $MaxLogLines) { $lines = $lines[0..($MaxLogLines - 1)] }
    return ($lines -join "`n")
}

function Limit-Text([string]$Text, [int]$Max) {
    if ($Text.Length -le $Max) { return $Text }
    return $Text.Substring(0, $Max)
}

function New-ResultValues {
    param($Fv, [string]$Message)
    $v = @{}
    $v['LastResult']    = Limit-Text $Message 255
    $v['ProcessingLog'] = Add-LogLine ([string](Get-FieldValue $Fv 'ProcessingLog')) $Message
    $v['LastProcessed'] = [datetime]::Now   # not Get-Date: its PSObject wrapper is sent as text and saved as UTC
    $v['ProcessedBy']   = $env:COMPUTERNAME
    return $v
}

# ----- site collection admin (via the tenant admin site, so the script account need not own the site) -----
function Get-Tenant {
    if ($null -eq $script:Tenant) {
        $ctx = $script:AdminConn.Context
        $dir = Split-Path -Parent $ctx.GetType().Assembly.Location
        # Build the Tenant object from the assembly PnP itself loaded; a type from another
        # CSOM copy on the machine does not bind to PnP's context.
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

function Set-SiteAdmin {
    param([string]$Url, [string]$Principal, [bool]$IsAdmin)
    $tenant = Get-Tenant
    [void]$tenant.SetSiteAdmin($Url.Trim().TrimEnd('/'), $Principal, $IsAdmin)
    try { $script:AdminConn.Context.ExecuteQuery() }
    catch {
        $m = $_.Exception.Message
        if ($m -match 'File Not Found') {
            if (-not $IsAdmin) { return 'site no longer exists' }
            throw "Site '$Url' does not exist."
        }
        if ((-not $IsAdmin) -and $m -match 'not found|does not exist|cannot be found|could not be found') { return 'was not an admin' }
        throw
    }
    if ($IsAdmin) { return 'added' }
    return 'removed'
}

# ----- archive -----
function Convert-FieldValue {
    param([string]$Type, $Value)
    if ($null -eq $Value) { return $null }
    # Person / lookup / hyperlink values are passed through as read from the source item. A value
    # built with New-Object can bind to a different CSOM assembly than PnP's and SharePoint
    # rejects it ("Invalid data has been used to update the list item").
    if (@('User','Lookup','URL') -contains $Type) { return $Value }
    if (@('UserMulti','LookupMulti') -contains $Type) {
        if (@($Value).Count -eq 0) { return $null }
        return ,$Value
    }
    if ($Type -eq 'DateTime') {
        $d = ConvertTo-Utc $Value
        if ($null -eq $d) { return $null }
        return $d.ToLocalTime()
    }
    if ($Type -eq 'MultiChoice') { return ,([string[]]@($Value)) }
    return $Value
}

# Copies the item to the archive, stamps Created / Created By / Modified / Modified By from the
# source and returns the archive id.
function Copy-ToArchive {
    param($Item)
    $ctx = $script:Main.Context
    $fv = $Item.FieldValues
    $key = '{0}:{1}' -f $script:ReqListId, $Item.Id
    $alist = $ctx.Web.Lists.GetById($script:ArchListId)

    $ci = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
    $new = $alist.AddItem($ci)
    foreach ($f in $script:CopyFields) {
        $v = Convert-FieldValue -Type $f.TypeAsString -Value (Get-FieldValue $fv $f.InternalName)
        if ($null -ne $v) { $new[$f.InternalName] = $v }
    }
    $new['OriginalItemId'] = [int]$Item.Id
    $new['SourceKey'] = $key
    $new['ArchivedDate'] = [datetime]::Now
    $new['ArchivedBy'] = $env:COMPUTERNAME
    $new['LastResult'] = 'Archived'
    $new['ProcessingLog'] = Add-LogLine ([string](Get-FieldValue $fv 'ProcessingLog')) 'Archived'
    $new.Update()
    $ctx.Load($new)
    $ctx.ExecuteQuery()
    $archId = [int]$new.Id

    $authorId = [int](Get-FieldValue $fv 'Author').LookupId
    $EditorId = [int](Get-FieldValue $fv 'Editor').LookupId
    $createdUtc = ConvertTo-Utc (Get-FieldValue $fv 'Created')
    $ModifiedUtc = ConvertTo-Utc (Get-FieldValue $fv 'Modified')

    $a = $alist.GetItemById($archId)
    $a['Author']   = $authorId          # bare int ids: see powershell-pnp skill
    $a['Editor']   = $EditorId
    $a['Created']  = $createdUtc.ToLocalTime()
    $a['Modified'] = $ModifiedUtc.ToLocalTime()
    $a.UpdateOverwriteVersion()
    $ctx.ExecuteQuery()

    return $archId
}

# ----- one request -----
function Invoke-Request {
    param($Snap, $Items)
    $id = [int]$Snap.Id
    $fv = $null
    $action = ''
    try {
        $fv = (Get-PnPListItem -List $script:ReqList -Id $id -Connection $script:Main).FieldValues
        $now = [datetime]::UtcNow
        $action = Get-RequestAction -Fv $fv -NowUtc $now
        $site = Get-LookupText (Get-FieldValue $fv 'SiteUrl')

        if ($action -eq 'Grant') {
            $site = Get-SiteRoot $site
            Assert-SiteUrl $site
            if (-not (Test-InTracker $site)) { throw "Site URL '$site' is not in the Teams Attestation Tracker. Fix the Site URL and save the item." }
            [void](Set-SiteAdmin -Url $site -Principal $GroupClaim -IsAdmin $true)
            $msg = "Added $GroupClaim as site collection admin on $site"
            $vals = New-ResultValues -Fv $fv -Message $msg
            $vals['RequestStatus']     = 'Access Granted'
            $vals['AccessGrantedDate'] = [datetime]::Now
            $vals['GrantedPrincipal']  = $GroupClaim
            Save-Request -Id $id -Values $vals
            Write-Line ('#{0} GRANTED  site admin -> {1}' -f $id, $site) 'Green'
            return 'Done'
        }

        if ($action -eq 'Revoke') {
            $site = Get-SiteRoot $site
            $principal = [string](Get-FieldValue $fv 'GrantedPrincipal')
            $grantedAt = ConvertTo-Utc (Get-FieldValue $fv 'AccessGrantedDate')
            if (-not $principal -and $null -ne $grantedAt) { $principal = $GroupClaim }
            if (-not $principal) {
                $msg = 'Access was never granted; nothing to remove.'
            } else {
                Assert-SiteUrl $site
                $res = Set-SiteAdmin -Url $site -Principal $principal -IsAdmin $false
                $msg = "Removed $principal as site collection admin on $site ($res)"
            }
            $vals = New-ResultValues -Fv $fv -Message $msg
            $vals['RequestStatus']     = 'Access Revoked'
            $vals['AccessRevokedDate'] = [datetime]::Now
            Save-Request -Id $id -Values $vals
            Write-Line ('#{0} REVOKED  site admin -> {1}' -f $id, $site) 'Yellow'
            $action = 'Archive'
        }

        if ($action -eq 'Archive') {
            $final = Get-PnPListItem -List $script:ReqList -Id $id -Connection $script:Main
            $archId = Copy-ToArchive -Item $final
            Remove-PnPListItem -List $script:ReqList -Identity $id -Recycle -Force -Connection $script:Main | Out-Null
            Write-Line ('#{0} ARCHIVED -> archive #{1}' -f $id, $archId) 'Yellow'
            return 'Done'
        }

        return 'Skipped'
    }
    catch {
        # retried every cycle; written back (and shown) only when the error changes
        $err = Limit-Text ('ERROR ({0}): {1}' -f $action, $_.Exception.Message) 255
        if ([string](Get-FieldValue $fv 'LastResult') -eq $err) { return 'Failed' }
        $vals = New-ResultValues -Fv $fv -Message $err
        try { Save-Request -Id $id -Values $vals }
        catch { Write-Line ('#{0} could not write the error back: {1}' -f $id, $_.Exception.Message) 'Red' }
        Write-Line ('#{0} {1}' -f $id, $err) 'Red'
        return 'Failed'
    }
}

function Invoke-Cycle {
    $now = [datetime]::UtcNow
    try { Update-TrackerCache $TrackerRefreshMinutes }
    catch { Write-Line ('Tracker refresh failed (using the previous copy): {0}' -f $_.Exception.Message) 'Red' }
    $items = @(Get-PnPListItem -List $script:ReqList -PageSize $PageSize -Fields $SnapFields -Connection $script:Main)
    $stat = @{ Open = $items.Count; Due = 0; Done = 0; Failed = 0; Scheduled = 0 }
    $due = [System.Collections.Generic.List[object]]::new()
    foreach ($it in $items) {
        $fv = $it.FieldValues
        $a = Get-RequestAction -Fv $fv -NowUtc $now
        if ($a -eq 'Scheduled') { $stat['Scheduled'] += 1; continue }
        if (-not $a) { continue }
        $due.Add($it)
    }
    $stat['Due'] = $due.Count
    $order = $due.ToArray()
    for ($i = 0; $i -lt $order.Count; $i++) {
        $r = Invoke-Request -Snap $order[$i] -Items $items
        if ($r -eq 'Done')   { $stat['Done'] += 1 }
        if ($r -eq 'Failed') { $stat['Failed'] += 1 }
    }
    return $stat
}

# ---------- Start ----------
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host 'PnP.PowerShell is not installed. Run: Install-Module PnP.PowerShell -Scope CurrentUser' -ForegroundColor Red
    exit 1
}
if ($GroupClaim -match '00000000-0000-0000-0000-000000000000') {
    Write-Host 'Set $GroupClaim at the top of the script to the Entra ID group claim (c:0t.c|tenant|<group id>).' -ForegroundColor Red
    exit 1
}
if (-not (Test-Path -LiteralPath $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }

Write-Host ''
Write-Host 'Access request processor (B)' -ForegroundColor Cyan
Write-Host ("  Site   : {0}" -f $SiteUrl)
Write-Host ("  Group  : {0}" -f $GroupClaim)
Write-Host ("  Every  : {0}s   Ctrl+C to stop" -f $PollSeconds)
Write-Host ''

Connect-Main
Write-Line ('Connected. {0} archive column(s) will be copied.' -f $script:CopyFields.Count) 'Cyan'

$cycle = 0
$quiet = 0
while ($true) {
    $cycle++
    try {
        $s = Invoke-Cycle
        $summary = '{0} open | {1} due | {2} done | {3} failed | {4} scheduled' -f `
            $s['Open'], $s['Due'], $s['Done'], $s['Failed'], $s['Scheduled']
        if ($s['Done'] -gt 0) {
            Write-Line $summary 'Cyan'
            $quiet = 0
        } else {
            $quiet++
            if ($cycle -eq 1 -or $quiet -ge $HeartbeatEvery) { Write-Line ('idle: ' + $summary) 'DarkGray'; $quiet = 0 }
        }
    }
    catch {
        Write-Line ('Cycle failed: {0}' -f $_.Exception.Message) 'Red'
    }
    Start-Sleep -Seconds $PollSeconds
}
