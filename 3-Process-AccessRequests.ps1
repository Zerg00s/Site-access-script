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

# ---------- Script ----------
$scriptRoot = $PSScriptRoot
if (-not $scriptRoot) { $scriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Path }
if (-not $scriptRoot) { $scriptRoot = (Get-Location).Path }
$LogDir = Join-Path $scriptRoot 'Logs'
$inv = [Globalization.CultureInfo]::InvariantCulture
$siteHost = ([Uri]$SiteUrl).Host
$RunId = '{0}:{1}:{2}' -f $env:COMPUTERNAME, $PID, ([guid]::NewGuid().ToString('N').Substring(0, 6))
$SnapFields = @('ID','SiteUrl','RequestStatus','AccessStartDate','AccessEndDate','AccessRevokedDate','AccessGrantedDate',
                'GrantedPrincipal','NextAttempt','Editor','LockOwner','LockExpires','Modified')

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
    $key = Get-SiteKey $Url
    if ($null -ne $script:TrackerUrls -and $script:TrackerUrls.Contains($key)) { return $true }
    Update-TrackerCache 1          # maybe added to the tracker since the last refresh
    return $script:TrackerUrls.Contains($key)
}

function Get-HttpStatus($ErrorRecord) {
    $code = 0
    try { $code = [int]$ErrorRecord.Exception.Response.StatusCode } catch { $code = 0 }
    return $code
}

function Get-RequestAction {
    param($Fv, [datetime]$NowUtc)
    $status = [string](Get-FieldValue $Fv 'RequestStatus')
    $immediate = ($true -eq (Get-FieldValue $Fv 'RevokeAccessImmediately'))   # Yes = remove now, ignore the removal date
    if ($status -eq 'Pending Access Grant') {
        if ($immediate) { return 'Revoke' }                                   # cancelled before it was granted
        $start = ConvertTo-Utc (Get-FieldValue $Fv 'AccessStartDate')
        if ($null -eq $start -or $start -le $NowUtc) { return 'Grant' }
        return 'Scheduled'
    }
    if ($status -eq 'Access Granted') {
        $end = ConvertTo-Utc (Get-FieldValue $Fv 'AccessEndDate')
        if ($immediate -or ($null -ne $end -and $end -le $NowUtc)) { return 'Revoke' }
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

# ----- connections and REST -----
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
    $me = Get-PnPProperty -ClientObject (Get-PnPWeb -Connection $script:Main) -Property CurrentUser -Connection $script:Main
    $script:MyUserId = [int]$me.Id
    $script:SnapFields = @($SnapFields)
    $col = $null
    try { $col = Get-PnPField -List $script:ReqList -Identity 'RevokeAccessImmediately' -Connection $script:Main -ErrorAction SilentlyContinue } catch { $col = $null }
    if ($null -ne $col) { $script:SnapFields += 'RevokeAccessImmediately' }
    else { Write-Line 'Column RevokeAccessImmediately not found on the requests list; removal dates only.' 'Yellow' }
    $skip = @('ContentType','Attachments','Created','Modified','Author','Editor','OriginalItemId','SourceKey','ArchivedDate','ArchivedBy')
    $types = @('Text','Note','Choice','MultiChoice','Number','Currency','DateTime','Boolean','User','UserMulti','Lookup','LookupMulti','URL')
    $script:CopyFields = @(Get-PnPField -List $script:ArchList -Connection $script:Main | Where-Object {
        (-not $_.Hidden) -and (-not $_.ReadOnlyField) -and ($skip -notcontains $_.InternalName) -and
        ($types -contains $_.TypeAsString) -and (-not $_.InternalName.StartsWith('_'))
    })
}

function Get-SpoToken {
    $t = $null
    try { $t = Get-PnPAccessToken -ResourceTypeName SharePoint -Connection $script:Main } catch { $t = $null }
    if (-not $t) { $t = Get-PnPAppAuthAccessToken -Connection $script:Main }
    return [string]$t
}

function Invoke-Spo {
    param([string]$Method, [string]$Url, [hashtable]$Headers, [string]$Body)
    $auth = 'Bearer ' + (Get-SpoToken)
    $h = @{ 'Authorization' = $auth; 'Accept' = 'application/json;odata=nometadata' }
    if ($Headers) { foreach ($k in $Headers.Keys) { $h[$k] = $Headers[$k] } }
    $p = @{ Uri = $Url; Method = $Method; Headers = $h; UseBasicParsing = $true; ErrorAction = 'Stop' }
    if ($Body) {
        $p['Body'] = [System.Text.Encoding]::UTF8.GetBytes($Body)
        $p['ContentType'] = 'application/json;odata=nometadata'
    }
    if ($PSVersionTable.PSVersion.Major -ge 6) { $p['SkipHeaderValidation'] = $true }
    return Invoke-WebRequest @p
}

# Claims an item for this process. ETag (If-Match) makes the claim atomic:
# if another server wrote the item between our read and our write, SharePoint returns 412.
function Lock-Request([int]$Id) {
    $base = "$SiteUrl/_api/web/lists(guid'$($script:ReqListId)')/items($Id)"
    $r = $null
    try { $r = Invoke-Spo -Method 'GET' -Url ($base + '?$select=LockOwner,LockExpires') }
    catch { if ((Get-HttpStatus $_) -eq 404) { return @{ State = 'Gone' } }; throw }
    $etag = $r.Headers['ETag']
    if ($etag -is [array]) { $etag = $etag[0] }
    $json = ([string]$r.Content -creplace ',\s*"ID":\s*\d+', '') -creplace '"ID":\s*\d+\s*,?', ''   # PS 5.1 ConvertFrom-Json rejects Id + ID
    $j = $json | ConvertFrom-Json
    $owner = [string]$j.LockOwner
    $exp = ConvertTo-Utc $j.LockExpires
    if ($owner -and $owner -ne $RunId -and $null -ne $exp -and $exp -gt [datetime]::UtcNow) { return @{ State = 'Busy' } }

    $newExp = [datetime]::UtcNow.AddMinutes($LockMinutes).ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", $inv)
    $body = @{ LockOwner = $RunId; LockExpires = $newExp } | ConvertTo-Json -Compress
    $hdr = @{ 'X-HTTP-Method' = 'MERGE'; 'If-Match' = [string]$etag }
    try { Invoke-Spo -Method 'POST' -Url $base -Headers $hdr -Body $body | Out-Null }
    catch {
        $c = Get-HttpStatus $_
        if ($c -eq 412) { return @{ State = 'Busy' } }
        if ($c -eq 404) { return @{ State = 'Gone' } }
        throw
    }
    return @{ State = 'Locked' }
}

# Normal update: adds a version, Modified By = the account running the script
function Save-Request([int]$Id, [hashtable]$Values) {
    $ctx = $script:Main.Context
    $it = $ctx.Web.Lists.GetById($script:ReqListId).GetItemById($Id)
    foreach ($k in $Values.Keys) { $it[$k] = $Values[$k] }
    $it.Update()
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
    param($Fv, [string]$Message, [switch]$KeepLock)
    $v = @{}
    $v['LastResult']    = Limit-Text $Message 255
    $v['ProcessingLog'] = Add-LogLine ([string](Get-FieldValue $Fv 'ProcessingLog')) $Message
    $v['LastProcessed'] = [datetime]::Now   # not Get-Date: its PSObject wrapper is sent as text and saved as UTC
    $v['ProcessedBy']   = $env:COMPUTERNAME
    $v['AttemptCount']  = 0
    $v['NextAttempt']   = $null
    if (-not $KeepLock) { $v['LockOwner'] = $null; $v['LockExpires'] = $null }
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

# Another active request for the same site and group must keep the group as admin.
# Requests that are themselves due for removal do not count.
function Find-OtherActiveRequest {
    param($Items, [int]$SelfId, [string]$SiteKey, [string]$Principal, [datetime]$NowUtc)
    foreach ($o in $Items) {
        if ([int]$o.Id -eq $SelfId) { continue }
        $f = $o.FieldValues
        if ([string](Get-FieldValue $f 'RequestStatus') -ne 'Access Granted') { continue }
        $end = ConvertTo-Utc (Get-FieldValue $f 'AccessEndDate')
        if ($null -ne $end -and $end -le $NowUtc) { continue }
        if ($true -eq (Get-FieldValue $f 'RevokeAccessImmediately')) { continue }
        if ((Get-SiteKey (Get-LookupText (Get-FieldValue $f 'SiteUrl'))) -ne $SiteKey) { continue }
        $p = [string](Get-FieldValue $f 'GrantedPrincipal')
        if (-not $p) { $p = $GroupClaim }
        if ($p -eq $Principal) { return [int]$o.Id }
    }
    return 0
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

# Copies the item to the archive (reusing a half-finished copy from an earlier attempt), stamps
# Created / Created By / Modified / Modified By from the source, verifies them, returns the archive id.
function Copy-ToArchive {
    param($Item)
    $ctx = $script:Main.Context
    $fv = $Item.FieldValues
    $key = '{0}:{1}' -f $script:ReqListId, $Item.Id
    $q = "<View><Query><Where><Eq><FieldRef Name='SourceKey' /><Value Type='Text'>$key</Value></Eq></Where></Query><RowLimit>1</RowLimit></View>"
    $found = @(Get-PnPListItem -List $script:ArchList -Query $q -Connection $script:Main)
    $alist = $ctx.Web.Lists.GetById($script:ArchListId)

    $archId = 0
    if ($found.Count -gt 0) { $archId = [int]$found[0].Id }
    else {
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
    }

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

    $chk = (Get-PnPListItem -List $script:ArchList -Id $archId -Fields 'Author','Editor','Created','Modified' -Connection $script:Main).FieldValues
    $okAuthor = ([int]$chk['Author'].LookupId -eq $authorId)
    $okEditor = ([int]$chk['Editor'].LookupId -eq $EditorId)
    $okCre = ([Math]::Abs(((ConvertTo-Utc $chk['Created']) - $createdUtc).TotalSeconds) -le 2)
    $okMod = ([Math]::Abs(((ConvertTo-Utc $chk['Modified']) - $ModifiedUtc).TotalSeconds) -le 2)
    if (-not ($okAuthor -and $okEditor -and $okCre -and $okMod)) {
        throw ("Archive item #{0} created but Created/Modified stamps did not stick (Author {1}, Editor {2}, Created {3}, Modified {4}). Original kept." -f $archId, $okAuthor, $okEditor, $okCre, $okMod)
    }
    return $archId
}

# ----- one request -----
function Invoke-Request {
    param($Snap, $Items)
    $id = [int]$Snap.Id
    $lock = Lock-Request $id
    if ($lock.State -ne 'Locked') { return $lock.State }
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
                $other = Find-OtherActiveRequest -Items $Items -SelfId $id -SiteKey (Get-SiteKey $site) -Principal $principal -NowUtc $now
                if ($other -gt 0) {
                    $msg = "Kept site admin on $site - still needed by request #$other"
                } else {
                    Assert-SiteUrl $site
                    $res = Set-SiteAdmin -Url $site -Principal $principal -IsAdmin $false
                    $msg = "Removed $principal as site collection admin on $site ($res)"
                }
            }
            $vals = New-ResultValues -Fv $fv -Message $msg -KeepLock
            $vals['RequestStatus']     = 'Access Revoked'
            $vals['AccessRevokedDate'] = [datetime]::Now
            Save-Request -Id $id -Values $vals
            $label = 'REVOKED   site admin ->'
            if (-not $principal) { $label = 'CANCELLED (never granted) ->' }
            Write-Line ('#{0} {1} {2}' -f $id, $label, $site) 'Yellow'
            $action = 'Archive'
        }

        if ($action -eq 'Archive') {
            $final = Get-PnPListItem -List $script:ReqList -Id $id -Connection $script:Main
            $archId = Copy-ToArchive -Item $final
            Remove-PnPListItem -List $script:ReqList -Identity $id -Recycle -Force -Connection $script:Main | Out-Null
            Write-Line ('#{0} ARCHIVED -> archive #{1}' -f $id, $archId) 'Yellow'
            return 'Done'
        }

        # Nothing to do any more (another server finished it, or a user changed it): release.
        Save-Request -Id $id -Values @{ LockOwner = $null; LockExpires = $null }
        return 'Skipped'
    }
    catch {
        $err = $_.Exception.Message
        $attempts = 0
        try { $attempts = [int](Get-FieldValue $fv 'AttemptCount') } catch { $attempts = 0 }
        $attempts++
        $mins = [Math]::Min([Math]::Pow(2, $attempts), $MaxBackoffMinutes)
        $vals = New-ResultValues -Fv $fv -Message ('ERROR ({0}): {1}' -f $action, $err)
        $vals['AttemptCount'] = $attempts
        $vals['NextAttempt']  = [datetime]::Now.AddMinutes($mins)
        try { Save-Request -Id $id -Values $vals }
        catch { Write-Line ('#{0} could not write the error back: {1}' -f $id, $_.Exception.Message) 'Red' }
        Write-Line ('#{0} FAILED {1} (try {2}, retry in {3} min): {4}' -f $id, $action, $attempts, $mins, $err) 'Red'
        return 'Failed'
    }
}

function Invoke-Cycle {
    $now = [datetime]::UtcNow
    try { Update-TrackerCache $TrackerRefreshMinutes }
    catch { Write-Line ('Tracker refresh failed (using the previous copy): {0}' -f $_.Exception.Message) 'Red' }
    $items = @(Get-PnPListItem -List $script:ReqList -PageSize $PageSize -Fields $script:SnapFields -Connection $script:Main)
    $stat = @{ Open = $items.Count; Due = 0; Done = 0; Failed = 0; Busy = 0; Scheduled = 0; Waiting = 0 }
    $due = [System.Collections.Generic.List[object]]::new()
    foreach ($it in $items) {
        $fv = $it.FieldValues
        $a = Get-RequestAction -Fv $fv -NowUtc $now
        if ($a -eq 'Scheduled') { $stat['Scheduled'] += 1; continue }
        if (-not $a) { continue }
        $next = ConvertTo-Utc (Get-FieldValue $fv 'NextAttempt')
        if ($null -ne $next -and $next -gt $now) {
            # edited by someone other than this script's account since it failed (e.g. a corrected Site URL): retry now
            $ed = Get-FieldValue $fv 'Editor'
            $edId = 0
            if ($null -ne $ed) { $edId = [int]$ed.LookupId }
            if ($edId -eq $script:MyUserId) { $stat['Waiting'] += 1; continue }
        }
        $owner = [string](Get-FieldValue $fv 'LockOwner')
        $lockExp = ConvertTo-Utc (Get-FieldValue $fv 'LockExpires')
        if ($owner -and $owner -ne $RunId -and $null -ne $lockExp -and $lockExp -gt $now) { $stat['Busy'] += 1; continue }
        $due.Add($it)
    }
    $stat['Due'] = $due.Count
    # random order so several servers spread out instead of fighting over the same item
    $order = $due.ToArray()
    if ($order.Count -gt 1) { $order = @($order | Get-Random -Count $order.Count) }
    for ($i = 0; $i -lt $order.Count; $i++) {
        $r = Invoke-Request -Snap $order[$i] -Items $items
        if ($r -eq 'Done')   { $stat['Done'] += 1 }
        if ($r -eq 'Failed') { $stat['Failed'] += 1 }
        if ($r -eq 'Busy')   { $stat['Busy'] += 1 }
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
Write-Host 'Access request processor' -ForegroundColor Cyan
Write-Host ("  Site   : {0}" -f $SiteUrl)
Write-Host ("  Group  : {0}" -f $GroupClaim)
Write-Host ("  Every  : {0}s   Run id: {1}   Ctrl+C to stop" -f $PollSeconds, $RunId)
Write-Host ''

Connect-Main
[void](Get-SpoToken)
Write-Line ('Connected. {0} archive column(s) will be copied.' -f $script:CopyFields.Count) 'Cyan'

$cycle = 0
$quiet = 0
while ($true) {
    $cycle++
    try {
        $s = Invoke-Cycle
        $summary = '{0} open | {1} due | {2} done | {3} failed | {4} scheduled | {5} waiting retry | {6} on another server' -f `
            $s['Open'], $s['Due'], $s['Done'], $s['Failed'], $s['Scheduled'], $s['Waiting'], $s['Busy']
        if (($s['Done'] + $s['Failed']) -gt 0) {
            Write-Line $summary 'Cyan'
            $quiet = 0
        } else {
            $quiet++
            if ($cycle -eq 1 -or $quiet -ge $HeartbeatEvery) { Write-Line ('idle: ' + $summary) 'DarkGray'; $quiet = 0 }
        }
    }
    catch {
        Write-Line ('Cycle failed: {0}' -f $_.Exception.Message) 'Red'
        try { Connect-Main; Write-Line 'Reconnected.' 'Cyan' }
        catch { Write-Line ('Reconnect failed: {0}' -f $_.Exception.Message) 'Red' }
    }
    Start-Sleep -Seconds ($PollSeconds + (Get-Random -Minimum 0 -Maximum 4))
}
