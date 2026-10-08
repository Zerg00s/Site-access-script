$ErrorActionPreference = 'Stop'

# ---------- Settings ----------
$SiteUrl           = 'https://gocleverpointcom.sharepoint.com/sites/SharePointServices'
$ClientId          = 'e391b4e0-0151-4aa2-8ce7-dccf4b3921fa'
$RequestsListUrl   = 'Lists/AccessRequests'
$RequestsListTitle = 'Site Access Requests'      # not 'Access Requests': SharePoint has a hidden built-in list with that title
$ArchiveListUrl    = 'Lists/AccessRequestsArchive'
$ArchiveListTitle  = 'Site Access Requests Archive'

# ---------- Script ----------
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host 'PnP.PowerShell is not installed. Run: Install-Module PnP.PowerShell -Scope CurrentUser' -ForegroundColor Red
    exit 1
}

function Get-ListSafe([string]$Identity) {
    $l = $null
    try { $l = Get-PnPList -Identity $Identity -ErrorAction SilentlyContinue } catch { $l = $null }
    if ($null -eq $l -and $Identity.Contains('/')) {
        # URL lookup is not reliable in every PnP version: compare each list's real URL instead
        $want = '/' + $Identity.Trim('/').ToLowerInvariant()
        foreach ($x in @(Get-PnPList)) {
            $rel = Get-PnPProperty -ClientObject $x.RootFolder -Property ServerRelativeUrl
            if ($rel.ToLowerInvariant().EndsWith($want)) { $l = $x; break }
        }
    }
    return $l
}

function Get-OrCreateList([string]$Url, [string]$Title) {
    $l = Get-ListSafe $Url
    if ($null -eq $l) {
        $clash = Get-ListSafe $Title
        if ($null -ne $clash) {
            $clashUrl = Get-PnPProperty -ClientObject $clash.RootFolder -Property ServerRelativeUrl
            throw ("A list titled '{0}' already exists at {1} ({2} items), not at '{3}'. Set the list URL variable to that list, or change the title variable." -f $Title, $clashUrl, $clash.ItemCount, $Url)
        }
        New-PnPList -Title $Title -Url $Url -Template GenericList | Out-Null
        $l = Get-ListSafe $Url
        if ($null -eq $l) { throw "List '$Url' was not found after New-PnPList." }
        Write-Host "List created: $Title" -ForegroundColor Green
    } elseif ($l.Title -ne $Title) {
        $oldTitle = $l.Title
        Set-PnPList -Identity $l -Title $Title | Out-Null
        Write-Host "List renamed: $oldTitle -> $Title"
    } else {
        Write-Host "List exists: $($l.Title)"
    }
    return $l
}

function Add-FieldIfMissing {
    param($List, [string]$InternalName, [string]$Xml)
    $f = $null
    try { $f = Get-PnPField -List $List -Identity $InternalName -ErrorAction SilentlyContinue } catch { $f = $null }
    if ($null -ne $f) {
        $def = ([xml]$Xml.Replace('#ID#', '{00000000-0000-0000-0000-000000000000}')).Field
        if ($f.TypeAsString -ne $def.Type) {
            throw ("Column '{0}' on '{1}' is {2} but should be {3}. Delete the column (its data goes with it) and run again." -f $InternalName, $List.Title, $f.TypeAsString, $def.Type)
        }
        $want = $def.DisplayName
        if ($want -and $f.Title -ne $want) {
            Set-PnPField -List $List -Identity $InternalName -Values @{ Title = $want } | Out-Null
            Write-Host "  [name] $InternalName -> $want"
        } else { Write-Host "  [skip] $InternalName" }
        return
    }
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

function Set-FormOrder {
    param($List, [string[]]$Order)
    $ctx = Get-PnPContext
    foreach ($ct in @(Get-PnPContentType -List $List)) {
        if ($ct.Name -ne 'Item') { continue }
        $ctx.Load($ct.FieldLinks)
        $ctx.ExecuteQuery()
        $present = @($ct.FieldLinks | ForEach-Object { $_.Name })
        [string[]]$names = @($Order | Where-Object { $present -contains $_ })
        $ct.FieldLinks.Reorder($names)
        $ct.Update($false)
        $ctx.ExecuteQuery()
        Write-Host "  [upd ] form order"
    }
}

Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId


# ----- column definitions (Archive = also created on the archive list) -----
$statusChoices = ''
foreach ($c in @('Pending Access Grant','Access Granted','Pending Access Removal','Access Revoked')) { $statusChoices += "<CHOICE>$c</CHOICE>" }
$noForms = "ShowInNewForm='FALSE' ShowInEditForm='FALSE'"

$defs = [System.Collections.Generic.List[object]]::new()
function Add-Def([string]$Name, [bool]$Archive, [string]$Xml) {
    $defs.Add([pscustomobject]@{ Name = $Name; Archive = $Archive; Xml = $Xml })
}
Add-Def 'SiteUrl'               $true  "<Field Type='Text' Name='SiteUrl' StaticName='SiteUrl' DisplayName='Site URL' Description='Paste the site URL as listed in the Teams Attestation Tracker.' ID='#ID#' Required='TRUE' Indexed='TRUE' MaxLength='255' />"
Add-Def 'RequestStatus'         $true  "<Field Type='Choice' Name='RequestStatus' StaticName='RequestStatus' DisplayName='Current Status' ID='#ID#' Format='Dropdown' Required='TRUE' Indexed='TRUE'><Default>Pending Access Grant</Default><CHOICES>$statusChoices</CHOICES></Field>"
Add-Def 'Requestor'             $true  "<Field Type='User' Name='Requestor' StaticName='Requestor' DisplayName='Requestor' ID='#ID#' UserSelectionMode='PeopleOnly' Required='TRUE' />"
Add-Def 'BusinessJustification' $true  "<Field Type='Note' Name='BusinessJustification' StaticName='BusinessJustification' DisplayName='Business Justification' ID='#ID#' NumLines='4' RichText='FALSE' />"
Add-Def 'AccessStartDate'       $true  "<Field Type='DateTime' Name='AccessStartDate' StaticName='AccessStartDate' DisplayName='Date Access To Be Added' ID='#ID#' Format='DateTime' Required='TRUE'><Default>[today]</Default></Field>"
Add-Def 'AccessEndDate'         $true  "<Field Type='DateTime' Name='AccessEndDate' StaticName='AccessEndDate' DisplayName='Date Access To Be Removed' Description='Leave blank to keep access. Set it (or set status to Pending Access Removal) to remove access.' ID='#ID#' Format='DateTime' />"
Add-Def 'RevokeAccessImmediately' $false "<Field Type='Boolean' Name='RevokeAccessImmediately' StaticName='RevokeAccessImmediately' DisplayName='Revoke Access Immediately' Description='Yes = remove access now, ignoring the removal date.' ID='#ID#'><Default>0</Default></Field>"
Add-Def 'AccessGrantedDate'     $true  "<Field Type='DateTime' Name='AccessGrantedDate' StaticName='AccessGrantedDate' DisplayName='Access Added On (by Script)' ID='#ID#' Format='DateTime' $noForms />"
Add-Def 'AccessRevokedDate'     $true  "<Field Type='DateTime' Name='AccessRevokedDate' StaticName='AccessRevokedDate' DisplayName='Access Removed On (by Script)' ID='#ID#' Format='DateTime' $noForms />"
Add-Def 'GrantedPrincipal'      $true  "<Field Type='Text' Name='GrantedPrincipal' StaticName='GrantedPrincipal' DisplayName='Site Admin Group Added' ID='#ID#' MaxLength='255' $noForms />"
Add-Def 'LastResult'            $true  "<Field Type='Text' Name='LastResult' StaticName='LastResult' DisplayName='Last Result' ID='#ID#' MaxLength='255' $noForms />"
Add-Def 'LastProcessed'         $true  "<Field Type='DateTime' Name='LastProcessed' StaticName='LastProcessed' DisplayName='Last Processed' ID='#ID#' Format='DateTime' $noForms />"
Add-Def 'ProcessedBy'           $true  "<Field Type='Text' Name='ProcessedBy' StaticName='ProcessedBy' DisplayName='Processed By (Server)' ID='#ID#' MaxLength='100' $noForms />"
Add-Def 'ProcessingLog'         $true  "<Field Type='Note' Name='ProcessingLog' StaticName='ProcessingLog' DisplayName='Processing Log' ID='#ID#' NumLines='8' RichText='FALSE' $noForms />"
Add-Def 'AttemptCount'          $false "<Field Type='Number' Name='AttemptCount' StaticName='AttemptCount' DisplayName='Failed Attempts' ID='#ID#' Decimals='0' Min='0' $noForms />"
Add-Def 'NextAttempt'           $false "<Field Type='DateTime' Name='NextAttempt' StaticName='NextAttempt' DisplayName='Next Retry' ID='#ID#' Format='DateTime' $noForms />"
Add-Def 'LockOwner'             $false "<Field Type='Text' Name='LockOwner' StaticName='LockOwner' DisplayName='Lock Owner' ID='#ID#' MaxLength='255' $noForms />"
Add-Def 'LockExpires'           $false "<Field Type='DateTime' Name='LockExpires' StaticName='LockExpires' DisplayName='Lock Expires' ID='#ID#' Format='DateTime' $noForms />"

# ----- Access Requests -----
$req = Get-OrCreateList $RequestsListUrl $RequestsListTitle
Set-PnPList -Identity $req -EnableVersioning $true -MajorVersions 100 -EnableAttachments $false | Out-Null
Set-PnPField -List $req -Identity 'Title' -Values @{ Title = 'Request Title'; Required = $false } | Out-Null
Write-Host 'Columns:' -ForegroundColor Cyan
foreach ($d in $defs) { Add-FieldIfMissing $req $d.Name $d.Xml }
Set-PnPField -List $req -Identity 'AccessStartDate' -Values @{ Required = $true; DefaultValue = '[today]'; Description = '' } | Out-Null
$formOrder = @('Title','SiteUrl','RequestStatus','Requestor','BusinessJustification','AccessStartDate','AccessEndDate','RevokeAccessImmediately','AccessGrantedDate','AccessRevokedDate')
Set-FormOrder -List $req -Order $formOrder

Write-Host 'Views:' -ForegroundColor Cyan
$main  = @('ID','SiteUrl','RequestStatus','Requestor','AccessStartDate','AccessGrantedDate','AccessEndDate','RevokeAccessImmediately','AccessRevokedDate','LastResult','Modified','Editor')
$admin = @('ID','SiteUrl','RequestStatus','LastResult','LastProcessed','ProcessedBy','AttemptCount','NextAttempt','LockOwner','LockExpires','GrantedPrincipal')
Set-ViewSafe -List $req -Title 'All Items'            -Fields $main -Query "<OrderBy><FieldRef Name='ID' Ascending='FALSE' /></OrderBy>"
Set-ViewSafe -List $req -Title 'Pending Grant'        -Fields $main -Query "<Where><Eq><FieldRef Name='RequestStatus' /><Value Type='Choice'>Pending Access Grant</Value></Eq></Where><OrderBy><FieldRef Name='AccessStartDate' /></OrderBy>"
Set-ViewSafe -List $req -Title 'Active Access'        -Fields $main -Query "<Where><Eq><FieldRef Name='RequestStatus' /><Value Type='Choice'>Access Granted</Value></Eq></Where><OrderBy><FieldRef Name='AccessEndDate' /></OrderBy>"
Set-ViewSafe -List $req -Title 'Removal Due'          -Fields $main -Query "<Where><Or><And><Eq><FieldRef Name='RequestStatus' /><Value Type='Choice'>Access Granted</Value></Eq><Leq><FieldRef Name='AccessEndDate' /><Value Type='DateTime'><Today /></Value></Leq></And><Or><Eq><FieldRef Name='RequestStatus' /><Value Type='Choice'>Pending Access Removal</Value></Eq><Eq><FieldRef Name='RevokeAccessImmediately' /><Value Type='Boolean'>1</Value></Eq></Or></Or></Where>"
Set-ViewSafe -List $req -Title 'My Requests'          -Fields $main -Query "<Where><Or><Eq><FieldRef Name='Requestor' /><Value Type='Integer'><UserID /></Value></Eq><Eq><FieldRef Name='Author' /><Value Type='Integer'><UserID /></Value></Eq></Or></Where><OrderBy><FieldRef Name='ID' Ascending='FALSE' /></OrderBy>"
Set-ViewSafe -List $req -Title 'Needs Attention'      -Fields $admin -Query "<Where><Gt><FieldRef Name='AttemptCount' /><Value Type='Number'>0</Value></Gt></Where>"
Set-ViewSafe -List $req -Title 'Processing Details'   -Fields $admin -Query "<OrderBy><FieldRef Name='LastProcessed' Ascending='FALSE' /></OrderBy>"

# ----- Access Requests Archive -----
$arc = Get-OrCreateList $ArchiveListUrl $ArchiveListTitle
Set-PnPList -Identity $arc -EnableVersioning $false -EnableAttachments $false | Out-Null
Set-PnPField -List $arc -Identity 'Title' -Values @{ Title = 'Request Title'; Required = $false } | Out-Null
Write-Host 'Columns:' -ForegroundColor Cyan
foreach ($d in $defs) {
    if (-not $d.Archive) { continue }
    $axml = $d.Xml.Replace("Required='TRUE'", "Required='FALSE'")
    Add-FieldIfMissing $arc $d.Name $axml
}
Add-FieldIfMissing $arc 'OriginalItemId' "<Field Type='Number' Name='OriginalItemId' StaticName='OriginalItemId' DisplayName='Original Request ID' ID='#ID#' Decimals='0' Indexed='TRUE' />"
Add-FieldIfMissing $arc 'SourceKey'      "<Field Type='Text' Name='SourceKey' StaticName='SourceKey' DisplayName='Source Key' ID='#ID#' MaxLength='100' Indexed='TRUE' />"
Add-FieldIfMissing $arc 'ArchivedDate'   "<Field Type='DateTime' Name='ArchivedDate' StaticName='ArchivedDate' DisplayName='Archived Date' ID='#ID#' Format='DateTime' />"
Add-FieldIfMissing $arc 'ArchivedBy'     "<Field Type='Text' Name='ArchivedBy' StaticName='ArchivedBy' DisplayName='Archived By (Server)' ID='#ID#' MaxLength='100' />"
Set-FormOrder -List $arc -Order $formOrder

Write-Host 'Views:' -ForegroundColor Cyan
$arcFields = @('OriginalItemId','SiteUrl','RequestStatus','Requestor','AccessStartDate','AccessGrantedDate','AccessEndDate','AccessRevokedDate','Author','Created','Editor','Modified','ArchivedDate')
Set-ViewSafe -List $arc -Title 'All Items' -Fields $arcFields -Query "<OrderBy><FieldRef Name='ArchivedDate' Ascending='FALSE' /></OrderBy>"
Set-ViewSafe -List $arc -Title 'By Site'   -Fields $arcFields -Query "<GroupBy Collapse='TRUE'><FieldRef Name='SiteUrl' /></GroupBy><OrderBy><FieldRef Name='ArchivedDate' Ascending='FALSE' /></OrderBy>"

Write-Host "Done." -ForegroundColor Green
Write-Host "  $SiteUrl/$RequestsListUrl"
Write-Host "  $SiteUrl/$ArchiveListUrl"
exit 0
