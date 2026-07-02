# ============================================================================

#  GENERATED FILE - DO NOT EDIT DIRECTLY.

#  Built from src/ by build/Build.ps1. Edit the sources there and rebuild.

# ============================================================================

#requires -Version 5.1
<#
.SYNOPSIS
    Turns a Viedoc eTMF-EMS archive (.zip or extracted folder) into a human-friendly
    rendering: a clean eTMF folder tree that separates Final / In Progress / Superseded
    documents, plus a self-contained HTML mini-viewer.

.DESCRIPTION
    Reads the eTMF Exchange Mechanism Standard (eTMF-EMS v1.0) manifest (*_exchange.xml)
    and the referenced files, then:
      * parses every OBJECT / FILE / AUDITRECORD / SIGNATURE / METADATA into an object model,
      * classifies each document (per object) by lifecycle (audit events) AND version state,
      * copies + renames files into a readable tree, verifying MD5 on every copy,
      * writes inventory.csv, a run log, and a self-contained index.html viewer.

    The eTMF tree is the backbone: Zone > Section > Artifact, with the status split made
    *inside* each artifact:  Artifact > <status bucket> > <Country/Site> > files.
    Multi-milestone exports (more than one manifest) get one subfolder per event plus a
    landing page. The original archive is never modified.

.PARAMETER ArchivePath
    Path to the eTMF archive .zip, OR a folder that already contains the extracted archive
    (i.e. the folder holding the *_exchange.xml manifest next to the TransferID folder).

.PARAMETER OutputPath
    Destination folder for the humanized output. Defaults to "<input directory>\<StudyName>",
    where the folder name is built from $DefaultOutputNameTemplate near the top of the script.

.PARAMETER KeepOriginalNames
    Copy files with their original filenames instead of the readable rename convention.

.PARAMETER SkipIntegrityCheck
    Skip MD5 integrity verification of referenced files (verification is on by default).

.PARAMETER Force
    Overwrite the output folder if it already exists.

.PARAMETER Open
    Open the generated index.html when finished.

.PARAMETER DryRun
    Parse, classify and print the summary only — write nothing.

.EXAMPLE
    .\Convert-EtmfArchive.ps1 -ArchivePath ".\2022 - Demo Study_eTMFArchive_20260626085449.zip"

.EXAMPLE
    .\Convert-EtmfArchive.ps1 -ArchivePath ".\archive" -Force -Open

.NOTES
    Runs on Windows PowerShell 5.1 (built into Windows 10/11) or PowerShell 7 (pwsh). If the
    script is blocked as "not digitally signed", add -ExecutionPolicy Bypass (this one run only):
        powershell -ExecutionPolicy Bypass -File .\Convert-EtmfArchive.ps1 -ArchivePath "...zip"

    Spec: eTMF Exchange Mechanism Standard v1.0 (public domain, TMF Reference Model).
    Integrity values are base64-encoded MD5 (CHKSUMSTD omitted by Viedoc; verified empirically).
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Position = 0)]   # optional: with no value, the script prints friendly usage and exits
    [string]$ArchivePath,

    [Parameter()]
    [string]$OutputPath,

    [Parameter()]
    [switch]$KeepOriginalNames,

    [Parameter()]
    [switch]$SkipIntegrityCheck,

    [Parameter()]
    [switch]$Force,

    [Parameter()]
    [switch]$Open,

    [Parameter()]
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# --- Windows PowerShell 5.1 compatibility -----------------------------------------------
# 5.1 predates the $IsWindows/$IsMacOS/$IsLinux automatic variables. Define them (as Windows,
# since Windows PowerShell only runs there) so the strict-mode references in the -Open path
# don't throw. On PowerShell 7 these already exist as read-only automatics, so we leave them.
if (-not (Test-Path Variable:IsWindows)) {
    Set-Variable IsWindows -Value $true
    Set-Variable IsMacOS   -Value $false
    Set-Variable IsLinux   -Value $false
}

# ----------------------------------------------------------------------------------------
#  Constants
# ----------------------------------------------------------------------------------------
$script:EmsNamespace = 'https://tmfrefmodel.com/ems'
$script:LogLines     = [System.Collections.Generic.List[string]]::new()

# Default output folder name when -OutputPath is not supplied. {0} = study name.
# Edit this to set your house style, e.g. '{0}', '{0} - eTMF', or 'eTMF - {0}'.
$script:DefaultOutputNameTemplate = '{0}'


# ----------------------------------------------------------------------------------------
#  Logging
# ----------------------------------------------------------------------------------------
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'OK', 'STEP')]
        [string]$Level = 'INFO'
    )
    $stamp = (Get-Date).ToString('HH:mm:ss')
    $script:LogLines.Add("$stamp [$Level] $Message")
    $color = switch ($Level) {
        'OK'    { 'Green' }
        'WARN'  { 'Yellow' }
        'ERROR' { 'Red' }
        'STEP'  { 'Cyan' }
        default { 'Gray' }
    }
    $prefix = switch ($Level) {
        'OK'   { '  [OK]  ' }
        'WARN' { '  [!]   ' }
        'ERROR'{ '  [X]   ' }
        'STEP' { '==> ' }
        default{ '       ' }
    }
    Write-Host "$prefix$Message" -ForegroundColor $color
}


# ----------------------------------------------------------------------------------------
#  Small helpers
# ----------------------------------------------------------------------------------------

# EMS dates are "DD-MON-YYYY" (e.g. 12-Nov-2022). Returns [datetime] or $null.
function ConvertFrom-EmsDate {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $formats = @('dd-MMM-yyyy', 'd-MMM-yyyy')
    foreach ($f in $formats) {
        $dt = [datetime]::MinValue
        if ([datetime]::TryParseExact($Value.Trim(), $f,
                [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::None, [ref]$dt)) {
            return $dt
        }
    }
    return $null
}

# Get the text of a child element (by local name) under an XmlElement; '' if absent.
function Get-NodeText {
    param([System.Xml.XmlElement]$Node, [string]$Name)
    if ($null -eq $Node) { return '' }
    $child = $Node.SelectSingleNode("*[local-name()='$Name']")
    if ($null -eq $child) { return '' }
    return $child.InnerText.Trim()
}

# Map an EMS CONTENTURL to an on-disk path. CONTENTURL is "<zipname>.zip/<TransferID>/...";
# strip the leading zip-name segment and resolve the remainder beside the manifest.
# SECURITY: the manifest is untrusted, so the resolved path is canonicalized and must stay
# CONTAINED under $ManifestDir. A traversal ("../") or rooted CONTENTURL returns $null
# (the referenced file is then recorded as a gap, never read or copied).
function Resolve-ContentPath {
    param([string]$ContentUrl, [string]$ManifestDir)
    $segments = $ContentUrl -split '/'
    if ($segments.Count -lt 2) { return $null }
    $rel = ($segments | Select-Object -Skip 1) -join [IO.Path]::DirectorySeparatorChar
    $base = [IO.Path]::GetFullPath($ManifestDir)
    $full = [IO.Path]::GetFullPath((Join-Path $base $rel))
    $sep  = [IO.Path]::DirectorySeparatorChar
    $baseWithSep = if ($base.EndsWith($sep)) { $base } else { $base + $sep }
    if (-not $full.StartsWith($baseWithSep, [System.StringComparison]::Ordinal)) { return $null }  # escapes -> reject
    return $full
}

# Base64-encoded MD5 of a file (matches the EMS INTEGRITY encoding).
function Get-FileMd5Base64 {
    param([string]$Path)
    $md5 = [System.Security.Cryptography.MD5]::Create()
    try {
        $stream = [System.IO.File]::OpenRead($Path)
        try   { return [Convert]::ToBase64String($md5.ComputeHash($stream)) }
        finally { $stream.Dispose() }
    }
    finally { $md5.Dispose() }
}

# Make a string safe to use as a Windows + Linux file/folder name.
function ConvertTo-SafeName {
    param([string]$Name)
    $clean = ($Name -replace '[<>:"/\\|?*\x00-\x1F]', '_')
    $clean = $clean.Trim().TrimEnd('.', ' ')   # Windows dislikes trailing dot/space
    if ([string]::IsNullOrWhiteSpace($clean)) { $clean = 'file' }
    return $clean
}

# "01.01.trial oversight" -> "01.01 Trial Oversight"  (numeric prefix + Title Case name)
function Get-PrettyTmfSegment {
    param([string]$Segment)
    $ti = (Get-Culture).TextInfo
    $parts = $Segment -split '\.'
    $num = @(); $name = @(); $inNum = $true
    foreach ($p in $parts) {
        if ($inNum -and $p -match '^\d+$') { $num += $p } else { $inNum = $false; $name += $p }
    }
    $numStr  = ($num -join '.')
    $nameStr = ($name -join '.').Trim()
    if ($nameStr) { $nameStr = $ti.ToTitleCase($nameStr) }
    if ($numStr -and $nameStr) { return "$numStr $nameStr" }
    if ($numStr)  { return $numStr }
    return $ti.ToTitleCase($Segment)
}

# Country folder ("de" -> "DE"), site ("university medical center freiburg" -> Title Case), "signed" -> "Signed".
function Get-PrettyScope {
    param([string]$Segment)
    if ($Segment -ieq 'signed') { return 'Signed' }
    if ($Segment.Length -le 3)  { return $Segment.ToUpper() }
    return (Get-Culture).TextInfo.ToTitleCase($Segment)
}

# Split a CONTENTURL into TMF path parts using the TransferID as anchor.
# Bounds-safe: a short/malformed CONTENTURL yields empty parts (never throws under StrictMode),
# so the referenced document degrades to a recorded gap rather than aborting the whole run.
function Get-ContentPathParts {
    param([string]$ContentUrl, [string]$TransferId)
    $segs = $ContentUrl -split '/'
    $idx = [array]::IndexOf($segs, $TransferId)
    if ($idx -lt 0) { $idx = 1 }                 # fallback: <zip>/<transfer>/...
    $artIdx = $idx + 3
    $at = { param($i) if ($i -ge 0 -and $i -lt $segs.Count) { $segs[$i] } else { '' } }
    $scope = @()
    $lastScopeIdx = $segs.Count - 2              # element before the filename
    if ($lastScopeIdx -ge ($artIdx + 1)) { $scope = @($segs[($artIdx + 1)..$lastScopeIdx]) }
    return [pscustomobject]@{
        Zone     = (& $at ($idx + 1))
        Section  = (& $at ($idx + 2))
        Artifact = (& $at $artIdx)
        Scope    = $scope
        File     = if ($segs.Count -ge 1) { $segs[-1] } else { '' }
    }
}



# Lifecycle = the furthest stage reached across an object's audit events. Matching is tolerant
# (substring, case-insensitive) so producer phrasing variants like "Document Finalized" still map.
function Get-EtmfLifecycle {
    param([string[]]$AuditEvents)
    $events = @($AuditEvents)
    if ($events | Where-Object { $_ -like '*lock*' })            { return 'Locked' }
    if ($events | Where-Object { $_ -like '*finaliz*' })         { return 'Finalized' }
    if ($events | Where-Object { $_ -like '*awaiting review*' }) { return 'In Review' }
    return 'Draft'
}

# Bucket = the folder a document lands in (two-axis: version state + lifecycle).
function Get-EtmfBucket {
    param([string]$VersionState, [string]$Lifecycle)
    if ($VersionState -in @('Superseded', 'Obsolete')) { return '03 Superseded' }
    if ($Lifecycle    -in @('Finalized', 'Locked'))    { return '01 Final' }
    return '02 In Progress'
}

# Status token shown in filename / badge: precise lifecycle, or the supersession state.
function Get-EtmfStatusToken {
    param([string]$VersionState, [string]$Lifecycle)
    if ($VersionState -in @('Superseded', 'Obsolete')) { return $VersionState }
    return $Lifecycle
}


# ----------------------------------------------------------------------------------------
#  Input resolution: accept a .zip (extract to temp) or an already-extracted folder.
# ----------------------------------------------------------------------------------------
function Resolve-EtmfSource {
    param([string]$ArchivePath)

    if (-not (Test-Path -LiteralPath $ArchivePath)) {
        throw "ArchivePath not found: $ArchivePath"
    }
    $item = Get-Item -LiteralPath $ArchivePath

    if ($item.PSIsContainer) {
        Write-Log "Using extracted archive folder: $($item.FullName)" 'INFO'
        return [pscustomobject]@{ Root = $item.FullName; TempDir = $null; IsTemp = $false }
    }

    if ($item.Extension -ieq '.zip') {
        $temp = Join-Path ([IO.Path]::GetTempPath()) ("etmf_" + [Guid]::NewGuid().ToString('N').Substring(0, 8))
        New-Item -ItemType Directory -Path $temp -Force | Out-Null
        Write-Log "Extracting archive to temp: $temp" 'STEP'
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
        [System.IO.Compression.ZipFile]::ExtractToDirectory($item.FullName, $temp)
        return [pscustomobject]@{ Root = $temp; TempDir = $temp; IsTemp = $true }
    }

    throw "ArchivePath must be a .zip file or a folder. Got: $($item.FullName)"
}

# Find the EMS manifest(s): *_exchange.xml at (or under) the archive root.
function Find-ExchangeManifests {
    param([string]$Root)
    $found = @(Get-ChildItem -LiteralPath $Root -Recurse -File -Filter '*exchange.xml' -ErrorAction SilentlyContinue)
    if ($found.Count -eq 0) {
        throw "No '*exchange.xml' manifest found under: $Root"
    }
    return $found
}


# ----------------------------------------------------------------------------------------
#  Manifest parsing -> object model
# ----------------------------------------------------------------------------------------
function Import-EtmfManifest {
    param([System.IO.FileInfo]$ManifestFile)

    $xml = New-Object System.Xml.XmlDocument
    # Untrusted manifest: disable DTD / external-entity resolution so a malicious manifest can't
    # mount an XXE attack. .NET Core (PowerShell 7) already defaults to this; .NET Framework
    # (Windows PowerShell 5.1) does NOT, so setting it explicitly is required for a safe 5.1 run.
    $xml.XmlResolver = $null
    $xml.Load($ManifestFile.FullName)
    $batchEl = $xml.DocumentElement
    if ($batchEl.LocalName -ne 'BATCH') {
        throw "Unexpected root element '$($batchEl.LocalName)' (expected BATCH) in $($ManifestFile.Name)"
    }
    $manifestDir = $ManifestFile.DirectoryName

    $batch = [pscustomobject]@{
        StudySystemId   = $batchEl.GetAttribute('STUDYSYSTEMID')
        StudyId         = $batchEl.GetAttribute('STUDYID')
        EventId         = $batchEl.GetAttribute('EVENTID')
        TransferId      = $batchEl.GetAttribute('TRANSFERID')
        TransferSource  = $batchEl.GetAttribute('TRANSFERSOURCEID')
        SpecificationId = $batchEl.GetAttribute('SPECIFICATIONID')
        TmfRmVersion    = $batchEl.GetAttribute('TMFRMVERSION')
        ManifestFile    = $ManifestFile.Name
        ManifestDir     = $manifestDir
    }

    $documents = [System.Collections.Generic.List[object]]::new()
    $objectNodes = @($batchEl.SelectNodes("*[local-name()='OBJECT']"))

    foreach ($obj in $objectNodes) {
        # ---- object-level (artifact instance) fields ----
        $artifactNumber = Get-NodeText $obj 'ARTIFACTNUMBER'
        $versionState   = Get-NodeText $obj 'OBJECTVERSIONSTATE'
        $country        = Get-NodeText $obj 'COUNTRYID'
        $org            = Get-NodeText $obj 'ORGANIZATIONNAME'
        # an OBJECT may relate to multiple people; capture them all
        $person         = (@($obj.SelectNodes("*[local-name()='PERSONNAME']") |
                             ForEach-Object { $_.InnerText.Trim() } | Where-Object { $_ }) -join '; ')
        # title falls back to the sub-artifact, then a placeholder, so folders/names never go bare
        $artifactTitle  = Get-NodeText $obj 'OBJECTTITLE'
        if (-not $artifactTitle) { $artifactTitle = Get-NodeText $obj 'SUBARTIFACT' }
        if (-not $artifactTitle) { $artifactTitle = 'Untitled Document' }
        # additional EMS OBJECT fields (surfaced in inventory + viewer)
        $restricted   = Get-NodeText $obj 'RESTRICTED'
        $language     = Get-NodeText $obj 'OBJECTLANGUAGE'
        $translation  = Get-NodeText $obj 'TRANSLATION'
        $expiryDate   = Get-NodeText $obj 'OBJECTEXPIRYDATE'

        $metadata = @{}
        foreach ($m in $obj.SelectNodes("*[local-name()='METADATA']")) {
            $metadata[$m.GetAttribute('NAME')] = $m.InnerText.Trim()
        }

        # ---- classify ONCE per object (an OBJECT = one artifact record / Viedoc document).
        # Multiple FILE nodes are attachments of that single record; the lifecycle action
        # (e.g. "Finalized") is logged on just one file, so we aggregate events across them. ----
        $fileNodes = @($obj.SelectNodes("*[local-name()='FILE']"))
        $objEventHeads = foreach ($fn in $fileNodes) {
            foreach ($a in $fn.SelectNodes("*[local-name()='AUDITRECORD']")) {
                ((Get-NodeText $a 'AUDITEVENT') -replace "`r?`n.*", '').Trim()
            }
        }
        $lifecycle = Get-EtmfLifecycle @($objEventHeads)
        $bucket    = Get-EtmfBucket $versionState $lifecycle
        $status    = Get-EtmfStatusToken $versionState $lifecycle

        # ---- files (one humanized copy each), de-duplicated by CONTENTURL ----
        $seen = [System.Collections.Generic.HashSet[string]]::new()
        foreach ($file in $fileNodes) {
            $contentUrl = Get-NodeText $file 'CONTENTURL'
            if (-not $seen.Add($contentUrl)) { continue }   # skip duplicate refs (e.g. the doubled signed file)

            $auditTrail = foreach ($a in $file.SelectNodes("*[local-name()='AUDITRECORD']")) {
                [pscustomobject]@{
                    AuditId   = Get-NodeText $a 'AUDITID'
                    Timestamp = Get-NodeText $a 'DATETIMESTAMP'
                    User      = Get-NodeText $a 'USERREF'
                    EntryType = Get-NodeText $a 'AUDITENTRYTYPE'
                    Event     = Get-NodeText $a 'AUDITEVENT'
                }
            }
            $auditTrail = @($auditTrail)

            $signatures = foreach ($s in $file.SelectNodes("*[local-name()='SIGNATURE']")) {
                [pscustomobject]@{
                    Methodology = Get-NodeText $s 'SIGNATUREMETHODOLOGY'
                    UserId      = Get-NodeText $s 'USERID'
                    SignerName  = Get-NodeText $s 'SIGNATURENAME'
                    DateTime    = Get-NodeText $s 'SIGNATUREDATETIME'
                    Reason      = Get-NodeText $s 'SIGNATUREREASON'
                }
            }
            $signatures = @($signatures)

            $diskPath = Resolve-ContentPath $contentUrl $manifestDir
            $exists   = ($null -ne $diskPath) -and (Test-Path -LiteralPath $diskPath)

            $documents.Add([pscustomobject]@{
                ArtifactNumber = $artifactNumber
                ObjectLevel    = Get-NodeText $obj 'OBJECTLEVEL'
                Country        = $country
                SiteId         = Get-NodeText $obj 'SITEID'
                Organization   = $org
                Person         = $person
                UniqueId       = Get-NodeText $obj 'UNIQUEID'
                ArtifactTitle  = $artifactTitle
                SubArtifact    = Get-NodeText $obj 'SUBARTIFACT'
                Version        = Get-NodeText $obj 'OBJECTVERSION'
                VersionState   = $versionState
                IsCopy         = Get-NodeText $obj 'OBJECTCOPY'
                Restricted     = $restricted
                Language       = $language
                Translation    = $translation
                ExpiryDate     = $expiryDate
                ArtifactDate   = Get-NodeText $obj 'ARTIFACTDATE'
                DateDesc       = Get-NodeText $obj 'DATEDESCRIPTION'
                ObjectId       = Get-NodeText $obj 'OBJECTID'
                DocumentId     = ($metadata['DocumentId'])
                MetaVersion    = ($metadata['Version'])
                Lifecycle      = $lifecycle
                Bucket         = $bucket
                Status         = $status
                OriginalName   = Get-NodeText $file 'FILENAME'
                ContentUrl     = $contentUrl
                Integrity      = Get-NodeText $file 'INTEGRITY'
                DiskPath       = $diskPath
                FileExists     = $exists
                SizeBytes      = if ($exists) { (Get-Item -LiteralPath $diskPath).Length } else { 0 }
                IntegrityOk    = $null   # filled by integrity pass (source)
                AuditTrail     = $auditTrail
                Signatures     = $signatures
                Metadata       = $metadata
                NewName        = ''      # filled by Phase 5 (copy/rename)
                RelPath        = ''      # filled by Phase 5 (path relative to output root)
                CopyIntegrityOk = $null  # filled by Phase 5 (verify the copy)
            })
        }
    }

    return [pscustomobject]@{ Batch = $batch; Documents = $documents; ObjectCount = $objectNodes.Count }
}

# ----------------------------------------------------------------------------------------
#  Integrity verification (base64 MD5)
# ----------------------------------------------------------------------------------------
function Invoke-IntegrityCheck {
    param([System.Collections.Generic.List[object]]$Documents)
    $ok = 0; $bad = 0; $missing = 0
    foreach ($d in $Documents) {
        if (-not $d.FileExists) { $missing++; $d.IntegrityOk = $false; continue }
        $actual = Get-FileMd5Base64 $d.DiskPath
        $match  = ($actual -ceq $d.Integrity)   # base64 is case-sensitive
        $d.IntegrityOk = $match
        if ($match) { $ok++ } else { $bad++; Write-Log "Integrity MISMATCH: $($d.OriginalName)" 'WARN' }
    }
    return [pscustomobject]@{ Ok = $ok; Bad = $bad; Missing = $missing }
}


# Human-friendly destination filename (or original name if -KeepOriginalNames).
# The folder already encodes Zone/Section/Artifact, so the file name only needs the document's
# own name + version + status -- this keeps the full path within the Windows MAX_PATH (260) limit.
# -MaxLength (when > 0) caps the whole file name, truncating the document-name part to fit.
function Get-HumanFileName {
    param($Doc, [switch]$KeepOriginal, [int]$MaxLength = 0)
    $ext = [System.IO.Path]::GetExtension($Doc.OriginalName)
    if ($KeepOriginal) { return (ConvertTo-SafeName $Doc.OriginalName) }
    $stem   = [System.IO.Path]::GetFileNameWithoutExtension($Doc.OriginalName)
    $suffix = " v{0} [{1}]" -f $Doc.Version, $Doc.Status
    if ($MaxLength -gt 0) {
        $room = $MaxLength - $suffix.Length - $ext.Length
        if ($room -lt 8) { $room = 8 }
        if ($stem.Length -gt $room) { $stem = $stem.Substring(0, $room - 1).TrimEnd() + '~' }  # ~ marks truncation
    }
    return (ConvertTo-SafeName ($stem + $suffix)) + $ext
}

# Copy every document into  <out>/Documents/<Zone>/<Section>/<Artifact>/<Bucket>/<Scope>/<file>,
# re-verify each copy, and return the inventory rows.
function Invoke-Humanize {
    param(
        $Model,
        [string]$OutputRoot,
        [switch]$KeepOriginalNames,
        [switch]$VerifyCopy
    )
    $batch    = $Model.Batch
    $treeRoot = Join-Path $OutputRoot 'Documents'
    New-Item -ItemType Directory -Path $treeRoot -Force | Out-Null

    $used      = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $inventory = [System.Collections.Generic.List[object]]::new()
    $copied = 0; $missing = 0; $copyBad = 0; $shortened = 0
    $targetMaxPath = 250   # keep generated full paths comfortably under the Windows MAX_PATH (260)

    foreach ($d in $Model.Documents) {
        $pp   = Get-ContentPathParts $d.ContentUrl $batch.TransferId
        $note = ''

        if ($d.FileExists) {
            $zone     = ConvertTo-SafeName (Get-PrettyTmfSegment $pp.Zone)
            $section  = ConvertTo-SafeName (Get-PrettyTmfSegment $pp.Section)
            $artifact = ConvertTo-SafeName ("{0} {1}" -f $d.ArtifactNumber, $d.ArtifactTitle)

            $folder = Join-Path (Join-Path (Join-Path (Join-Path $treeRoot $zone) $section) $artifact) $d.Bucket
            foreach ($s in $pp.Scope) { $folder = Join-Path $folder (ConvertTo-SafeName (Get-PrettyScope $s)) }
            New-Item -ItemType Directory -Path $folder -Force | Out-Null

            # budget the file name so the full destination path stays within the path limit
            $budget   = $targetMaxPath - ($folder.Length + 1)
            if ($budget -lt 30) { $budget = 30 }
            $fileName = Get-HumanFileName $d -KeepOriginal:$KeepOriginalNames -MaxLength $budget
            if (-not $KeepOriginalNames -and $fileName -ne (Get-HumanFileName $d)) { $shortened++ }
            $dest     = Join-Path $folder $fileName
            if ($used.Contains($dest) -or (Test-Path -LiteralPath $dest)) {    # collision guard
                $stem = [IO.Path]::GetFileNameWithoutExtension($fileName)
                $ext  = [IO.Path]::GetExtension($fileName); $n = 2
                do   { $dest = Join-Path $folder ("{0} ({1}){2}" -f $stem, $n, $ext); $n++ }
                while ($used.Contains($dest) -or (Test-Path -LiteralPath $dest))
            }
            [void]$used.Add($dest)

            Copy-Item -LiteralPath $d.DiskPath -Destination $dest -Force
            $copyOk = $null
            if ($VerifyCopy) {
                $copyOk = ((Get-FileMd5Base64 $dest) -ceq $d.Integrity)   # base64 is case-sensitive
                if (-not $copyOk) { $copyBad++; Write-Log "Copy integrity MISMATCH: $fileName" 'WARN' }
            }
            $copied++

            $d.NewName         = Split-Path $dest -Leaf
            $d.RelPath         = ($dest.Substring($OutputRoot.Length)).TrimStart('/', '\') -replace '\\', '/'
            $d.CopyIntegrityOk = $copyOk
        }
        else {
            Write-Log "Missing source file (recorded as gap): $($d.OriginalName)" 'WARN'
            $missing++; $note = 'MISSING SOURCE FILE'
        }

        $inventory.Add([pscustomobject]@{
            ArtifactNumber = $d.ArtifactNumber
            Zone           = Get-PrettyTmfSegment $pp.Zone
            Section        = Get-PrettyTmfSegment $pp.Section
            ArtifactTitle  = $d.ArtifactTitle
            TmfLevel       = $d.ObjectLevel
            Country        = $d.Country
            Organization   = $d.Organization
            Person         = $d.Person
            Language       = $d.Language
            Restricted     = $d.Restricted
            ExpiryDate     = $d.ExpiryDate
            Version        = $d.Version
            VersionState   = $d.VersionState
            Lifecycle      = $d.Lifecycle
            Bucket         = $d.Bucket
            Status         = $d.Status
            DocumentId     = $d.DocumentId
            ObjectId       = $d.ObjectId
            ArtifactDate   = $d.ArtifactDate
            OriginalName   = $d.OriginalName
            NewName        = $d.NewName
            RelativePath   = $d.RelPath
            SizeBytes      = $d.SizeBytes
            Integrity      = $d.Integrity
            SourceMd5Ok    = $d.IntegrityOk
            CopyMd5Ok      = $d.CopyIntegrityOk
            Signed         = if ($d.Signatures.Count -gt 0) { 'Yes' } else { 'No' }
            AuditRecords   = $d.AuditTrail.Count
            Note           = $note
        })
    }

    if ($shortened -gt 0) {
        Write-Log ("{0} long file name(s) shortened (with a trailing ~) to stay within the Windows 260-char path limit; originals are in inventory.csv." -f $shortened) 'INFO'
    }

    return [pscustomobject]@{
        Inventory = $inventory; TreeRoot = $treeRoot
        Copied = $copied; Missing = $missing; CopyBad = $copyBad
    }
}


function New-HtmlViewer {
    param($Model, [string]$OutputRoot, $Integrity)

    $batch = $Model.Batch

    # ISO-3166 alpha-3 -> alpha-2 so the viewer's country agrees with the alpha-2 folder tree.
    $alpha2 = @{
        DEU='DE'; SWE='SE'; USA='US'; JPN='JP'; GBR='GB'; FRA='FR'; ITA='IT'; ESP='ES'; NLD='NL'
        BEL='BE'; CHE='CH'; AUT='AT'; POL='PL'; DNK='DK'; NOR='NO'; FIN='FI'; CAN='CA'; AUS='AU'
        BRA='BR'; CHN='CN'; IND='IN'; KOR='KR'; RUS='RU'; MEX='MX'; ZAF='ZA'; IRL='IE'; PRT='PT'
        CZE='CZ'; HUN='HU'; GRC='GR'; ISR='IL'; TUR='TR'; ARG='AR'; NZL='NZ'; SGP='SG'
    }

    # ---- per-document records for the front-end ----
    $records = foreach ($d in $Model.Documents) {
        $pp    = Get-ContentPathParts $d.ContentUrl $batch.TransferId
        $scope = @($pp.Scope | ForEach-Object { Get-PrettyScope $_ })
        $iso   = ConvertFrom-EmsDate $d.ArtifactDate
        $ctry  = if ($d.Country -and $alpha2.ContainsKey($d.Country)) { $alpha2[$d.Country] } else { $d.Country }
        [pscustomobject]@{
            artNo        = $d.ArtifactNumber
            zone         = (Get-PrettyTmfSegment $pp.Zone)
            section      = (Get-PrettyTmfSegment $pp.Section)
            artifact     = ("{0} {1}" -f $d.ArtifactNumber, $d.ArtifactTitle)
            title        = $d.ArtifactTitle
            level        = $d.ObjectLevel
            country      = $ctry
            scope        = ($scope -join ' / ')
            person       = $d.Person
            version      = $d.Version
            versionState = $d.VersionState
            lifecycle    = $d.Lifecycle
            bucket       = $d.Bucket
            status       = $d.Status
            restricted   = ($d.Restricted -ieq 'Yes')
            language     = $d.Language
            translation  = $d.Translation
            expiry       = $d.ExpiryDate
            date         = $d.ArtifactDate
            dateIso      = if ($iso) { $iso.ToString('yyyy-MM-dd') } else { '' }
            dateDesc     = $d.DateDesc
            documentId   = "$($d.DocumentId)"
            objectId     = $d.ObjectId
            originalName = $d.OriginalName
            newName      = $d.NewName
            relPath      = $d.RelPath
            sizeBytes    = $d.SizeBytes
            present      = [bool]$d.FileExists
            integrity    = $d.Integrity
            sourceOk     = [bool]$d.IntegrityOk
            copyOk       = [bool]$d.CopyIntegrityOk
            signatures   = @($d.Signatures | ForEach-Object {
                                [pscustomobject]@{ methodology = $_.Methodology; signer = $_.SignerName; dateTime = $_.DateTime; reason = $_.Reason } })
            audit        = @($d.AuditTrail | ForEach-Object {
                                [pscustomobject]@{ ts = $_.Timestamp; user = $_.User; type = $_.EntryType; event = $_.Event } })
        }
    }

    $intOk    = if ($Integrity) { $Integrity.Ok } else { 0 }
    $intTotal = if ($Integrity) { $Integrity.Ok + $Integrity.Bad + $Integrity.Missing } else { 0 }

    $meta = [pscustomobject]@{
        study          = $batch.SpecificationId
        studyId        = $batch.StudyId
        event          = $batch.EventId
        transferId     = $batch.TransferId
        tmfrm          = $batch.TmfRmVersion
        source         = $batch.TransferSource
        manifest       = $batch.ManifestFile
        generated      = (Get-Date).ToString('yyyy-MM-dd HH:mm')
        total          = $Model.Documents.Count
        final          = @($Model.Documents | Where-Object { $_.Bucket -eq '01 Final' }).Count
        inProgress     = @($Model.Documents | Where-Object { $_.Bucket -eq '02 In Progress' }).Count
        superseded     = @($Model.Documents | Where-Object { $_.Bucket -eq '03 Superseded' }).Count
        signed         = @($Model.Documents | Where-Object { $_.Signatures.Count -gt 0 }).Count
        missing        = @($Model.Documents | Where-Object { -not $_.FileExists }).Count
        integrityOk    = $intOk
        integrityTotal = $intTotal
        integrityChecked = [bool]$Integrity
    }

    $dataJson = ConvertTo-Json -InputObject @($records) -Depth 8 -Compress
    if (-not $dataJson.TrimStart().StartsWith('[')) { $dataJson = "[$dataJson]" }
    $metaJson = ConvertTo-Json -InputObject $meta -Depth 4 -Compress
    # Never let document text break out of the <script> data island. JSON structure contains no
    # '<', so every '<' is inside a string value -> escape ALL of them to the JSON unicode escape
    # <. This defeats </script>, <!--<script and any other tokenizer trick, and stays valid
    # JSON (the browser's JSON.parse decodes < back to '<').
    $jsonLt   = [string][char]0x5C + 'u003c'    # the 6 chars < (JSON-escaped '<')
    $dataJson = $dataJson.Replace('<', $jsonLt)
    $metaJson = $metaJson.Replace('<', $jsonLt)

    $template = Get-ViewerTemplate   # build-generated from src/assets/viewer.{html,css,js}
    $html = $template.Replace('/*__META__*/', $metaJson).Replace('/*__DATA__*/', $dataJson)

    $indexPath = Join-Path $OutputRoot 'index.html'
    # write UTF-8 (no BOM) so the browser reads it cleanly
    [System.IO.File]::WriteAllText($indexPath, $html, (New-Object System.Text.UTF8Encoding($false)))
    return $indexPath
}



function New-LandingPage {
    param([string]$OutputRoot, $Events, [string]$Study)
    # Inline HTML-encoder (no System.Web dependency, no load-order hazard).
    $enc = { param($s) ([string]$s).Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;') }
    $cards = foreach ($e in $Events) {
        $href = [Uri]::EscapeUriString("$($e.EventLabel)/index.html")
        $warn = if ($e.Missing -gt 0 -or $e.CopyBad -gt 0) {
            "<span class='w'>$($e.Missing) missing · $($e.CopyBad) mismatch</span>" } else { "<span class='ok'>integrity clean</span>" }
        @"
<a class="card" href="$href">
  <div class="ev">Event $(& $enc $e.Event)</div>
  <div class="meta">$($e.Docs) documents · $($e.Final) final · $($e.InProgress) in&nbsp;progress · $($e.Superseded) superseded</div>
  <div class="meta">$warn</div>
</a>
"@
    }
    $tmpl = Get-LandingTemplate   # build-generated from src/assets/landing.html
    $studyEnc = & $enc $Study
    $html = $tmpl.Replace('__STUDY__', $studyEnc).Replace('__N__', "$($Events.Count)").Replace('__CARDS__', ($cards -join "`n"))
    $landingPath = Join-Path $OutputRoot 'index.html'
    [System.IO.File]::WriteAllText($landingPath, $html, (New-Object System.Text.UTF8Encoding($false)))
    return $landingPath
}



# ----------------------------------------------------------------------------------------
#  Friendly usage (shown when the script is run with no archive)
# ----------------------------------------------------------------------------------------
function Show-Usage {
    $w = 'White'; $c = 'Cyan'; $g = 'Gray'; $gr = 'Green'
    Write-Host ""
    Write-Host "  ===================================================================" -ForegroundColor $c
    Write-Host "   eTMF Archive Humanizer" -ForegroundColor $w
    Write-Host "  ===================================================================" -ForegroundColor $c
    Write-Host ""
    Write-Host "  This tool turns an eTMF export (a .zip file from Viedoc) into" -ForegroundColor $g
    Write-Host "  something you can easily read:" -ForegroundColor $g
    Write-Host ""
    Write-Host "    * tidy folders you can browse - with the final/approved documents" -ForegroundColor $g
    Write-Host "      kept separate from drafts and older (superseded) versions, and" -ForegroundColor $g
    Write-Host "    * a single web page you open in your browser to search and read" -ForegroundColor $g
    Write-Host "      every document, with its history and signatures." -ForegroundColor $g
    Write-Host ""
    Write-Host "  Your original .zip is never changed." -ForegroundColor $g
    Write-Host ""
    Write-Host "  HOW TO USE IT" -ForegroundColor $w
    Write-Host "  -------------" -ForegroundColor $w
    Write-Host "  Keep this script and your .zip in the same folder, and make sure" -ForegroundColor $g
    Write-Host "  this window is IN that folder (use  cd ""C:\path\to\folder""  to" -ForegroundColor $g
    Write-Host "  move there - the prompt shows your current folder).  Then point" -ForegroundColor $g
    Write-Host "  the tool at your .zip, for example:" -ForegroundColor $g
    Write-Host ""
    Write-Host '      .\Convert-EtmfArchive.ps1 -ArchivePath ".\My Study eTMF.zip"' -ForegroundColor $gr
    Write-Host ""
    Write-Host "  Tip: you can drag your .zip file into this window to fill in the path." -ForegroundColor $g
    Write-Host ""
    Write-Host "  When it finishes, open the new folder it creates and double-click" -ForegroundColor $g
    Write-Host "  'index.html' to start browsing." -ForegroundColor $g
    Write-Host ""
    Write-Host "  Add -Open to open it for you automatically:" -ForegroundColor $g
    Write-Host '      .\Convert-EtmfArchive.ps1 -ArchivePath "...your.zip" -Open' -ForegroundColor $gr
    Write-Host ""
    Write-Host "  On a locked-down PC, or sharing this with a colleague?  This longer" -ForegroundColor $g
    Write-Host "  form starts on any Windows machine and lets this one script run" -ForegroundColor $g
    Write-Host "  even if Windows would otherwise block it as unsigned:" -ForegroundColor $g
    Write-Host '      powershell -ExecutionPolicy Bypass -File .\Convert-EtmfArchive.ps1 -ArchivePath "...your.zip"' -ForegroundColor $gr
    Write-Host ""
    Write-Host "  For all options, run:  Get-Help .\Convert-EtmfArchive.ps1 -Full" -ForegroundColor $g
    Write-Host "  ===================================================================" -ForegroundColor $c
    Write-Host ""
}

# ----------------------------------------------------------------------------------------
#  Dry-run summary
# ----------------------------------------------------------------------------------------
function Show-Summary {
    param($Model, $Integrity)

    $b = $Model.Batch
    $docs = $Model.Documents

    Write-Host ""
    Write-Host "==================== eTMF ARCHIVE SUMMARY ====================" -ForegroundColor White
    Write-Host ("  Study (Specification) : {0}" -f $b.SpecificationId)
    Write-Host ("  Study ID              : {0}" -f $b.StudyId)
    Write-Host ("  Event / Transfer ID   : {0} / {1}" -f $b.EventId, $b.TransferId)
    Write-Host ("  Source / TMF-RM ver   : {0} / {1}" -f $b.TransferSource, $b.TmfRmVersion)
    Write-Host ("  Manifest              : {0}" -f $b.ManifestFile)
    Write-Host ("  Documents (files)     : {0}" -f $docs.Count)

    Write-Host "`n  -- By status bucket --" -ForegroundColor White
    $docs | Group-Object Bucket | Sort-Object Name | ForEach-Object {
        Write-Host ("     {0,-16} {1,3}" -f $_.Name, $_.Count)
    }
    Write-Host "`n  -- By lifecycle --" -ForegroundColor White
    $docs | Group-Object Lifecycle | Sort-Object Name | ForEach-Object {
        Write-Host ("     {0,-16} {1,3}" -f $_.Name, $_.Count)
    }
    Write-Host "`n  -- By TMF level --" -ForegroundColor White
    $docs | Group-Object ObjectLevel | Sort-Object Name | ForEach-Object {
        Write-Host ("     {0,-16} {1,3}" -f $_.Name, $_.Count)
    }

    $sigCount = @($docs | Where-Object { $_.Signatures.Count -gt 0 }).Count
    Write-Host ("`n  E-signed documents     : {0}" -f $sigCount)

    if ($Integrity) {
        $tot = $Integrity.Ok + $Integrity.Bad + $Integrity.Missing
        $col = if ($Integrity.Bad -eq 0 -and $Integrity.Missing -eq 0) { 'Green' } else { 'Yellow' }
        Write-Host ("  Integrity (MD5)        : {0}/{1} OK, {2} mismatch, {3} missing" -f `
            $Integrity.Ok, $tot, $Integrity.Bad, $Integrity.Missing) -ForegroundColor $col
    } else {
        Write-Host "  Integrity (MD5)        : skipped" -ForegroundColor DarkGray
    }

    Write-Host "`n  -- Document inventory (preview) --" -ForegroundColor White
    $docs |
        Sort-Object ArtifactNumber, Bucket, Country, OriginalName |
        Select-Object @{n='Art#';e={$_.ArtifactNumber}},
                      @{n='Bucket';e={$_.Bucket}},
                      @{n='Status';e={$_.Status}},
                      @{n='Ver';e={$_.Version}},
                      @{n='Ctry';e={$_.Country}},
                      @{n='Int';e={ if($null -eq $_.IntegrityOk){'-'} elseif($_.IntegrityOk){'ok'} else {'X'} }},
                      @{n='Original file';e={ if($_.OriginalName.Length -gt 42){$_.OriginalName.Substring(0,39)+'...'}else{$_.OriginalName} }} |
        Format-Table -AutoSize | Out-String -Width 160 | Write-Host
    Write-Host "==============================================================" -ForegroundColor White
}

function Show-Recap {
    param([string]$OutputRoot, $Written, [string]$Landing)
    Write-Host ""
    Write-Host "==================== DONE ====================" -ForegroundColor White
    Write-Host ("  Output folder : {0}" -f $OutputRoot) -ForegroundColor Green
    foreach ($e in $Written) {
        $flag = if ($e.Missing -gt 0 -or $e.CopyBad -gt 0) { '  [!]' } else { '' }
        $col  = if ($flag) { 'Yellow' } else { 'Gray' }
        Write-Host ("    Event {0,-4} {1,3} docs  ({2} final / {3} in-progress / {4} superseded)  copied {5}{6}" -f `
            $e.Event, $e.Docs, $e.Final, $e.InProgress, $e.Superseded, $e.Copied, $flag) -ForegroundColor $col
        if ($e.Missing -gt 0) { Write-Host ("           {0} document(s) missing source files" -f $e.Missing) -ForegroundColor Yellow }
        if ($e.CopyBad -gt 0) { Write-Host ("           {0} copy integrity mismatch(es)" -f $e.CopyBad) -ForegroundColor Yellow }
    }
    $open = if ($Landing) { $Landing } else { $Written[0].Index }
    Write-Host ("  Open viewer   : {0}" -f $open) -ForegroundColor Green
    Write-Host "=============================================" -ForegroundColor White
}


function Get-ViewerTemplate {
@'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>eTMF Viewer</title>
<style>  :root{
    --bg:#f4f6f8; --panel:#ffffff; --ink:#1f2733; --muted:#67727f; --line:#e3e8ee;
    --head:#243140; --head2:#2f4055; --accent:#2563a8;
    --final:#1b7f4b; --final-bg:#e6f4ec; --prog:#9a6700; --prog-bg:#fdf3d7; --sup:#5a6573; --sup-bg:#eceef1;
  }
  *{box-sizing:border-box}
  html,body{margin:0;height:100%}
  body{font:14px/1.45 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;color:var(--ink);background:var(--bg)}
  header{background:linear-gradient(180deg,var(--head),var(--head2));color:#fff;padding:14px 20px}
  .hgrid{display:flex;justify-content:space-between;align-items:center;gap:24px;flex-wrap:wrap}
  header h1{margin:0;font-size:18px;font-weight:600}
  header .sub{color:#b9c6d6;font-size:12px;margin-top:4px}
  header .sub b{color:#e7eef6;font-weight:600}
  .counts{display:flex;gap:10px;flex-wrap:wrap}
  .count{background:rgba(255,255,255,.10);border:1px solid rgba(255,255,255,.18);border-radius:8px;padding:6px 12px;text-align:center;min-width:74px}
  .count .n{font-size:18px;font-weight:700;line-height:1}
  .count .l{font-size:10.5px;color:#c4d2e2;text-transform:uppercase;letter-spacing:.4px;margin-top:3px}
  .layout{display:flex;height:calc(100vh - 70px)}
  aside#sidebar{width:300px;min-width:240px;background:var(--panel);border-right:1px solid var(--line);overflow:auto;padding:10px 6px}
  .side-title{font-size:11px;text-transform:uppercase;letter-spacing:.6px;color:var(--muted);padding:6px 10px}
  nav#tree ul{list-style:none;margin:0;padding-left:14px}
  nav#tree li{margin:1px 0}
  .node{display:flex;align-items:center;gap:6px;padding:3px 8px;border-radius:6px;cursor:pointer;color:var(--ink);user-select:none}
  .node:hover{background:#eef3f8}
  .node.active{background:#e2edf8;color:var(--accent);font-weight:600}
  .node .dot{width:8px;height:8px;border-radius:50%;flex:0 0 auto}
  .node .cnt{margin-left:auto;color:var(--muted);font-size:11px;background:#eef1f5;border-radius:10px;padding:0 7px}
  .node.zone{font-weight:600}
  .twist{width:12px;display:inline-block;color:var(--muted);font-size:10px;transition:transform .12s}
  .collapsed > ul{display:none}
  .collapsed .twist{transform:rotate(-90deg)}
  main{flex:1;display:flex;flex-direction:column;overflow:hidden}
  .toolbar{display:flex;gap:10px;align-items:center;padding:10px 16px;border-bottom:1px solid var(--line);background:var(--panel);flex-wrap:wrap}
  #search{flex:1;min-width:200px;padding:8px 11px;border:1px solid var(--line);border-radius:8px;font-size:13px}
  #search:focus{outline:none;border-color:var(--accent);box-shadow:0 0 0 3px rgba(37,99,168,.12)}
  select{padding:7px 9px;border:1px solid var(--line);border-radius:8px;background:#fff;font-size:13px}
  .chips{display:flex;gap:6px}
  .chip{padding:6px 11px;border:1px solid var(--line);border-radius:20px;cursor:pointer;font-size:12.5px;background:#fff;color:var(--muted)}
  .chip:hover{border-color:#c7d2de}
  .chip.on{color:#fff;border-color:transparent}
  .chip.on[data-s="all"]{background:var(--head2)}
  .chip.on[data-s="01 Final"]{background:var(--final)}
  .chip.on[data-s="02 In Progress"]{background:var(--prog)}
  .chip.on[data-s="03 Superseded"]{background:var(--sup)}
  .resultCount{margin-left:auto;color:var(--muted);font-size:12px;white-space:nowrap}
  .tablewrap{overflow:auto;flex:1}
  table{border-collapse:collapse;width:100%;font-size:13px}
  thead th{position:sticky;top:0;background:#f0f3f7;text-align:left;padding:9px 12px;border-bottom:1px solid var(--line);cursor:pointer;white-space:nowrap;font-weight:600;color:#3a4654}
  thead th:hover{background:#e8edf3}
  th .arrow{color:var(--accent);font-size:10px;margin-left:3px}
  tbody td{padding:8px 12px;border-bottom:1px solid #eef2f6;vertical-align:top}
  tbody tr{cursor:pointer}
  tbody tr:hover{background:#f6f9fc}
  td.art{font-variant-numeric:tabular-nums;color:var(--muted);white-space:nowrap}
  td.doc{color:var(--muted);max-width:340px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .badge{display:inline-block;padding:2px 9px;border-radius:20px;font-size:11.5px;font-weight:600;white-space:nowrap}
  .b-final{color:var(--final);background:var(--final-bg)}
  .b-prog{color:var(--prog);background:var(--prog-bg)}
  .b-sup{color:var(--sup);background:var(--sup-bg)}
  .ok{color:var(--final);font-weight:700}.no{color:#b23b3b;font-weight:700}
  .sig{color:var(--accent);font-size:12px}
  /* drawer */
  #overlay{position:fixed;inset:0;background:rgba(20,28,38,.38);opacity:0;pointer-events:none;transition:opacity .15s;z-index:5}
  #overlay.show{opacity:1;pointer-events:auto}
  #drawer{position:fixed;top:0;right:0;height:100%;width:520px;max-width:94vw;background:#fff;box-shadow:-8px 0 26px rgba(0,0,0,.18);transform:translateX(100%);transition:transform .18s;z-index:6;overflow:auto}
  #drawer.show{transform:none}
  .d-head{padding:18px 20px;border-bottom:1px solid var(--line);position:sticky;top:0;background:#fff}
  .d-head h2{margin:0 0 6px;font-size:16px}
  .d-head .x{position:absolute;top:14px;right:16px;border:none;background:#eef1f5;border-radius:8px;width:30px;height:30px;cursor:pointer;font-size:16px;color:var(--muted)}
  .d-body{padding:16px 20px}
  .openbtn{display:inline-block;margin:6px 0 14px;padding:9px 14px;background:var(--accent);color:#fff;border-radius:8px;text-decoration:none;font-weight:600;font-size:13px}
  .openbtn:hover{background:#1d4f86}
  .sec{margin:18px 0 8px;font-size:11px;text-transform:uppercase;letter-spacing:.6px;color:var(--muted);border-bottom:1px solid var(--line);padding-bottom:5px}
  .kv{display:grid;grid-template-columns:130px 1fr;gap:4px 12px;font-size:13px}
  .kv .k{color:var(--muted)}
  .kv .v{word-break:break-word}
  .tl{border-left:2px solid var(--line);margin-left:6px;padding-left:14px}
  .tl .ev{position:relative;margin-bottom:12px}
  .tl .ev::before{content:"";position:absolute;left:-21px;top:3px;width:9px;height:9px;border-radius:50%;background:var(--accent);border:2px solid #fff}
  .tl .ev .when{font-size:11.5px;color:var(--muted)}
  .tl .ev .what{white-space:pre-wrap;margin-top:1px}
  .tl .ev .who{font-size:12px;color:var(--muted)}
  .vh{display:flex;flex-direction:column;gap:5px}
  .vh a{display:flex;align-items:center;gap:8px;padding:6px 9px;border:1px solid var(--line);border-radius:7px;text-decoration:none;color:var(--ink);font-size:12.5px}
  .vh a:hover{background:#f6f9fc}
  .vh a.cur{border-color:var(--accent);background:#eef5fc}
  .sigbox{border:1px solid var(--line);border-radius:8px;padding:10px 12px;margin-bottom:8px;font-size:12.5px;background:#fafbfc}
  .sigbox .nm{font-weight:600}
  footer{padding:8px 16px;color:var(--muted);font-size:11px;background:var(--panel);border-top:1px solid var(--line)}
  .empty{padding:40px;text-align:center;color:var(--muted)}</style>
</head>
<body>
<header>
  <div class="hgrid">
    <div>
      <h1 id="studyName"></h1>
      <div class="sub" id="studySub"></div>
    </div>
    <div class="counts" id="counts"></div>
  </div>
</header>
<div class="layout">
  <aside id="sidebar">
    <div class="side-title">Trial Master File</div>
    <nav id="tree"></nav>
  </aside>
  <main>
    <div class="toolbar">
      <input id="search" type="search" placeholder="Search title, document, artifact #…" autocomplete="off">
      <div class="chips" id="statusChips"></div>
      <select id="countryFilter"><option value="all">All countries</option></select>
      <span class="resultCount" id="resultCount"></span>
    </div>
    <div class="tablewrap">
      <table id="docTable">
        <thead><tr>
          <th data-key="artNo">Artifact #</th>
          <th data-key="title">Title</th>
          <th data-key="originalName">Document</th>
          <th data-key="scope">Country / Site</th>
          <th data-key="version">Ver</th>
          <th data-key="status">Status</th>
          <th data-key="date">Date</th>
          <th data-key="copyOk" title="Integrity">✓</th>
        </tr></thead>
        <tbody id="tbody"></tbody>
      </table>
      <div class="empty" id="empty" style="display:none">No documents match the current filters.</div>
    </div>
  </main>
</div>
<div id="overlay"></div>
<aside id="drawer"></aside>
<footer id="footer"></footer>

<script>const META = /*__META__*/;
const DATA = /*__DATA__*/;

const $ = s => document.querySelector(s);
const esc = s => (s==null?'':String(s)).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const href = p => p ? encodeURI(p) : '#';
// Documents open in a new tab so the viewer stays put.
const linkTgt = ' target="_blank" rel="noopener noreferrer"';
const fmtSize = b => { if(!b) return ''; const u=['B','KB','MB','GB']; let i=0,n=b; while(n>=1024&&i<u.length-1){n/=1024;i++;} return (n<10&&i>0?n.toFixed(1):Math.round(n))+' '+u[i]; };
const fmtTs = s => { if(!s) return ''; const d=new Date(s); return isNaN(d)? s : d.toISOString().slice(0,16).replace('T',' ')+' UTC'; };
const bClass = b => b==='01 Final'?'b-final':b==='02 In Progress'?'b-prog':'b-sup';

const state = { search:'', status:'all', country:'all', node:null, sortKey:'artNo', sortDir:1 };

// ---- header ----
$('#studyName').textContent = META.study || META.studyId || 'eTMF Archive';
$('#studySub').innerHTML =
  `Event <b>${esc(META.event)}</b> · Transfer <b>${esc(META.transferId)}</b> · TMF-RM <b>${esc(META.tmfrm)}</b> · Source <b>${esc(META.source)}</b>`;
const intTxt = META.integrityChecked ? `${META.integrityOk}/${META.integrityTotal}` : 'n/a';
const countDefs = [['Total',META.total],['Final',META.final],['In&nbsp;Progress',META.inProgress],
  ['Superseded',META.superseded],['E-signed',META.signed],['Integrity',intTxt]];
if (META.missing) countDefs.push(['Missing',META.missing]);
$('#counts').innerHTML = countDefs
  .map(([l,n])=>`<div class="count"><div class="n">${n}</div><div class="l">${l}</div></div>`).join('');
$('#footer').innerHTML = `Generated ${esc(META.generated)} from <b>${esc(META.manifest)}</b> · ${META.total} documents · eTMF Archive Humanizer`;

// ---- status chips ----
const chips = [['all','All'],['01 Final','Final'],['02 In Progress','In Progress'],['03 Superseded','Superseded']];
$('#statusChips').innerHTML = chips.map(([s,l])=>`<span class="chip${s==='all'?' on':''}" data-s="${s}">${l}</span>`).join('');
$('#statusChips').addEventListener('click',e=>{ const c=e.target.closest('.chip'); if(!c) return;
  state.status=c.dataset.s; document.querySelectorAll('.chip').forEach(x=>x.classList.toggle('on',x===c)); render(); });

// ---- country filter ----
const countries = [...new Set(DATA.map(d=>d.country).filter(Boolean))].sort();
$('#countryFilter').innerHTML = '<option value="all">All countries</option>' + countries.map(c=>`<option>${esc(c)}</option>`).join('');
$('#countryFilter').addEventListener('change',e=>{ state.country=e.target.value; render(); });

// ---- search ----
$('#search').addEventListener('input',e=>{ state.search=e.target.value.toLowerCase(); render(); });

// ---- tree ----
function buildTree(){
  const t={};
  DATA.forEach(d=>{
    const z=t[d.zone]??=({c:0,f:0,s:{}}); z.c++; if(d.bucket!=='01 Final')z.f++;
    const se=z.s[d.section]??=({c:0,f:0,a:{}}); se.c++; if(d.bucket!=='01 Final')se.f++;
    const a=se.a[d.artifact]??=({c:0,f:0}); a.c++; if(d.bucket!=='01 Final')a.f++;
  });
  const dot=o=>`<span class="dot" style="background:${o.f?'var(--prog)':'var(--final)'}"></span>`;
  let h='<ul><li><div class="node active" data-type="all">All documents<span class="cnt">'+DATA.length+'</span></div></li>';
  for(const zn of Object.keys(t).sort()){ const z=t[zn];
    h+=`<li class="znode"><div class="node zone" data-type="zone" data-v="${esc(zn)}"><span class="twist">▼</span>${dot(z)}${esc(zn)}<span class="cnt">${z.c}</span></div><ul>`;
    for(const sn of Object.keys(z.s).sort()){ const se=z.s[sn];
      h+=`<li class="snode"><div class="node" data-type="section" data-v="${esc(sn)}"><span class="twist">▼</span>${dot(se)}${esc(sn)}<span class="cnt">${se.c}</span></div><ul>`;
      for(const an of Object.keys(se.a).sort()){ const a=se.a[an];
        h+=`<li><div class="node" data-type="artifact" data-v="${esc(an)}">${dot(a)}${esc(an)}<span class="cnt">${a.c}</span></div></li>`;
      }
      h+='</ul></li>';
    }
    h+='</ul></li>';
  }
  $('#tree').innerHTML=h+'</ul>';
}
$('#tree').addEventListener('click',e=>{
  const tw=e.target.closest('.twist');
  if(tw){ tw.closest('li').classList.toggle('collapsed'); e.stopPropagation(); return; }
  const n=e.target.closest('.node'); if(!n) return;
  state.node = n.dataset.type==='all' ? null : {type:n.dataset.type,v:n.dataset.v};
  document.querySelectorAll('#tree .node').forEach(x=>x.classList.toggle('active',x===n));
  render();
});

// ---- filtering + table ----
function matches(d){
  if(state.status!=='all' && d.bucket!==state.status) return false;
  if(state.country!=='all' && d.country!==state.country) return false;
  if(state.node){ if(state.node.type==='zone'&&d.zone!==state.node.v) return false;
    if(state.node.type==='section'&&d.section!==state.node.v) return false;
    if(state.node.type==='artifact'&&d.artifact!==state.node.v) return false; }
  if(state.search){ const h=(d.artNo+' '+d.title+' '+d.originalName+' '+d.newName+' '+d.artifact+' '+d.scope).toLowerCase();
    if(!h.includes(state.search)) return false; }
  return true;
}
function render(){
  let rows=DATA.filter(matches);
  const k=state.sortKey, dir=state.sortDir;
  rows.sort((a,b)=>{ let x=a[k],y=b[k];
    if(k==='version'){x=+x||0;y=+y||0;}
    else if(k==='date'){x=a.dateIso||'';y=b.dateIso||'';}        // sort by ISO date, not the display string
    else {x=(''+x).toLowerCase();y=(''+y).toLowerCase();}
    return x<y?-1*dir:x>y?dir:0; });
  const tb=$('#tbody');
  tb.innerHTML=rows.map((d,i)=>{
    const idx=DATA.indexOf(d);
    const sig=d.signatures&&d.signatures.length?' <span class="sig" title="Electronically signed">✎</span>':'';
    const restr=d.restricted?' <span title="Restricted artifact" style="color:#b23b3b">🔒</span>':'';
    const integ = !d.present ? '<span class="no" title="Source file missing">⚠</span>'
                : d.copyOk ? '<span class="ok" title="MD5 verified">✓</span>'
                : (META.integrityChecked?'<span class="no" title="Integrity mismatch">✗</span>':'');
    return `<tr data-i="${idx}">
      <td class="art">${esc(d.artNo)}</td>
      <td>${esc(d.title)}${sig}${restr}</td>
      <td class="doc" title="${esc(d.originalName)}">${esc(d.originalName)}</td>
      <td>${esc(d.scope||(d.level==='Trial'?'—':d.country))}</td>
      <td>${esc(d.version)}</td>
      <td><span class="badge ${bClass(d.bucket)}">${esc(d.status)}</span></td>
      <td>${esc(d.date)}</td>
      <td>${integ}</td></tr>`;
  }).join('');
  $('#empty').style.display = rows.length?'none':'block';
  $('#resultCount').textContent = rows.length+' of '+DATA.length+' documents';
  document.querySelectorAll('th[data-key]').forEach(th=>{
    const a=th.querySelector('.arrow'); if(a)a.remove();
    if(th.dataset.key===k){ const s=document.createElement('span'); s.className='arrow'; s.textContent=dir>0?'▲':'▼'; th.appendChild(s); }
  });
}
document.querySelectorAll('th[data-key]').forEach(th=>th.addEventListener('click',()=>{
  const k=th.dataset.key; if(state.sortKey===k)state.sortDir*=-1; else {state.sortKey=k;state.sortDir=1;} render();
}));
$('#tbody').addEventListener('click',e=>{ const tr=e.target.closest('tr'); if(tr) openDrawer(+tr.dataset.i); });

// ---- drawer ----
function openDrawer(i){
  const d=DATA[i];
  const versions=DATA.map((x,ix)=>({x,ix})).filter(o=>o.x.documentId&&o.x.documentId===d.documentId&&o.x.artNo===d.artNo&&o.x.country===d.country)
    .sort((a,b)=>(+a.x.version||0)-(+b.x.version||0));
  const vh = versions.length>1 ? `<div class="sec">Version history</div><div class="vh">`+
    versions.map(o=>`<a class="${o.ix===i?'cur':''}" href="${href(o.x.relPath)}"${linkTgt}>
        <span class="badge ${bClass(o.x.bucket)}">${esc(o.x.status)}</span> v${esc(o.x.version)} <span style="color:var(--muted)">${esc(o.x.originalName)}</span></a>`).join('')+`</div>` : '';
  const sigs = d.signatures&&d.signatures.length ? `<div class="sec">Signatures (${d.signatures.length})</div>`+
    d.signatures.map(s=>`<div class="sigbox"><div class="nm">${esc(s.signer)}</div>
      <div>${esc(s.methodology)} signature · ${esc(fmtTs(s.dateTime))}</div>
      <div style="color:var(--muted)">${esc(s.reason)}</div></div>`).join('') : '';
  const audit = d.audit&&d.audit.length ? `<div class="sec">Audit trail (${d.audit.length})</div><div class="tl">`+
    d.audit.slice().reverse().map(a=>`<div class="ev"><div class="when">${esc(fmtTs(a.ts))} · ${esc(a.type)}</div>
      <div class="what">${esc(a.event)}</div><div class="who">${esc(a.user)}</div></div>`).join('')+`</div>` : '';
  const kv = (k,v)=>v?`<div class="k">${esc(k)}</div><div class="v">${esc(v)}</div>`:'';
  $('#drawer').innerHTML=`
    <div class="d-head"><button class="x" onclick="closeDrawer()">×</button>
      <h2>${esc(d.title)}</h2>
      <span class="badge ${bClass(d.bucket)}">${esc(d.status)}</span>
      <span style="color:var(--muted);font-size:12px"> ${esc(d.artNo)} · v${esc(d.version)}</span>
    </div>
    <div class="d-body">
      ${d.present && d.relPath ? `<a class="openbtn" href="${href(d.relPath)}"${linkTgt}>Open document ↗</a>`
        : `<div class="openbtn" style="background:var(--sup);cursor:default">⚠ Source file not available</div>`}
      <div class="sec">Document</div>
      <div class="kv">
        ${kv('Artifact',d.artifact)}${kv('Sub-artifact / title',d.title)}${kv('TMF level',d.level)}
        ${kv('Country',d.country)}${kv('Site / org',d.scope)}${kv('Person',d.person)}
        ${kv('Version',d.version)}${kv('Version state',d.versionState)}${kv('Lifecycle',d.lifecycle)}
        ${kv('Restricted',d.restricted?'Yes':'')}${kv('Language',d.language)}${kv('Translation',d.translation)}
        ${kv(d.dateDesc||'Artifact date',d.date)}${kv('Expires',d.expiry)}${kv('Document ID',d.documentId)}
        ${kv('Size',fmtSize(d.sizeBytes))}
      </div>
      <div class="sec">Files</div>
      <div class="kv">
        ${kv('Original name',d.originalName)}${kv('Stored as',d.newName)}
        ${kv('Integrity (MD5)',d.integrity)}
        <div class="k">Verified</div><div class="v">${d.copyOk?'<span class="ok">✓ copy matches manifest</span>':(META.integrityChecked?'<span class="no">✗ mismatch</span>':'not checked')}</div>
      </div>
      ${vh}${sigs}${audit}
    </div>`;
  $('#overlay').classList.add('show'); $('#drawer').classList.add('show');
}
function closeDrawer(){ $('#overlay').classList.remove('show'); $('#drawer').classList.remove('show'); }
$('#overlay').addEventListener('click',closeDrawer);
document.addEventListener('keydown',e=>{ if(e.key==='Escape')closeDrawer(); });

buildTree(); render();</script>
</body>
</html>
'@
}

function Get-LandingTemplate {
@'
<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>eTMF Archive — __STUDY__</title><style>
body{font:15px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f4f6f8;color:#1f2733;margin:0}
header{background:linear-gradient(180deg,#243140,#2f4055);color:#fff;padding:22px 26px}
header h1{margin:0;font-size:20px}header p{margin:6px 0 0;color:#b9c6d6;font-size:13px}
.wrap{max-width:820px;margin:24px auto;padding:0 18px;display:flex;flex-direction:column;gap:12px}
.card{display:block;background:#fff;border:1px solid #e3e8ee;border-radius:10px;padding:16px 18px;text-decoration:none;color:inherit;transition:box-shadow .12s,border-color .12s}
.card:hover{border-color:#2563a8;box-shadow:0 4px 14px rgba(20,40,70,.10)}
.card .ev{font-weight:600;font-size:16px}.card .meta{color:#67727f;font-size:13px;margin-top:4px}
.ok{color:#1b7f4b}.w{color:#9a6700}
footer{max-width:820px;margin:8px auto 30px;padding:0 18px;color:#67727f;font-size:12px}
</style></head><body>
<header><h1>__STUDY__</h1><p>eTMF archive — __N__ milestone export(s). Select an event to open its viewer.</p></header>
<div class="wrap">__CARDS__</div>
<footer>Generated by eTMF Archive Humanizer.</footer>
</body></html>
'@
}

# ========================================================================================
#  MAIN
# ========================================================================================
$script:OutputRoot = $null
$source = $null

# No archive supplied -> show friendly instructions and stop (instead of prompting).
if ([string]::IsNullOrWhiteSpace($ArchivePath)) {
    Show-Usage
    return
}

try {
    Write-Log "eTMF Archive Humanizer - parse, classify & humanize" 'STEP'

    # Resolve where the output will live (sibling of the input by default).
    $inputItem = Get-Item -LiteralPath $ArchivePath
    $baseDir   = if ($inputItem.PSIsContainer) {
                     if ($inputItem.Parent) { $inputItem.Parent.FullName } else { $inputItem.FullName }  # drive/fs root has no parent
                 } else { $inputItem.DirectoryName }

    $source = Resolve-EtmfSource -ArchivePath $ArchivePath
    $manifests = @(Find-ExchangeManifests -Root $source.Root)
    Write-Log ("Found {0} manifest file(s)." -f $manifests.Count) 'OK'

    # ---- Phase 4: parse + classify + verify every manifest (one per study event/milestone) ----
    $events = foreach ($mf in $manifests) {
        Write-Log "Parsing manifest: $($mf.Name)" 'STEP'
        $model = Import-EtmfManifest -ManifestFile $mf
        Write-Log ("Parsed {0} document(s) from {1} object(s)." -f $model.Documents.Count, $model.ObjectCount) 'OK'
        $integrity = $null
        if (-not $SkipIntegrityCheck) {
            Write-Log "Verifying source integrity (base64 MD5)..." 'STEP'
            $integrity = Invoke-IntegrityCheck -Documents $model.Documents
        }
        Show-Summary -Model $model -Integrity $integrity
        [pscustomobject]@{ Model = $model; Integrity = $integrity; Manifest = $mf }
    }
    $events = @($events)

    if ($DryRun) {
        Write-Log "DryRun: parsed & classified only; nothing written." 'INFO'
        return
    }

    # ---- output root (computed once from the study) ----
    $b0 = $events[0].Model.Batch
    $studyName = ConvertTo-SafeName $b0.SpecificationId
    if ([string]::IsNullOrWhiteSpace($studyName)) { $studyName = ConvertTo-SafeName $b0.StudyId }
    $outputRoot = if ($OutputPath) { $OutputPath }
                  else { Join-Path $baseDir (ConvertTo-SafeName ($script:DefaultOutputNameTemplate -f $studyName)) }
    $script:OutputRoot = $outputRoot

    if (Test-Path -LiteralPath $outputRoot) {
        if ($Force) {
            # Safety: only ever auto-delete a directory that looks like our own prior output,
            # so a mistyped -OutputPath + -Force cannot wipe an unrelated folder.
            $looksLikeOurs = @('index.html', 'inventory.csv', '_log.txt', 'Documents') |
                Where-Object { Test-Path -LiteralPath (Join-Path $outputRoot $_) }
            if (-not $looksLikeOurs) {
                throw "Refusing to -Force delete '$outputRoot': it does not look like a previous run's output " +
                      "(no index.html / inventory.csv / _log.txt / Documents). Remove it manually or choose another -OutputPath."
            }
            if ($PSCmdlet.ShouldProcess($outputRoot, 'Remove existing output directory')) {
                Write-Log "Output exists; -Force given, removing: $outputRoot" 'WARN'
                Remove-Item -LiteralPath $outputRoot -Recurse -Force
            }
        } else {
            throw "Output folder already exists: $outputRoot  (use -Force to overwrite)"
        }
    }
    New-Item -ItemType Directory -Path $outputRoot -Force | Out-Null

    # ---- Phases 5 + 6: per event, build tree + inventory + viewer ----
    $multi      = $events.Count -gt 1
    $written    = [System.Collections.Generic.List[object]]::new()
    $usedLabels = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($ev in $events) {
        $model = $ev.Model; $integrity = $ev.Integrity; $batch = $model.Batch

        $eventLabel = ''
        $eventDir   = $outputRoot
        if ($multi) {
            $eventLabel = ConvertTo-SafeName ("{0} {1}" -f $batch.EventId, $batch.TransferId).Trim()
            if ([string]::IsNullOrWhiteSpace($eventLabel)) { $eventLabel = ConvertTo-SafeName $batch.TransferId }
            if (-not $usedLabels.Add($eventLabel)) {                    # two manifests share EventId+TransferId
                $n = 2
                while (-not $usedLabels.Add("$eventLabel ($n)")) { $n++ }
                $eventLabel = "$eventLabel ($n)"
                Write-Log "Duplicate event label; disambiguated to '$eventLabel'" 'WARN'
            }
            $eventDir = Join-Path $outputRoot $eventLabel
            New-Item -ItemType Directory -Path $eventDir -Force | Out-Null
        }

        Write-Log ("Humanizing event '{0}' -> {1}" -f $batch.EventId, $eventDir) 'STEP'
        $result = Invoke-Humanize -Model $model -OutputRoot $eventDir `
                    -KeepOriginalNames:$KeepOriginalNames -VerifyCopy:(-not $SkipIntegrityCheck)
        Write-Log ("Copied {0} file(s); {1} missing; {2} copy mismatch(es)." -f `
            $result.Copied, $result.Missing, $result.CopyBad) ($(if ($result.CopyBad -eq 0 -and $result.Missing -eq 0) { 'OK' } else { 'WARN' }))

        $invPath = Join-Path $eventDir 'inventory.csv'
        $result.Inventory | Export-Csv -LiteralPath $invPath -NoTypeInformation -Encoding UTF8
        Write-Log "Inventory written: $invPath" 'OK'

        $indexPath = New-HtmlViewer -Model $model -OutputRoot $eventDir -Integrity $integrity
        Write-Log "Viewer written: $indexPath" 'OK'

        $docs = $model.Documents
        $written.Add([pscustomobject]@{
            Event = $batch.EventId; EventLabel = $eventLabel; Dir = $eventDir; Index = $indexPath
            Docs = $docs.Count; Copied = $result.Copied; Missing = $result.Missing; CopyBad = $result.CopyBad
            Final      = @($docs | Where-Object { $_.Bucket -eq '01 Final' }).Count
            InProgress = @($docs | Where-Object { $_.Bucket -eq '02 In Progress' }).Count
            Superseded = @($docs | Where-Object { $_.Bucket -eq '03 Superseded' }).Count
        })
    }

    $landing = $null
    if ($multi) {
        $landing = New-LandingPage -OutputRoot $outputRoot -Events $written -Study $b0.SpecificationId
        Write-Log "Landing page written: $landing" 'OK'
    }

    Show-Recap -OutputRoot $outputRoot -Written $written -Landing $landing

    if ($Open) {
        $toOpen = if ($landing) { $landing } else { $written[0].Index }
        try {
            if ($IsWindows -or $env:OS -match 'Windows') { Start-Process $toOpen }
            elseif ($IsMacOS) { & open $toOpen }
            else { & xdg-open $toOpen 2>$null }
        } catch { Write-Log "Could not auto-open viewer: $($_.Exception.Message)" 'WARN' }
    }
}
catch {
    Write-Log $_.Exception.Message 'ERROR'
    throw
}
finally {
    # Persist the run log next to the output (if we produced one).
    if ($script:OutputRoot -and (Test-Path -LiteralPath $script:OutputRoot)) {
        $logPath = Join-Path $script:OutputRoot '_log.txt'
        $script:LogLines | Set-Content -LiteralPath $logPath -Encoding UTF8 -ErrorAction SilentlyContinue
    }
    # Remove the temp extraction (files have been copied out by now).
    if ($null -ne $source -and $source.IsTemp -and (Test-Path -LiteralPath $source.TempDir)) {
        Remove-Item -LiteralPath $source.TempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

