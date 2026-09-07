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
      * flags records worth a human look (identical content filed twice; a name that mentions
        another study's code) and applies the reviewer's own -ExclusionList decisions,
      * copies + renames files into a readable tree, verifying MD5 on every copy,
      * writes inventory.csv, a run log, and a self-contained index.html viewer.

    The eTMF tree is the backbone: Zone > Section > Artifact, with the status split made
    *inside* each artifact:  Artifact > <status bucket> > <Country/Site> > files.
    -FinalOnly produces the client-facing counterpart: current approved documents only, with no
    status subfolders, plus a note recording what was left out.
    A multi-milestone export (more than one manifest) is MERGED into that one tree, because each
    milestone export is a full snapshot of the TMF: the same file exported in several milestones
    is published once (the newest snapshot's record wins) and the milestone is recorded as data,
    in inventory.csv and in the viewer, instead of as a folder level. Use -SplitByMilestone for
    the older per-milestone view. The original archive is never modified.

.PARAMETER ArchivePath
    Path to the eTMF archive .zip, OR a folder that already contains the extracted archive
    (i.e. the folder holding the *_exchange.xml manifest next to the TransferID folder).

.PARAMETER OutputPath
    Destination folder for the humanized output. Defaults to "<input directory>\<StudyName>",
    where the folder name is built from $DefaultOutputNameTemplate near the top of the script.

.PARAMETER FinalOnly
    Produce a copy containing ONLY the current, approved documents - the version the eTMF records
    as Current AND finalised or locked - for handing to a client or sponsor. Superseded versions,
    drafts and anything excluded by -ExclusionList are left out, and the status subfolders are
    dropped (everything in this copy is final). An "About this folder.txt" states what the copy
    contains, what was left out and how many of each, so the scope of the handover is on record.
    This is a convenience copy, NOT the complete trial master file - run without the switch for that.

.PARAMETER ExclusionList
    Path to a CSV of documents to keep out of the main tree - the reviewer's decisions about
    documents that do not belong in this TMF (wrongly uploaded, from another study, ...).
    Columns: DocumentId (required), Version (optional - only that version), Reason (optional).
    Matching documents are filed under "_Excluded" with the reason recorded; nothing is deleted,
    and a row that matches no document is reported rather than ignored.

.PARAMETER SplitByMilestone
    Write one subfolder (and one viewer) per milestone export instead of merging them into a
    single TMF tree. Only has an effect when the archive holds more than one manifest.

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

# ============================================================================
#  GENERATED FILE - DO NOT EDIT DIRECTLY.
#  Built from src/ by build/Build.ps1. Edit the sources there and rebuild.
# ============================================================================
# NOTE ON THE HELP BLOCK ABOVE - two things silently break `Get-Help -Full`, degrading it to the
# bare syntax line (the command README.md and .NOTES both tell users to run):
#   1. ANY free text inside the block before the first .KEYWORD - which is why this banner lives
#      here and not in there;
#   2. #requires above it. Above it, PowerShell stops associating the
# comment-based help with the script and 'Get-Help -Full' degrades to the bare syntax line - which
# is the command README.md and the .NOTES section both tell users to run.
#requires -Version 5.1
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Position = 0)]   # optional: with no value, the script prints friendly usage and exits
    [string]$ArchivePath,

    [Parameter()]
    [string]$OutputPath,

    [Parameter()]
    [switch]$FinalOnly,

    [Parameter()]
    [string]$ExclusionList,

    [Parameter()]
    [switch]$SplitByMilestone,

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
# Stamped into the run log, the viewer and inventory.csv. A support mail that says "it still fails"
# is unanswerable without it: we could not tell whether a reported log came from the build that
# contains a given fix. Bump it with any change to behaviour or output.
$script:ToolVersion  = '1.1.0'
$script:RunStamp     = (Get-Date).ToString('yyyyMMdd-HHmmss')   # names the log; one file per run
$script:LogLines     = [System.Collections.Generic.List[string]]::new()

# Default output folder name when -OutputPath is not supplied. {0} = study name.
# Edit this to set your house style, e.g. '{0}', '{0} - eTMF', or 'eTMF - {0}'.
$script:DefaultOutputNameTemplate   = '{0}'
$script:FinalOnlyOutputNameTemplate = '{0} (final documents only)'


# ----------------------------------------------------------------------------------------
#  Logging
# ----------------------------------------------------------------------------------------
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'OK', 'STEP')]
        [string]$Level = 'INFO'
    )
    # Messages can quote untrusted text (zip entry names, manifest FILENAME). One line per entry
    # is the log's whole structure, so a name carrying CR/LF could forge convincing "[OK]" lines
    # in the audit record - strip control characters before the message is ever recorded.
    $Message = $Message -replace '[\x00-\x1F]', '?'
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

# Console AND log, in one call. Everything the operator sees on screen belongs in the run record:
# the summary tables and the final recap are the most useful part of a run, and for as long as they
# went only to Write-Host, the log a customer emailed us was missing exactly what they were reading.
function Write-Out {
    param([string]$Message = '', [string]$ForegroundColor = 'Gray')
    foreach ($line in ($Message -split "`r?`n")) {
        # One line per entry is the log's whole structure; strip the control characters that
        # manifest-derived text could carry (newlines already handled by the split above).
        $script:LogLines.Add(($line -replace '[\x00-\x09\x0B\x0C\x0E-\x1F]', '?').TrimEnd())
    }
    Write-Host $Message -ForegroundColor $ForegroundColor
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

# Split an archive-relative path into its segments: both separators, empties and "." removed.
# ONE rule, shared by the extractor (Expand-EtmfZip) and the lookup (Resolve-ContentPath). If the
# two ever normalise differently, a document that IS on disk gets reported as a missing source
# file - so keep this the single definition and never inline a variant of it.
function Get-ContentPathSegments {
    param([string]$Path)
    # Callers MUST wrap the call in @(): a function return unrolls, so a single-segment result
    # would arrive as a bare string and $parts.Count would throw under Set-StrictMode. (Do not
    # "fix" that with a ",@(...)" return here - combined with the caller's @() it nests instead.)
    return @($Path -split '[\\/]+' | Where-Object { $_ -ne '' -and $_ -ne '.' })
}

# The archive path a CONTENTURL points at, verbatim ('/'-joined, leading zip-name segment
# dropped). This is the record's identity and the key into the extraction index, so it is kept
# RAW: any normalisation here would be a lossy derivation, and deriving identity is exactly what
# put one document's bytes under another document's record in earlier designs.
function Get-ContentUrlTail {
    param([string]$ContentUrl)
    # Both separators: Get-ContentPathSegments and Resolve-ContentPath already accept backslashes,
    # and splitting on '/' alone made a backslash-separated CONTENTURL report itself as "malformed",
    # which tells a customer their export is broken when it is merely written the other way round.
    $segments = $ContentUrl -split '[\\/]+'
    if ($segments.Count -lt 2) { return $null }
    $parts = @(Get-ContentPathSegments (($segments | Select-Object -Skip 1) -join '/'))
    if ($parts.Count -eq 0 -or ($parts -contains '..')) { return $null }
    return ($parts -join '/')
}

# Map an EMS CONTENTURL to an on-disk path. CONTENTURL is "<zipname>.zip/<TransferID>/...";
# strip the leading zip-name segment and resolve the remainder beside the manifest.
# SECURITY: the manifest is untrusted, so the resolved path is canonicalized and must stay
# CONTAINED under $ManifestDir. A traversal ("../") or rooted CONTENTURL returns $null
# (the referenced file is then recorded as a gap, never read or copied).
# FOLDER INPUT ONLY. Zip input never comes here: it resolves through the extraction index, which
# is exact. -SanitizeSegments matches a folder a user pre-extracted with 7-Zip or Explorer, which
# substitute "_" for the forbidden characters; that substitution is many-to-one, so its caller
# must confirm the match by checksum before binding.
function Resolve-ContentPath {
    param([string]$ContentUrl, [string]$ManifestDir, [switch]$SanitizeSegments)
    $segments = $ContentUrl -split '/'
    if ($segments.Count -lt 2) { return $null }
    $parts = @(Get-ContentPathSegments (($segments | Select-Object -Skip 1) -join '/'))
    if ($parts.Count -eq 0) { return $null }
    if ($parts -contains '..') { return $null }                        # traversal -> reject
    if ($SanitizeSegments) { $parts = @($parts | ForEach-Object { ConvertTo-SafeName $_ }) }
    $rel  = $parts -join [IO.Path]::DirectorySeparatorChar
    $base = [IO.Path]::GetFullPath($ManifestDir)
    # .NET Framework (5.1) throws on "*?<>|" in a path; treat that as unresolvable, not fatal.
    try   { $full = [IO.Path]::GetFullPath((Join-Path $base $rel)) }
    catch { return $null }
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


# Human milestone label for a manifest: "09 Database Lock_exchange.xml" -> "09 Database Lock".
# The BATCH's EVENTID is Viedoc's internal event id and does NOT match the milestone number shown
# in the eTMF UI (a reviewer reading "11" for what the study calls milestone 9 cannot reconcile the
# output against the system) - the milestone's real name is carried by the manifest's file name.
# Falls back to the event id only when the name carries no label of its own.
function Get-MilestoneLabel {
    param([string]$ManifestFileName, [string]$EventId)
    # The extractor may have appended " (2)" to de-duplicate two identically named manifests.
    $stem = ([string]$ManifestFileName) -replace '(?i)[_\- ]?exchange(\s*\(\d+\))?\.xml$', ''
    $stem = $stem.Trim().Trim('_', '-', ' ')
    if (-not [string]::IsNullOrWhiteSpace($stem)) { return $stem }
    if (-not [string]::IsNullOrWhiteSpace($EventId)) { return "Event $EventId" }
    return 'Milestone'
}

# A document's path INSIDE its milestone snapshot: everything below the TransferID segment.
# Every milestone export re-exports the same documents under a new TransferID root, so this is the
# half of the cross-snapshot identity that stays constant (the MD5 is the other half).
function Get-SnapshotRelativePath {
    param([string]$ContentUrl, [string]$TransferId)
    $segs = @($ContentUrl -split '[\\/]+' | Where-Object { $_ -ne '' -and $_ -ne '.' })
    if ($segs.Count -eq 0) { return '' }
    $idx = -1
    if (-not [string]::IsNullOrWhiteSpace($TransferId)) { $idx = [array]::IndexOf($segs, $TransferId) }
    if ($idx -lt 0) { $idx = 0 }                       # fallback: drop the leading "<name>.zip" segment
    if (($idx + 1) -ge $segs.Count) { return $segs[$segs.Count - 1] }
    return (@($segs[($idx + 1)..($segs.Count - 1)]) -join '/')
}

# Cross-snapshot identity of one document VERSION RECORD.
#
# Content alone is NOT identity, and assuming it was cost this tool two critical defects:
#   * A real Viedoc export records a document's version history as several OBJECTs that share ONE
#     physical file - same CONTENTURL, same INTEGRITY - differing only in OBJECTVERSION. Keying on
#     content collapsed v1/v2 into v3 and deleted their audit trails from a regulated archive.
#   * Import-EtmfManifest already de-dupes CONTENTURL *within* an OBJECT, so any repeat that reaches
#     the merge is BY CONSTRUCTION a different artifact record - one document filed under two
#     artifact numbers is an ordinary TMF pattern. Keying on content threw one of them away, with
#     its title, its audit trail and its e-signature.
# So the key is the record's full identity plus its content: the same document version, re-exported
# at a later milestone, still matches on every one of these fields (which is the collapse we DO
# want), while anything that differs in identity or version stays a separate record.
# Returns $null when the manifest states no checksum: identity cannot be established, and merging on
# a guess is the failure mode this whole design exists to prevent.
function Get-EtmfDocumentKey {
    param($Doc)
    if ([string]::IsNullOrWhiteSpace($Doc.Integrity)) { return $null }
    $rel = Get-SnapshotRelativePath $Doc.ContentUrl $Doc.TransferId
    if ([string]::IsNullOrWhiteSpace($rel)) { return $null }
    return (@($Doc.Integrity, $rel, $Doc.DocumentId, $Doc.ArtifactNumber,
              $Doc.Country, $Doc.SiteId, $Doc.Version) -join '|')
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
#  Version resolution (which version of a document is the current one)
# ----------------------------------------------------------------------------------------

# The key that ties every VERSION of one Viedoc document together, across snapshots.
# DocumentId is Viedoc's own document identity; artifact + country + site are a safety net (and
# the same grouping the viewer's version history already uses). No DocumentId means no identity:
# return $null and let the record stand alone rather than group on a resemblance.
function Get-EtmfVersionGroupKey {
    param($Doc)
    if ([string]::IsNullOrWhiteSpace($Doc.DocumentId)) { return $null }
    return ('{0}|{1}|{2}|{3}' -f $Doc.DocumentId, $Doc.ArtifactNumber, $Doc.Country, $Doc.SiteId)
}

# Compare two OBJECTVERSION values. Dotted numeric versions are compared SEGMENT BY SEGMENT, so
# v10 beats v9 (a string compare gets that backwards) and v1.10 beats v1.9 (parsing them as decimals
# gets THAT backwards, and would additionally depend on the machine's decimal separator - which must
# never decide which version of a clinical document is the current one). Anything else: ordinal.
function Compare-EtmfVersion {
    param([string]$A, [string]$B)
    $sa = @(([string]$A).Trim() -split '\.')
    $sb = @(([string]$B).Trim() -split '\.')
    $numeric = $true
    foreach ($seg in ($sa + $sb)) { if ($seg -notmatch '^\d+$') { $numeric = $false; break } }
    if ($numeric) {
        # Compared as DIGIT STRINGS (leading zeros stripped, longer number wins, then ordinal), never
        # cast to a number: a manifest is untrusted text and an absurdly long version segment would
        # overflow [int64] and abort the entire conversion over a cosmetic field.
        $max = [Math]::Max($sa.Count, $sb.Count)
        for ($i = 0; $i -lt $max; $i++) {
            $x = if ($i -lt $sa.Count) { ([string]$sa[$i]).TrimStart('0') } else { '' }   # "1.0" equals "1"
            $y = if ($i -lt $sb.Count) { ([string]$sb[$i]).TrimStart('0') } else { '' }
            if ($x.Length -ne $y.Length) { if ($x.Length -lt $y.Length) { return -1 } else { return 1 } }
            $c = [string]::CompareOrdinal($x, $y)
            if ($c -lt 0) { return -1 }
            if ($c -gt 0) { return 1 }
        }
        return 0
    }
    $c = [string]::CompareOrdinal([string]$A, [string]$B)
    if ($c -lt 0) { return -1 }
    if ($c -gt 0) { return 1 }
    return 0
}

# ----------------------------------------------------------------------------------------
#  Review heuristics (flag for a human; never decide)
# ----------------------------------------------------------------------------------------

# Does this text mention a study code that is NOT this study's? Returns the offending token, or ''.
#
# Deliberately narrow, because a noisy flag is worse than none: the token must be UPPERCASE letters
# followed by at least THREE digits (the shape of a study code - "XYZ-002"), which is what keeps
# ordinary document numbering out of it ("Consent Form-01", "Amendment-001" in mixed case). The
# study's own code must have that shape too, and text that already names this study is never
# flagged. Even then this only ever RAISES A FLAG - see Set-EtmfReviewFlags.
function Get-EtmfForeignStudyCode {
    param([string]$Text, [string[]]$StudyCodes)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $norm = { param($s) ([string]$s).ToUpperInvariant() -replace '[-_ ]', '' }
    # Shape-test the study's own id UPPER-CASED: a study registered as "test-001" would otherwise
    # fail the test and silently switch the whole check off.
    $own = @($StudyCodes | Where-Object { $_ -and ((& $norm $_) -cmatch '^[A-Z]{2,10}\d{3,6}$') } | ForEach-Object { & $norm $_ })
    if ($own.Count -eq 0) { return '' }                     # no code-shaped study id -> nothing to compare against
    $textNorm = & $norm $Text
    foreach ($c in $own) { if ($textNorm.Contains($c)) { return '' } }   # names this study -> fine
    # Tokens that look like study codes but never are: standards, quality marks, CVs, dates. Without
    # this, "ISO 9001" and "CV_2021" in a file name raise a "belongs to another study" flag, and a
    # heuristic that cries wolf is one nobody reads.
    $notStudyCodes = @('ISO', 'IEC', 'EN', 'DIN', 'ASTM', 'ICH', 'GCP', 'CFR', 'EU', 'GDPR', 'CV', 'ID', 'NO', 'REF', 'VER', 'V')
    foreach ($m in [regex]::Matches($Text, '\b([A-Z]{2,10})[-_ ]?\d{3,6}\b')) {
        if ($notStudyCodes -contains $m.Groups[1].Value) { continue }
        if ($own -notcontains (& $norm $m.Value)) { return $m.Value }
    }
    return ''
}


# ----------------------------------------------------------------------------------------
#  Input resolution: accept a .zip (extract to temp) or an already-extracted folder.
# ----------------------------------------------------------------------------------------

# Extract a .zip entry by entry instead of via ZipFile::ExtractToDirectory, and give every
# content file a NEUTRAL name (00001.bin, 00002.bin, ...) in a flat temp folder.
#
# The temp tree is deleted at the end of the run and no user ever sees it, so it does not need
# readable names - and the moment we stop deriving temp paths from the archive's own text, a
# whole class of defects disappears rather than having to be defended against:
#   * "Illegal characters in path." on Windows PowerShell 5.1 - the reported bug. .NET Framework
#     rejects " < > | * ? and control chars in a path; we no longer put archive text in a path.
#   * Name collisions - sanitizing is many-to-one, so two documents could land on one file and
#     one document's bytes could be published under the other's identity. Index names are unique.
#   * Zip-slip ("../" entries) - the destination never comes from the entry, so it cannot escape.
#   * MAX_PATH - temp paths are now short regardless of how deep the archive nests.
# Identity is therefore never derived: Files maps the entry's VERBATIM archive path to the file
# we wrote, and Import-EtmfManifest looks each CONTENTURL up by exact key.
# Manifests are the one exception - they keep their (sanitized) name because the recap and the
# viewer display it. A single unextractable entry is logged and skipped, so it shows up as a
# missing-source gap rather than killing the run.
function Expand-EtmfZip {
    param([string]$ZipPath, [string]$Destination)

    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    # ORDINAL: keys are raw entry paths, where two spellings differing only in case are two
    # different documents. The value is the temp file; '' marks an ambiguous key (see below).
    $files     = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::Ordinal)
    $manifests = New-Object 'System.Collections.Generic.List[object]'
    $skipped = 0; $count = 0

    $zip = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
    try {
        foreach ($entry in $zip.Entries) {
            if ([string]::IsNullOrEmpty($entry.Name)) { continue }        # directory entry
            $parts = @(Get-ContentPathSegments $entry.FullName)
            if ($parts.Count -eq 0) { continue }
            $rawRel = $parts -join '/'
            $count++

            $isManifest = $entry.Name -like '*exchange.xml'
            if ($isManifest) {
                # Keep a readable name: it is shown in the recap and the viewer. Sanitized because
                # this one IS a real filename, and de-duplicated because two events could share it.
                $leaf = ConvertTo-SafeName $entry.Name
                $dest = Join-Path $Destination $leaf
                $n = 2
                while (Test-Path -LiteralPath $dest) {
                    $dest = Join-Path $Destination ('{0} ({1}){2}' -f [IO.Path]::GetFileNameWithoutExtension($leaf), $n, [IO.Path]::GetExtension($leaf))
                    $n++
                }
            }
            else {
                $dest = Join-Path $Destination ('{0:d5}.bin' -f $count)
            }

            try {
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $dest, $true)
            }
            catch {
                Write-Log "Could not extract '$($entry.FullName)': $($_.Exception.Message)" 'WARN'
                $skipped++; continue
            }

            if ($isManifest) {
                # Prefix = where this manifest sits in the archive. A CONTENTURL is relative to it,
                # so prefix + tail reconstructs the entry path exactly.
                $prefix = ''
                if ($parts.Count -gt 1) { $prefix = (@($parts[0..($parts.Count - 2)]) -join '/') }
                $manifests.Add([pscustomobject]@{ Path = $dest; Prefix = $prefix })
                continue
            }

            if ($files.ContainsKey($rawRel)) {
                # Two entries with the identical archive path: nothing can tell which one a
                # CONTENTURL means, so refuse both rather than guess (both become visible gaps).
                Write-Log "Ambiguous archive: more than one entry is named '$rawRel' - neither will be matched to a manifest record." 'WARN'
                $files[$rawRel] = ''
            }
            else { $files[$rawRel] = $dest }
        }
    }
    finally { $zip.Dispose() }

    if ($skipped -gt 0) {
        Write-Log "$skipped zip entr(ies) could not be extracted - they will show as missing source files." 'WARN'
    }
    return [pscustomobject]@{ FileCount = $files.Count; Skipped = $skipped; Files = $files; Manifests = $manifests }
}

function Resolve-EtmfSource {
    param([string]$ArchivePath)

    if (-not (Test-Path -LiteralPath $ArchivePath)) {
        throw "ArchivePath not found: $ArchivePath"
    }
    $item = Get-Item -LiteralPath $ArchivePath

    if ($item.PSIsContainer) {
        Write-Log "Using extracted archive folder: $($item.FullName)" 'INFO'
        # No index: we extracted nothing, so the lookup resolves names against the folder itself.
        return [pscustomobject]@{ Root = $item.FullName; TempDir = $null; IsTemp = $false; ContentMap = $null; Manifests = $null }
    }

    if ($item.Extension -ieq '.zip') {
        $temp = Join-Path ([IO.Path]::GetTempPath()) ("etmf_" + [Guid]::NewGuid().ToString('N').Substring(0, 8))
        New-Item -ItemType Directory -Path $temp -Force | Out-Null
        $script:TempDir = $temp        # recorded now so the finally in main.ps1 can clean up
        Write-Log "Extracting archive to temp: $temp" 'STEP'                # even if extraction throws
        $x = Expand-EtmfZip -ZipPath $item.FullName -Destination $temp
        Write-Log "Extracted $($x.FileCount) file(s)" 'INFO'
        # Manifests come from the extractor, not from a directory scan: their temp names are
        # neutral, and each carries the archive prefix its CONTENTURLs are relative to.
        return [pscustomobject]@{ Root = $temp; TempDir = $temp; IsTemp = $true; ContentMap = $x.Files; Manifests = $x.Manifests }
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
    # -ContentMap: raw archive path -> extracted temp file, from Expand-EtmfZip ($null for folder
    # input). -ContentPrefix: where this manifest sits in the archive, since a CONTENTURL is
    # relative to the manifest but the index is keyed from the archive root.
    param([System.IO.FileInfo]$ManifestFile, $ContentMap = $null, [string]$ContentPrefix = "")

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

    $eventId    = $batchEl.GetAttribute('EVENTID')
    # Trimmed: the TransferID anchors every CONTENTURL lookup and the snapshot-relative path, so a
    # stray space in the attribute would silently switch the merge to its fallback and stop it
    # collapsing anything.
    $transferId = ([string]$batchEl.GetAttribute('TRANSFERID')).Trim()
    # Milestone name comes from the manifest's FILE NAME, not EVENTID - see Get-MilestoneLabel.
    $milestone  = Get-MilestoneLabel $ManifestFile.Name $eventId

    # Milestones/IsMerged are present on every batch (a single manifest is a one-milestone merge),
    # so the humanizer and the viewer never have to branch on how the model was built.
    $batch = [pscustomobject]@{
        StudySystemId   = $batchEl.GetAttribute('STUDYSYSTEMID')
        StudyId         = $batchEl.GetAttribute('STUDYID')
        EventId         = $eventId
        TransferId      = $transferId
        TransferSource  = $batchEl.GetAttribute('TRANSFERSOURCEID')
        SpecificationId = $batchEl.GetAttribute('SPECIFICATIONID')
        TmfRmVersion    = $batchEl.GetAttribute('TMFRMVERSION')
        ManifestFile    = $ManifestFile.Name
        ManifestDir     = $manifestDir
        Milestone       = $milestone
        Milestones      = @($milestone)
        IsMerged        = $false
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

            $integrity = Get-NodeText $file 'INTEGRITY'
            $diskPath  = $null
            $exists    = $false
            # A gap has to explain ITSELF. "Missing source file: X" told a customer nothing, cost a
            # round of mail asking for the archive, and still left us guessing; these three states
            # are distinguishable at the point of failure, so record which one it was.
            $gapReason = ''
            $lookupKey = ''

            if ($null -ne $ContentMap) {
                # Zip input: we extracted this archive, so the index knows exactly which file this
                # record is. The key is the entry's VERBATIM archive path, so nothing is derived -
                # there is nothing to collide and no guess to make. Missing key, or '' (an
                # ambiguous archive), means the document is honestly recorded as a gap.
                $tail = Get-ContentUrlTail $contentUrl
                if ($null -eq $tail) {
                    $gapReason = 'the CONTENTURL is malformed or points outside the archive'
                }
                else {
                    $key = $tail
                    if ($ContentPrefix) { $key = $ContentPrefix + '/' + $tail }
                    $lookupKey = $key
                    $mapped = $null
                    if (-not $ContentMap.TryGetValue($key, [ref]$mapped)) {
                        $gapReason = 'the archive contains no entry with this exact path'
                    }
                    elseif ([string]::IsNullOrEmpty($mapped)) {
                        $gapReason = 'two or more archive entries share this exact path, so neither can be identified'
                    }
                    elseif (-not (Test-Path -LiteralPath $mapped -PathType Leaf)) {
                        $gapReason = 'the archive entry could not be extracted (see the extraction warning above)'
                    }
                    else { $diskPath = $mapped; $exists = $true }
                }
            }
            else {
                # Folder input: resolve against the folder the user handed us.
                $diskPath  = Resolve-ContentPath $contentUrl $manifestDir
                $lookupKey = $diskPath
                $exists    = ($null -ne $diskPath) -and (Test-Path -LiteralPath $diskPath -PathType Leaf)
                if (-not $exists) {
                    # 7-Zip and Explorer substitute "_" for the characters Windows forbids, but that
                    # substitution is many-to-one and we have no index to disambiguate it. Bind only
                    # if the bytes are the ones the manifest asked for - checked even under
                    # -SkipIntegrityCheck, because a wrong match here is what stays silent.
                    $alt = Resolve-ContentPath $contentUrl $manifestDir -SanitizeSegments
                    if (($null -ne $alt) -and (Test-Path -LiteralPath $alt -PathType Leaf)) {
                        if ($integrity -and ((Get-FileMd5Base64 $alt) -ceq $integrity)) {
                            $diskPath = $alt; $exists = $true
                        }
                        else {
                            $gapReason = 'a file with a sanitized name is present but its content does not match the manifest checksum'
                            Write-Log ("Ignored a sanitized-name match for '{0}': its content does not match the manifest checksum, so it is recorded as a gap rather than risk the wrong document." -f (Get-NodeText $file 'FILENAME')) 'WARN'
                        }
                    }
                    elseif ($null -eq $diskPath) {
                        $gapReason = 'the CONTENTURL does not resolve to a path inside the manifest folder'
                    }
                    else {
                        $gapReason = 'no file exists at this path in the extracted folder'
                    }
                }
            }

            $documents.Add([pscustomobject]@{
                Milestone      = $milestone
                AllMilestones  = @($milestone)   # every snapshot this exact file appeared in (merge)
                TransferId     = $transferId     # anchors this record's CONTENTURL, even once merged
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
                Integrity      = $integrity
                DiskPath       = $diskPath
                FileExists     = $exists
                GapReason      = $gapReason   # why it could not be bound (empty when it was)
                LookupKey      = $lookupKey   # the exact path/key we looked for
                SizeBytes      = if ($exists) { (Get-Item -LiteralPath $diskPath).Length } else { 0 }
                IntegrityOk    = $null   # filled by integrity pass (source)
                AuditTrail     = $auditTrail
                Signatures     = $signatures
                Metadata       = $metadata
                SupersededBy   = ''      # filled by Resolve-EtmfVersions when a newer version exists
                ReviewFlag     = ''      # filled by Set-EtmfReviewFlags (flag only - nothing moves)
                ReviewReason   = ''
                Excluded       = $false  # filled by Set-EtmfExclusions from -ExclusionList
                ExclusionReason = ''
                NewName        = ''      # filled by Phase 5 (copy/rename)
                RelPath        = ''      # filled by Phase 5 (path relative to output root)
                CopyIntegrityOk = $null  # filled by Phase 5 (verify the copy)
            })
        }
    }

    return [pscustomobject]@{ Batch = $batch; Documents = $documents; ObjectCount = $objectNodes.Count }
}

# ----------------------------------------------------------------------------------------
#  Merge the per-milestone models into ONE model
# ----------------------------------------------------------------------------------------
# A milestone export is a SNAPSHOT of the whole eTMF at that moment, so an archive with 11
# milestones re-exports the same documents 11 times. Publishing a folder per milestone (what this
# tool used to do) hands the reader 11 partial TMFs and no way to search across them, so the
# default is now ONE tree - which means the repeats have to be collapsed, or the copier's collision
# guard would file the same document as "... (2)", "... (3)", ... eleven times over.
#
# Identity is EXACT, never approximate, and it is the record's IDENTITY, not its content:
# Get-EtmfDocumentKey requires the checksum, the path, DocumentId, artifact, country, site AND
# OBJECTVERSION to match. Content alone is not identity - a real export gives every version of a
# document the SAME file and checksum - so collapsing on content deleted version history. Deciding
# which version supersedes which is then a separate concern (Resolve-EtmfVersions).
#
# The record kept is the one from the LATEST snapshot, so the newest milestone's verdict on
# version state and lifecycle is what the reader sees; the labels of every snapshot the file
# appeared in are carried along for the inventory.
function Merge-EtmfEvents {
    param($Events)

    # Ascending by TransferID (a yyyyMMddHHmmss stamp) so that "last one wins" = newest wins.
    # The manifest name is a tie-breaker: PowerShell 5.1's Sort-Object is not stable, and a
    # merge whose winner depends on enumeration order would not be reproducible.
    $ordered = @($Events | Sort-Object -Property @{ Expression = { [string]$_.Model.Batch.TransferId } },
                                                 @{ Expression = { [string]$_.Model.Batch.ManifestFile } })

    # ORDINAL: the key carries a base64 checksum and a file path, both case-sensitive. (A plain
    # [ordered]@{} would compare them case-insensitively.) Insertion order is kept so the merged
    # model lists documents oldest-snapshot-first rather than in hash order.
    $byKey = New-Object System.Collections.Specialized.OrderedDictionary ([System.StringComparer]::Ordinal)
    $collapsed = 0
    $objectCount = 0

    foreach ($ev in $ordered) {
        $objectCount += $ev.Model.ObjectCount
        $label = $ev.Model.Batch.Milestone
        # Keys already used by THIS manifest. Collapse is a cross-snapshot operation: two records in
        # one manifest are two records even when every identity field matches, because lifecycle is
        # derived from their audit events - which no key can contain - so collapsing them let XML
        # order decide whether a document came out "Finalized / 01 Final" or "In Review".
        $seenThisEvent = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
        foreach ($d in $ev.Model.Documents) {
            $key = Get-EtmfDocumentKey $d
            if ($null -eq $key) {
                # No checksum to identify it by: keep the record on its own rather than guess.
                $key = 'unkeyed:' + [Guid]::NewGuid().ToString('N')
            }
            elseif ($seenThisEvent.Contains($key)) {
                $n = 2
                while ($seenThisEvent.Contains(('{0}#{1}' -f $key, $n))) { $n++ }
                # Numbered BY ORDER OF APPEARANCE, deliberately: the same suffix is derived in every
                # snapshot, so a within-manifest duplicate still pairs with its counterpart in the
                # next export. A GUID here would make each snapshot's copy unique and publish N of
                # them. Pairing therefore assumes these records keep their order between exports.
                $key = '{0}#{1}' -f $key, $n
            }
            [void]$seenThisEvent.Add($key)
            if ($byKey.Contains($key)) {
                $prev = $byKey[$key]
                $d.AllMilestones = @(@($prev.AllMilestones) + $label | Select-Object -Unique)
                # "Later wins" must never cost us a FILE. If the newer snapshot lists this record but
                # its file is absent, unextractable or ambiguous there, while an older snapshot bound
                # and verified it, carry the binding across: the key contains INTEGRITY, so those
                # bytes are provably this record's content. Without this, an export defect in the
                # newest milestone silently deletes a document that IS in the archive - the tool
                # manufacturing the very "missing source file" symptom it was asked to explain.
                if ((-not $d.FileExists) -and $prev.FileExists) {
                    # MEASURE the donor before trusting it. The INTEGRITY in the key is what the
                    # manifest CLAIMS, not what the file on disk contains - and under
                    # -SkipIntegrityCheck nothing has ever hashed it - so without this the transplant
                    # would be the D6 wrong-bytes bug wearing a new hat. Checked always, exactly like
                    # the sanitized-name folder bind below.
                    $donorOk = $false
                    try { $donorOk = ((Get-FileMd5Base64 $prev.DiskPath) -ceq $d.Integrity) } catch { $donorOk = $false }
                    if ($donorOk) {
                        $d.DiskPath    = $prev.DiskPath
                        $d.FileExists  = $true
                        $d.IntegrityOk = $true
                        $d.SizeBytes   = $prev.SizeBytes
                        $d.LookupKey   = $prev.LookupKey
                        $d.GapReason   = ''
                        Write-Log ("'{0}' has no usable file in '{1}'; using the verified identical bytes exported in '{2}'." -f `
                            $d.OriginalName, $label, $prev.Milestone) 'INFO'
                    }
                    else {
                        Write-Log ("'{0}' has no file in '{1}' and the copy exported in '{2}' does not match the manifest checksum - recorded as a gap rather than publishing unverified bytes." -f `
                            $d.OriginalName, $label, $prev.Milestone) 'WARN'
                    }
                }
                $byKey[$key] = $d          # later snapshot wins on metadata
                $collapsed++
            }
            else {
                $d.AllMilestones = @($label)
                $byKey[$key] = $d
            }
        }
    }

    $documents = [System.Collections.Generic.List[object]]::new()
    foreach ($k in @($byKey.Keys)) { $documents.Add($byKey[$k]) }

    $unkeyed = @($documents | Where-Object { [string]::IsNullOrWhiteSpace($_.Integrity) }).Count
    if ($unkeyed -gt 0 -and $ordered.Count -gt 1) {
        Write-Log ("{0} record(s) state no checksum, so they cannot be recognised as the same document across milestones and appear once per milestone export. They are listed individually in inventory.csv." -f $unkeyed) 'WARN'
    }

    $latest = $ordered[$ordered.Count - 1].Model.Batch
    $batch  = $latest.PSObject.Copy()
    $batch.Milestones = @($ordered | ForEach-Object { $_.Model.Batch.Milestone })
    $batch.IsMerged   = ($ordered.Count -gt 1)

    return [pscustomobject]@{
        Batch = $batch; Documents = $documents; ObjectCount = $objectCount; Collapsed = $collapsed
    }
}

# ----------------------------------------------------------------------------------------
#  Version resolution across snapshots
# ----------------------------------------------------------------------------------------
# THE reported problem: "we have different version numbers of the same document, without knowing
# which one is the correct final and approved version."
#
# OBJECTVERSIONSTATE is a per-SNAPSHOT fact. A v1 exported at milestone 3 says "Current" in that
# manifest forever, even though milestone 11 carries a v2 that replaced it - so after merging, both
# versions land in "01 Final" and the reader is left to guess. Supersession by a newer version of
# the same document is not a judgement: it is derivable, and this derives it.
#
# Rule: inside one version group, the winner is the HIGHEST version the archive still calls
# Current (if none is Current the group is superseded anyway, and the highest version stands in for
# it). Every OTHER version is filed under "03 Superseded" and records what replaced it. Records
# sharing the winner's version are left alone - they are the attachments of that one document (D5).
function Resolve-EtmfVersions {
    param($Model)

    $groups = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[object]]' ([System.StringComparer]::Ordinal)
    foreach ($d in $Model.Documents) {
        $key = Get-EtmfVersionGroupKey $d
        if ($null -eq $key) { continue }             # no document identity -> never grouped
        if (-not $groups.ContainsKey($key)) { $groups[$key] = [System.Collections.Generic.List[object]]::new() }
        $groups[$key].Add($d)
    }

    $demoted = 0; $resolved = 0
    foreach ($key in @($groups.Keys)) {
        $members = $groups[$key]
        # More than one DISTINCT version? Several records of the same version are one document's files.
        $distinct = @($members | Group-Object Version).Count
        if ($distinct -lt 2) { continue }

        $currents = @($members | Where-Object { $_.VersionState -eq 'Current' })
        # @() around the whole conditional: a one-element result would otherwise unroll to a scalar
        # and $pool[0] throws under StrictMode.
        $pool     = @(if ($currents.Count -gt 0) { $currents } else { $members })
        $winner   = $pool[0]
        foreach ($m in $pool) { if ((Compare-EtmfVersion $m.Version $winner.Version) -gt 0) { $winner = $m } }
        # The WINNING VERSION is whatever the archive says is current, present or not: preferring a
        # version whose file happens to be present would quietly promote an older document to
        # "current", which misrepresents the TMF. Presence only breaks a TIE - two records of the
        # SAME version (possible since collapse is cross-snapshot only), one bound and one a gap.
        if (-not $winner.FileExists) {
            $sameVersionPresent = @($pool | Where-Object { $_.FileExists -and (Compare-EtmfVersion $_.Version $winner.Version) -eq 0 })
            if ($sameVersionPresent.Count -gt 0) { $winner = $sameVersionPresent[0] }
        }
        $resolved++

        if (-not $winner.FileExists) {
            Write-Log ("The current version of '{0}' (v{1}) has no file in the archive; older versions are still filed as superseded." -f `
                $winner.OriginalName, $winner.Version) 'WARN'
        }
        $sharedChecksum = @($members | Group-Object Integrity | Where-Object { $_.Count -gt 1 -and $_.Name })
        if ($sharedChecksum.Count -gt 0) {
            Write-Log ("'{0}': versions {1} share one file and checksum - normal for this export format (the version history of one document), recorded here because it is also what a version RELABELLED between exports would look like." -f `
                $winner.OriginalName, (@($members | ForEach-Object { 'v' + $_.Version } | Select-Object -Unique) -join ', ')) 'INFO'
        }
        $replacedBy = 'v{0} ({1}, {2})' -f $winner.Version, $winner.Status, $winner.Milestone
        foreach ($m in $members) {
            if ((Compare-EtmfVersion $m.Version $winner.Version) -eq 0) { continue }
            $m.SupersededBy = $replacedBy
            if ($m.Bucket -ne '03 Superseded') {
                # It called itself Current in its own snapshot; a newer version says otherwise.
                $m.Bucket = '03 Superseded'
                $m.Status = 'Superseded'
                $demoted++
            }
        }
    }

    return [pscustomobject]@{ Demoted = $demoted; Groups = $resolved }
}

# ----------------------------------------------------------------------------------------
#  Review flags + exclusions ("sort out the documents that should not be there")
# ----------------------------------------------------------------------------------------
# The other half of the customer's ask: an export "may contain incorrect documents (which in the
# worst case belong to another study)". NOTHING in the manifest says a document does not belong -
# that is a human judgement about the study, not a fact in the data - so the tool must not pretend
# to make it. It does two separate things instead:
#
#   Set-EtmfReviewFlags  RAISES A FLAG and moves nothing. One flag is a fact (two records with
#                        identical content), one is an explicitly conservative heuristic (a file
#                        name naming another study's code). A false positive costs a glance.
#   Set-EtmfExclusions   ACTS, but only on a decision a person wrote down in -ExclusionList. The
#                        document is filed under _Excluded/ with the stated reason - kept, not
#                        deleted, and reproduced identically on every future run.
function Set-EtmfReviewFlags {
    param($Model)

    $flagged = 0
    $addFlag = {
        param($Doc, [string]$Flag, [string]$Reason)
        $Doc.ReviewFlag   = (@($Doc.ReviewFlag, $Flag)     | Where-Object { $_ }) -join '; '
        $Doc.ReviewReason = (@($Doc.ReviewReason, $Reason) | Where-Object { $_ }) -join '; '
    }

    # (1) FACT: the same bytes filed as two different documents. This is the "duplicates" half of
    # the complaint, and it is exact - no threshold, no similarity, just one checksum in two records.
    $byContent = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[object]]' ([System.StringComparer]::Ordinal)
    foreach ($d in $Model.Documents) {
        if ([string]::IsNullOrWhiteSpace($d.Integrity)) { continue }
        if (-not $byContent.ContainsKey($d.Integrity)) { $byContent[$d.Integrity] = [System.Collections.Generic.List[object]]::new() }
        $byContent[$d.Integrity].Add($d)
    }
    foreach ($key in @($byContent.Keys)) {
        $members = $byContent[$key]
        if ($members.Count -lt 2) { continue }
        # Flag only what a TMF manager would ACT on: the same file filed under two different
        # ARTIFACTS, i.e. in two different places in the TMF, which is a filing question someone has
        # to answer. Everything else that shares bytes is ordinary: a document's own attachments and
        # versions, and above all a template - one cover letter reused for DEU/SWE/JPN/USA as
        # several country records under one artifact. Flagging those produced 10 flags on a
        # 36-document archive, 8 of them that single cover letter, which buries the flag that
        # matters ("Possible foreign study") in noise and teaches the reader to ignore the column.
        $artifacts = @($members | ForEach-Object { $_.ArtifactNumber } | Select-Object -Unique)
        if ($artifacts.Count -lt 2) { continue }
        foreach ($m in $members) {
            $others = @($members | Where-Object { $_.ArtifactNumber -ne $m.ArtifactNumber } |
                        ForEach-Object { '{0} {1}' -f $_.ArtifactNumber, $_.OriginalName } | Select-Object -Unique)
            & $addFlag $m 'Filed under two artifacts' ("identical content is also filed under: {0}" -f (($others | Select-Object -First 3) -join ', '))
            $flagged++
        }
    }

    # (2) HEURISTIC: the name mentions another study's code.
    $codes = @($Model.Batch.StudyId, $Model.Batch.SpecificationId)
    foreach ($d in $Model.Documents) {
        $foreign = Get-EtmfForeignStudyCode ("{0} {1}" -f $d.OriginalName, $d.ArtifactTitle) $codes
        if ($foreign) {
            & $addFlag $d 'Possible foreign study' ("the name mentions '{0}', which is not this study's code ({1}) - check it belongs here" -f $foreign, $Model.Batch.StudyId)
            $flagged++
        }
    }

    return @($Model.Documents | Where-Object { $_.ReviewFlag }).Count
}

# Read the reviewer's exclusion decisions. Keyed on DocumentId (the id the inventory and the viewer
# both show), optionally narrowed to one Version; Reason is carried into the output so the decision
# is auditable where it took effect.
function Import-EtmfExclusions {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "-ExclusionList file not found: $Path" }
    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -eq 0) {
        Write-Log "Exclusion list is empty: $Path" 'WARN'
        return @()
    }
    $columns = @($rows[0].PSObject.Properties.Name)
    if ($columns -notcontains 'DocumentId') {
        throw "-ExclusionList must have a 'DocumentId' column (found: $($columns -join ', ')). Optional: Version, ArtifactNumber, Country, SiteId, Reason."
    }
    # DocumentId ALONE IS NOT UNIQUE. In a real export the same id is reused across countries (3 of
    # 12 in the reference archive span two), so a bare id can take out documents the reviewer never
    # meant to touch - in the one feature whose job is controlling what reaches a client. These
    # optional columns let a row say exactly which document it means; Set-EtmfExclusions warns when
    # a row is coarse enough to hit several.
    $get = { param($Row, $Name) if ($columns -contains $Name) { ([string]$Row.$Name).Trim() } else { '' } }

    $out = foreach ($r in $rows) {
        $id = ([string]$r.DocumentId).Trim()
        if (-not $id) { continue }
        [pscustomobject]@{
            DocumentId     = $id
            Version        = & $get $r 'Version'
            ArtifactNumber = & $get $r 'ArtifactNumber'
            Country        = & $get $r 'Country'
            SiteId         = & $get $r 'SiteId'
            Reason         = & $get $r 'Reason'
            Matched        = 0
            # DISTINCT documents hit, not raw matches: under -SplitByMilestone one document is
            # matched once per snapshot, and counting those as "several documents" would cry wolf
            # about the very precision warning that exists to be believed.
            MatchedKeys    = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
        }
    }
    return @($out)
}

function Set-EtmfExclusions {
    param($Model, $Exclusions)

    $excluded = 0
    foreach ($d in $Model.Documents) {
        foreach ($x in $Exclusions) {
            if ([string]$d.DocumentId -ne [string]$x.DocumentId) { continue }
            if ($x.Version        -and ((Compare-EtmfVersion $d.Version $x.Version) -ne 0)) { continue }
            if ($x.ArtifactNumber -and ($d.ArtifactNumber -ne $x.ArtifactNumber)) { continue }
            if ($x.Country        -and ($d.Country -ne $x.Country)) { continue }
            if ($x.SiteId         -and ([string]$d.SiteId -ne $x.SiteId)) { continue }
            if ($d.Excluded) { break }                       # already excluded by an earlier row
            $d.Excluded        = $true
            $d.ExclusionReason = if ($x.Reason) { $x.Reason } else { 'excluded by the exclusion list (no reason given)' }
            $d.Bucket          = '04 Excluded'
            $d.Status          = 'Excluded'
            $x.Matched         = $x.Matched + 1
            [void]$x.MatchedKeys.Add(('{0}|{1}|{2}|{3}' -f $d.DocumentId, $d.ArtifactNumber, $d.Country, $d.SiteId))
            $excluded++
            # NAME what was excluded. A count alone cannot show a reviewer that their one row for a
            # German cover letter also removed the Swedish one.
            Write-Log ("Excluded: DocumentId {0} | {1} | {2} | v{3} | '{4}' - {5}" -f `
                $d.DocumentId, $d.ArtifactNumber, $(if ($d.Country) { $d.Country } else { 'Trial' }), `
                $d.Version, $d.OriginalName, $d.ExclusionReason) 'WARN'
            break
        }
    }
    return $excluded
}

# Report exclusion rows that did nothing, or did more than a reviewer is likely to have meant.
# Called ONCE after every model has been processed, so a row that legitimately matches in one
# milestone but not another is not reported as a miss.
# After exclusions: a SupersededBy that names an excluded document sends the reader looking for
# something no longer in the tree. Say so instead.
function Update-EtmfSupersededReferences {
    param($Model)
    $excludedKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
    foreach ($d in $Model.Documents) {
        if ($d.Excluded) {
            $k = Get-EtmfVersionGroupKey $d
            if ($k) { [void]$excludedKeys.Add(('{0}|v{1}' -f $k, $d.Version)) }
        }
    }
    if ($excludedKeys.Count -eq 0) { return }
    foreach ($d in $Model.Documents) {
        if ($d.Excluded -or -not $d.SupersededBy) { continue }
        $k = Get-EtmfVersionGroupKey $d
        if (-not $k) { continue }
        # SupersededBy starts "v<version> (" - recover the version it names
        $ver = ([regex]::Match($d.SupersededBy, '^v([^ ]+) ')).Groups[1].Value
        if ($ver -and $excludedKeys.Contains(('{0}|v{1}' -f $k, $ver))) {
            $d.SupersededBy = '{0} - EXCLUDED, see _Excluded' -f $d.SupersededBy
        }
    }
}

function Write-EtmfExclusionOutcome {
    param($Exclusions)
    foreach ($x in $Exclusions) {
        if ($x.Matched -eq 0) {
            # Nearly always a mistyped id, and silently excluding nothing is the worst outcome
            # available: the reviewer believes a document was removed when it was not.
            $narrow = @(
                if ($x.Version)        { "version '$($x.Version)'" }
                if ($x.ArtifactNumber) { "artifact '$($x.ArtifactNumber)'" }
                if ($x.Country)        { "country '$($x.Country)'" }
                if ($x.SiteId)         { "site '$($x.SiteId)'" }
            ) -join ', '
            $where = if ($narrow) { " with $narrow" } else { '' }
            Write-Log ("Exclusion list: no document has DocumentId '{0}'{1} - nothing was excluded for that row. Check it against inventory.csv." -f $x.DocumentId, $where) 'WARN'
        }
        elseif ($x.MatchedKeys.Count -gt 1 -and -not ($x.ArtifactNumber -or $x.Country -or $x.SiteId)) {
            Write-Log ("Exclusion list: DocumentId '{0}' matched {1} different documents (the same id is reused across artifacts/countries). If you meant only one, add an ArtifactNumber, Country or SiteId column to that row - the excluded documents are named above." -f `
                $x.DocumentId, $x.MatchedKeys.Count) 'WARN'
        }
    }
}

# Reduce a model to the documents a client copy may contain: the version the eTMF calls Current
# AND finalised/locked, minus anything the reviewer excluded. Returns what was left out, because a
# partial TMF that does not say it is partial is the problem this tool exists to fix.
function Select-EtmfFinalDocuments {
    param($Model)

    $kept = [System.Collections.Generic.List[object]]::new()
    $withheld = [pscustomobject]@{ Superseded = 0; InProgress = 0; Excluded = 0; NoFile = 0; Flagged = 0 }
    foreach ($d in $Model.Documents) {
        if ($d.Excluded)                     { $withheld.Excluded   = $withheld.Excluded + 1;   continue }
        if ($d.Bucket -eq '03 Superseded')   { $withheld.Superseded = $withheld.Superseded + 1; continue }
        if ($d.Bucket -eq '02 In Progress')  { $withheld.InProgress = $withheld.InProgress + 1; continue }
        if (-not $d.FileExists)              { $withheld.NoFile     = $withheld.NoFile + 1 }
        if ($d.ReviewFlag)                   { $withheld.Flagged    = $withheld.Flagged + 1 }
        $kept.Add($d)
    }
    $Model.Documents = $kept
    return $withheld
}

# ----------------------------------------------------------------------------------------
#  Integrity verification (base64 MD5)
# ----------------------------------------------------------------------------------------
function Invoke-IntegrityCheck {
    param([System.Collections.Generic.List[object]]$Documents)
    $ok = 0; $bad = 0; $missing = 0
    foreach ($d in $Documents) {
        if (-not $d.FileExists) { $missing++; $d.IntegrityOk = $false; continue }
        if ([string]::IsNullOrWhiteSpace($d.Integrity)) {
            # No checksum stated is not corruption. Reporting it as a MISMATCH would send a QA
            # reviewer looking for a damaged file that is fine.
            $bad++; $d.IntegrityOk = $false
            Write-Log "No checksum stated in the manifest for '$($d.OriginalName)' - the copy could not be verified." 'WARN'
            continue
        }
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
    # Sanitize BEFORE any path API sees it: OriginalName is the verbatim manifest <FILENAME>, and
    # on .NET Framework (5.1) GetExtension/GetFileNameWithoutExtension throw "Illegal characters
    # in path." for " < > | and control chars - the same crash this tool now handles at extraction.
    # Doing it here also stops an unsafe extension surviving into the destination via "+ $ext".
    $safeOriginal = ConvertTo-SafeName $Doc.OriginalName
    $ext = [System.IO.Path]::GetExtension($safeOriginal)
    if ($KeepOriginal) { return $safeOriginal }
    $stem   = [System.IO.Path]::GetFileNameWithoutExtension($safeOriginal)
    $suffix = " v{0} [{1}]" -f $Doc.Version, $Doc.Status
    if ($MaxLength -gt 0) {
        $room = $MaxLength - $suffix.Length - $ext.Length
        if ($room -lt 8) { $room = 8 }
        if ($stem.Length -gt $room) {
            # Keep the START **and the END** of the name, with '~' marking the cut. Clinical file
            # names are distinguished by their tail ("...-01" / "...-02", a date, a version), so
            # head-only truncation collapses different documents onto one name; the collision guard
            # then files them as " (2)", which reads to a reviewer like a duplicate of the first.
            [int]$tailRoom = [Math]::Min(12, [int][Math]::Floor(($room - 1) / 2))
            $head = $stem.Substring(0, $room - 1 - $tailRoom).TrimEnd()
            $stem = $head + '~' + $stem.Substring($stem.Length - $tailRoom)
        }
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
        [switch]$VerifyCopy,
        # -FinalOnly: every document in the tree is final, so a "01 Final" level under every
        # artifact would be a folder that says nothing.
        [switch]$OmitBucketFolder
    )
    $treeRoot = Join-Path $OutputRoot 'Documents'
    New-Item -ItemType Directory -Path $treeRoot -Force | Out-Null
    # Documents a reviewer excluded are KEPT (this is an archive) but filed outside the TMF tree,
    # so the folder that gets handed to a client contains only what belongs in it.
    $excludedRoot = Join-Path $OutputRoot '_Excluded'

    $used      = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $inventory = [System.Collections.Generic.List[object]]::new()
    $copied = 0; $missing = 0; $copyBad = 0; $shortened = 0
    $targetMaxPath = 250   # keep generated full paths comfortably under the Windows MAX_PATH (260)

    foreach ($d in $Model.Documents) {
        # Anchor on the record's OWN TransferID: after a merge the documents come from several
        # snapshots, so the model's batch is not the one this CONTENTURL was written against.
        $pp   = Get-ContentPathParts $d.ContentUrl $d.TransferId
        $note = ''

        if ($d.FileExists) {
            $zone     = ConvertTo-SafeName (Get-PrettyTmfSegment $pp.Zone)
            $section  = ConvertTo-SafeName (Get-PrettyTmfSegment $pp.Section)
            $artifact = ConvertTo-SafeName ("{0} {1}" -f $d.ArtifactNumber, $d.ArtifactTitle)

            $root   = if ($d.Excluded) { $excludedRoot } else { $treeRoot }
            $folder = Join-Path (Join-Path (Join-Path $root $zone) $section) $artifact
            if (-not $OmitBucketFolder) { $folder = Join-Path $folder $d.Bucket }
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
                # Name the file we actually WROTE: the collision guard above may have suffixed it,
                # and a QA reviewer reconciling the log against the tree must land on the real one.
                if (-not $copyOk) { $copyBad++; Write-Log "Copy integrity MISMATCH: $(Split-Path $dest -Leaf)" 'WARN' }
            }
            $copied++

            $d.NewName         = Split-Path $dest -Leaf
            $d.RelPath         = ($dest.Substring($OutputRoot.Length)).TrimStart('/', '\') -replace '\\', '/'
            $d.CopyIntegrityOk = $copyOk
        }
        else {
            # Say WHY, and say what was looked for - this log is what a customer emails back to us.
            Write-Log ("Missing source file: '{0}' - {1}. CONTENTURL: {2} | looked for: {3}" -f `
                $d.OriginalName, $d.GapReason, $d.ContentUrl, $d.LookupKey) 'WARN'
            $missing++
            $note = if ($d.GapReason) { "MISSING SOURCE FILE: $($d.GapReason)" } else { 'MISSING SOURCE FILE' }
        }

        $inventory.Add([pscustomobject]@{
            Milestone      = $d.Milestone                            # snapshot this record came from
            AllMilestones  = (@($d.AllMilestones) -join '; ')        # every snapshot it appeared in
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
            SupersededBy   = $d.SupersededBy      # the version that replaced this one, if any
            ReviewFlag      = $d.ReviewFlag       # raised by the tool, decided by a human
            ReviewReason    = $d.ReviewReason
            Excluded        = if ($d.Excluded) { 'Yes' } else { 'No' }
            ExclusionReason = $d.ExclusionReason
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
            ToolVersion    = $script:ToolVersion   # so a CSV separated from its folder still says what made it
        })
    }

    if ($shortened -gt 0) {
        Write-Log ("{0} long file name(s) shortened (the ~ marks where text was cut) to stay within the Windows 260-char path limit; originals are in inventory.csv." -f $shortened) 'INFO'
    }

    return [pscustomobject]@{
        Inventory = $inventory; TreeRoot = $treeRoot
        Copied = $copied; Missing = $missing; CopyBad = $copyBad
    }
}


function New-HtmlViewer {
    param($Model, [string]$OutputRoot, $Integrity, [switch]$FinalOnlyCopy)

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
        # Anchor on the record's own TransferID - after a merge the batch belongs to one snapshot.
        $pp    = Get-ContentPathParts $d.ContentUrl $d.TransferId
        $scope = @($pp.Scope | ForEach-Object { Get-PrettyScope $_ })
        $iso   = ConvertFrom-EmsDate $d.ArtifactDate
        $ctry  = if ($d.Country -and $alpha2.ContainsKey($d.Country)) { $alpha2[$d.Country] } else { $d.Country }
        [pscustomobject]@{
            milestone    = $d.Milestone
            milestones   = @($d.AllMilestones)      # every snapshot it appeared in (viewer filters on this)
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
            supersededBy = $d.SupersededBy
            reviewFlag   = $d.ReviewFlag
            reviewReason = $d.ReviewReason
            excluded     = [bool]$d.Excluded
            exclusionReason = $d.ExclusionReason
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
            gapReason    = $d.GapReason
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
        toolVersion    = $script:ToolVersion
        runStamp       = $script:RunStamp
        milestone      = $batch.Milestone                  # the latest snapshot in this view
        milestones     = @($batch.Milestones)
        milestoneCount = @($batch.Milestones).Count
        merged         = [bool]$batch.IsMerged
        finalOnly      = [bool]$FinalOnlyCopy
        generated      = (Get-Date).ToString('yyyy-MM-dd HH:mm')
        total          = $Model.Documents.Count
        final          = @($Model.Documents | Where-Object { $_.Bucket -eq '01 Final' }).Count
        inProgress     = @($Model.Documents | Where-Object { $_.Bucket -eq '02 In Progress' }).Count
        superseded     = @($Model.Documents | Where-Object { $_.Bucket -eq '03 Superseded' }).Count
        excluded       = @($Model.Documents | Where-Object { $_.Excluded }).Count
        # Flags on documents the reviewer already excluded are settled: counting them here would
        # advertise a "Flagged (1)" filter that lists nothing, because the excluded ones live
        # behind their own chip.
        flagged        = @($Model.Documents | Where-Object { $_.ReviewFlag -and -not $_.Excluded }).Count
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
  <div class="ev">$(& $enc $e.Event)</div>
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
#  "About this folder" - the scope statement that travels WITH a -FinalOnly copy
# ----------------------------------------------------------------------------------------
# A curated copy of a trial master file has to say, on its face, that it is curated: who made it,
# from what, what is in it and what was deliberately left out. Without that a recipient cannot tell
# a convenience copy from the archive, and the CRO has no record of what they handed over.
function Write-EtmfScopeNote {
    param([string]$Path, $Batch, [int]$Included, $Withheld, [string]$SourceInput, [string]$ExclusionList)

    $lines = [System.Collections.Generic.List[string]]::new()
    $add = { param([string]$Text = '') $lines.Add($Text) }

    & $add "ABOUT THIS FOLDER"
    & $add "================="
    & $add ""
    & $add "This folder holds the CURRENT, APPROVED documents of the trial master file for"
    & $add ("{0}, taken from the eTMF export named below." -f $Batch.SpecificationId)
    & $add ""
    & $add ("  Source export : {0}" -f (Split-Path $SourceInput -Leaf))
    & $add ("  Milestone     : {0}" -f $Batch.Milestone)
    & $add ("  Prepared      : {0}" -f (Get-Date).ToString('yyyy-MM-dd HH:mm'))
    & $add ("  Prepared with : eTMF Archive Humanizer v{0}" -f $script:ToolVersion)
    if ($ExclusionList) { & $add ("  Exclusion list: {0}" -f (Split-Path $ExclusionList -Leaf)) }
    & $add ""
    & $add "WHAT IS IN HERE"
    & $add ("  {0} document(s) that the eTMF records as the current version AND as finalised" -f $Included)
    & $add "  or locked. They are filed the familiar TMF way: Documents\Zone\Section\Artifact."
    & $add ""
    & $add "WHAT IS NOT IN HERE"
    & $add ("  {0} superseded or replaced version(s)" -f $Withheld.Superseded)
    & $add ("  {0} draft(s) or document(s) still under review" -f $Withheld.InProgress)
    & $add ("  {0} document(s) set aside by the study team as not belonging to this trial" -f $Withheld.Excluded)
    if ($Withheld.NoFile -gt 0) {
        & $add ("  {0} of the included document(s) are listed but their file is not present in the" -f $Withheld.NoFile)
        & $add "  export; they appear in index.html marked "source file not available"."
    }
    & $add ""
    & $add "THIS IS NOT THE COMPLETE TRIAL MASTER FILE."
    & $add "Superseded versions, drafts and their audit trails are retained in full by the"
    & $add "study team. This copy exists so the current approved documents can be read and"
    & $add "searched without the rest."
    & $add ""
    & $add "HOW TO USE IT"
    & $add "  Open index.html in any web browser - no internet needed - to search every"
    & $add "  document and read its details, history and signatures."
    & $add "  inventory.csv lists the same documents as a spreadsheet."
    & $add ""

    [System.IO.File]::WriteAllLines($Path, $lines, (New-Object System.Text.UTF8Encoding($false)))
}

# ----------------------------------------------------------------------------------------
#  Run header - the first thing in every log
# ----------------------------------------------------------------------------------------
# A log is only worth emailing if it says what produced it: which build, which archive, which
# options, on what. Without the version we could not tell whether a reported failure came from a
# build that already contained the fix for it.
function Show-RunHeader {
    param([string]$ArchivePath, $BoundParameters)

    $inputDesc = $ArchivePath
    try {
        $item = Get-Item -LiteralPath $ArchivePath -ErrorAction Stop
        $kind = if ($item.PSIsContainer) { 'folder' } else { '{0:N1} MB' -f ($item.Length / 1MB) }
        $inputDesc = '{0}  ({1}, modified {2:yyyy-MM-dd HH:mm})' -f $item.FullName, $kind, $item.LastWriteTime
    }
    catch { $inputDesc = "$ArchivePath  (not readable)" }

    # Switches print as -Name, everything else as -Name <value>; ArchivePath is on its own line.
    $opts = @()
    if ($null -ne $BoundParameters) {
        foreach ($k in @($BoundParameters.Keys | Sort-Object)) {
            if ($k -eq 'ArchivePath') { continue }
            $v = $BoundParameters[$k]
            if ($v -is [switch] -or $v -is [bool]) { if ($v) { $opts += "-$k" } }
            else { $opts += "-$k `"$v`"" }
        }
    }
    $optText = if ($opts.Count) { $opts -join ' ' } else { '(defaults)' }

    $os = try { [System.Runtime.InteropServices.RuntimeInformation]::OSDescription.Trim() } catch { 'unknown OS' }

    Write-Out ""
    Write-Out "==================== eTMF ARCHIVE HUMANIZER ====================" 'White'
    Write-Out ("  Version    : {0}" -f $script:ToolVersion)
    Write-Out ("  Run at     : {0}" -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz'))
    Write-Out ("  Input      : {0}" -f $inputDesc)
    Write-Out ("  Options    : {0}" -f $optText)
    Write-Out ("  Running on : PowerShell {0} / {1} / machine {2}" -f $PSVersionTable.PSVersion, $os, [Environment]::MachineName)
    Write-Out "===============================================================" 'White'
}

# ----------------------------------------------------------------------------------------
#  Dry-run summary
# ----------------------------------------------------------------------------------------
function Show-Summary {
    param($Model, $Integrity)

    $b = $Model.Batch
    $docs = $Model.Documents

    Write-Out ""
    Write-Out "==================== eTMF ARCHIVE SUMMARY ====================" -ForegroundColor White
    Write-Out ("  Study (Specification) : {0}" -f $b.SpecificationId)
    Write-Out ("  Study ID              : {0}" -f $b.StudyId)
    Write-Out ("  Milestone             : {0}" -f $b.Milestone)
    Write-Out ("  Event / Transfer ID   : {0} / {1}" -f $b.EventId, $b.TransferId)
    Write-Out ("  Source / TMF-RM ver   : {0} / {1}" -f $b.TransferSource, $b.TmfRmVersion)
    Write-Out ("  Manifest              : {0}" -f $b.ManifestFile)
    Write-Out ("  Documents (files)     : {0}" -f $docs.Count)

    Write-Out "`n  -- By status bucket --" -ForegroundColor White
    $docs | Group-Object Bucket | Sort-Object Name | ForEach-Object {
        Write-Out ("     {0,-16} {1,3}" -f $_.Name, $_.Count)
    }
    Write-Out "`n  -- By lifecycle --" -ForegroundColor White
    $docs | Group-Object Lifecycle | Sort-Object Name | ForEach-Object {
        Write-Out ("     {0,-16} {1,3}" -f $_.Name, $_.Count)
    }
    Write-Out "`n  -- By TMF level --" -ForegroundColor White
    $docs | Group-Object ObjectLevel | Sort-Object Name | ForEach-Object {
        Write-Out ("     {0,-16} {1,3}" -f $_.Name, $_.Count)
    }

    $sigCount = @($docs | Where-Object { $_.Signatures.Count -gt 0 }).Count
    Write-Out ("`n  E-signed documents     : {0}" -f $sigCount)

    if ($Integrity) {
        $tot = $Integrity.Ok + $Integrity.Bad + $Integrity.Missing
        $col = if ($Integrity.Bad -eq 0 -and $Integrity.Missing -eq 0) { 'Green' } else { 'Yellow' }
        Write-Out ("  Integrity (MD5)        : {0}/{1} OK, {2} mismatch, {3} missing" -f `
            $Integrity.Ok, $tot, $Integrity.Bad, $Integrity.Missing) -ForegroundColor $col
    } else {
        Write-Out "  Integrity (MD5)        : skipped" -ForegroundColor DarkGray
    }

    Write-Out "`n  -- Document inventory (preview) --" -ForegroundColor White
    $docs |
        Sort-Object ArtifactNumber, Bucket, Country, OriginalName |
        Select-Object @{n='Art#';e={$_.ArtifactNumber}},
                      @{n='Bucket';e={$_.Bucket}},
                      @{n='Status';e={$_.Status}},
                      @{n='Ver';e={$_.Version}},
                      @{n='Ctry';e={$_.Country}},
                      @{n='Int';e={ if($null -eq $_.IntegrityOk){'-'} elseif($_.IntegrityOk){'ok'} else {'X'} }},
                      @{n='Original file';e={ if($_.OriginalName.Length -gt 42){$_.OriginalName.Substring(0,39)+'...'}else{$_.OriginalName} }} |
        Format-Table -AutoSize | Out-String -Width 160 | ForEach-Object { Write-Out $_ }
    Write-Out "==============================================================" -ForegroundColor White
}

function Show-Recap {
    param([string]$OutputRoot, $Written, [string]$Landing, $Gaps, [int]$Demoted = 0, [int]$Flagged = 0, [int]$Excluded = 0, $Withheld = $null)
    Write-Out ""
    Write-Out "==================== DONE ====================" -ForegroundColor White
    Write-Out ("  Output folder : {0}" -f $OutputRoot) -ForegroundColor Green
    if ($null -ne $Withheld) {
        Write-Out  "  Contents      : CURRENT APPROVED DOCUMENTS ONLY - a copy to hand over," -ForegroundColor Cyan
        Write-Out  "                  not the complete trial master file." -ForegroundColor Cyan
        Write-Out ("                  Left out: {0} superseded, {1} in progress, {2} excluded." -f `
            $Withheld.Superseded, $Withheld.InProgress, $Withheld.Excluded) -ForegroundColor Cyan
        Write-Out  "                  'About this folder.txt' states this for the recipient." -ForegroundColor Cyan
    }
    foreach ($e in $Written) {
        $flag = if ($e.Missing -gt 0 -or $e.CopyBad -gt 0) { '  [!]' } else { '' }
        $col  = if ($flag) { 'Yellow' } else { 'Gray' }
        # The counts must add up to the total, or the reader is left hunting for the difference.
        $exc = if ($e.Excluded -gt 0) { ' / {0} excluded' -f $e.Excluded } else { '' }
        Write-Out ("    {0,-30} {1,3} docs  ({2} final / {3} in-progress / {4} superseded{5})  copied {6}{7}" -f `
            $e.Event, $e.Docs, $e.Final, $e.InProgress, $e.Superseded, $exc, $e.Copied, $flag) -ForegroundColor $col
        if ($e.Missing -gt 0) { Write-Out ("           {0} document(s) missing source files" -f $e.Missing) -ForegroundColor Yellow }
        if ($e.CopyBad -gt 0) { Write-Out ("           {0} copy integrity mismatch(es)" -f $e.CopyBad) -ForegroundColor Yellow }
    }
    # The headline answer to "which one is the current approved version?".
    if ($Demoted -gt 0) {
        Write-Out ""
        Write-Out ("  {0} older version(s) were filed under '03 Superseded' because a newer" -f $Demoted) -ForegroundColor Cyan
        Write-Out  "  version of the same document exists in the archive. inventory.csv's" -ForegroundColor Cyan
        Write-Out  "  SupersededBy column names the version that replaced each one." -ForegroundColor Cyan
    }

    if ($Excluded -gt 0) {
        Write-Out ""
        Write-Out ("  {0} document(s) from your exclusion list were filed under '_Excluded'" -f $Excluded) -ForegroundColor Cyan
        Write-Out  "  (kept, with the reason, but outside the TMF tree)." -ForegroundColor Cyan
    }
    if ($Flagged -gt 0) {
        Write-Out ""
        Write-Out ("  {0} document(s) are FLAGGED FOR REVIEW and still in the TMF tree; the" -f $Flagged) -ForegroundColor Yellow
        Write-Out  "  tool cannot know whether they belong here. See inventory.csv's ReviewFlag" -ForegroundColor Yellow
        Write-Out  "  column (or the viewer's 'Flagged' filter), then list the ones that do not" -ForegroundColor Yellow
        Write-Out  "  belong in a CSV and re-run with -ExclusionList to file them separately." -ForegroundColor Yellow
    }

    # Group the gaps by CAUSE. "12 missing" prompts a mail asking why; "12 - the archive contains no
    # entry with this exact path" is already the answer, and points at the export rather than at us.
    if ($Gaps -and @($Gaps).Count -gt 0) {
        Write-Out ""
        Write-Out "  Missing source files, by cause:" -ForegroundColor Yellow
        foreach ($g in $Gaps) {
            Write-Out ("    {0,4}  {1}" -f $g.Count, $g.Reason) -ForegroundColor Yellow
        }
        Write-Out "          Each one is listed in inventory.csv (Note column) with its" -ForegroundColor Yellow
        Write-Out "          CONTENTURL, and in the run log with the exact path that was looked for." -ForegroundColor Yellow
    }
    $open = if ($Landing) { $Landing } else { $Written[0].Index }
    Write-Out ""
    Write-Out ("  Open viewer   : {0}" -f $open) -ForegroundColor Green
    Write-Out "=============================================" -ForegroundColor White
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
  .chip.on[data-s="04 Excluded"]{background:#8a4b4b}
  .b-exc{color:#8a4b4b;background:#f7e9e9}
  .flagchip{border-color:#e0c98a;color:#8a6d1e;background:#fdf8e8}
  .flagchip.on{background:#9a6700;color:#fff;border-color:transparent}
  .flag{color:#9a6700}
  .flagbox{background:#fdf8e8;border:1px solid #ecdfae;color:#6b5a1e;border-radius:8px;padding:9px 11px;margin:0 0 12px;font-size:12.5px;line-height:1.45}
  .excbox{background:#f7e9e9;border:1px solid #e6cccc;color:#7d4444;border-radius:8px;padding:9px 11px;margin:0 0 12px;font-size:12.5px;line-height:1.45}
  .resultCount{margin-left:auto;color:var(--muted);font-size:12px;white-space:nowrap}
  [hidden]{display:none !important}
  .filterbar{display:flex;align-items:center;gap:8px;flex-wrap:wrap;padding:8px 16px;background:#fdf8e8;border-bottom:1px solid #ecdfae;font-size:12.5px;color:#6b5a1e}
  .fb-label{font-weight:600}
  .fchip{background:#fff;border:1px solid #e5d9ae;border-radius:20px;padding:2px 10px;white-space:nowrap}
  #filterList{display:flex;flex-wrap:wrap;gap:6px;min-width:0}
  .fb-show{border:1px solid #d9c78a;background:#fff;color:#6b5a1e;border-radius:5px;padding:0 6px;margin-left:6px;cursor:pointer;font-size:11.5px}
  .vh{overflow-wrap:anywhere}
  .vh .vh-none{display:flex;align-items:center;gap:8px;padding:6px 9px;border:1px dashed var(--line);border-radius:7px;color:var(--muted);font-size:12.5px}
  .fb-clear{flex:0 0 auto;margin-left:auto;border:1px solid #d9c78a;background:#fff;color:#6b5a1e;border-radius:7px;padding:4px 10px;cursor:pointer;font-size:12.5px}
  .fb-clear:hover{background:#fffaf0}
  #docTable.no-ms .col-ms{display:none}
  td.ms{color:var(--muted);white-space:nowrap;max-width:170px;overflow:hidden;text-overflow:ellipsis}
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
  .gapwhy{background:#fdf3f3;border:1px solid #f0d5d5;color:#8a4b4b;border-radius:8px;padding:9px 11px;margin:0 0 14px;font-size:12.5px;line-height:1.45}
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
      <input id="search" type="search" placeholder="Search title, document, artifact #, zone, site, status, ID…" autocomplete="off">
      <div class="chips" id="statusChips"></div>
      <select id="countryFilter"><option value="all">All countries</option></select>
      <select id="milestoneFilter" hidden><option value="all">All milestones</option></select>
      <button type="button" class="chip flagchip" id="flagToggle" hidden></button>
      <span class="resultCount" id="resultCount"></span>
    </div>
    <!-- Shown only while something is filtered out: a reviewer who cannot find a document must be
         able to see that the view is narrowed, and undo it in one click. -->
    <div class="filterbar" id="filterBar" hidden>
      <span class="fb-label">Showing a filtered view:</span>
      <span id="filterList"></span>
      <button type="button" class="fb-clear" id="clearFilters">Clear all filters</button>
    </div>
    <div class="tablewrap">
      <table id="docTable">
        <thead><tr>
          <th data-key="artNo">Artifact #</th>
          <th data-key="title">Title</th>
          <th data-key="originalName">Document</th>
          <th data-key="scope">Country / Site</th>
          <th class="col-ms" data-key="milestone" title="The most recent milestone export that still contained this document">Last export</th>
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
// Encode each path SEGMENT: encodeURI() deliberately leaves '#' (and '?', '&') alone, so a document
// called "Protocol Amendment #2.pdf" - a name Windows allows and our sanitizer therefore keeps -
// produced a link the browser cut at the '#'. The file was on disk; only the link was broken.
const href = p => p ? p.split('/').map(encodeURIComponent).join('/') : '#';
// Documents open in a new tab so the viewer stays put.
const linkTgt = ' target="_blank" rel="noopener noreferrer"';
const fmtSize = b => { if(!b) return ''; const u=['B','KB','MB','GB']; let i=0,n=b; while(n>=1024&&i<u.length-1){n/=1024;i++;} return (n<10&&i>0?n.toFixed(1):Math.round(n))+' '+u[i]; };
const fmtTs = s => { if(!s) return ''; const d=new Date(s); return isNaN(d)? s : d.toISOString().slice(0,16).replace('T',' ')+' UTC'; };
// Same rule as Compare-EtmfVersion on the PowerShell side: segment-wise, so v10 > v9 and
// v1.10 > v1.9. Parsing "1.10" as a number makes it 1.1 and sorts it below 1.9.
const cmpVersion = (a,b) => {
  const sa=String(a==null?'':a).trim().split('.'), sb=String(b==null?'':b).trim().split('.');
  const num = s => s.length>0 && s.every(x=>/^\d+$/.test(x));
  if(num(sa)&&num(sb)){
    for(let i=0;i<Math.max(sa.length,sb.length);i++){
      const x=+(sa[i]||0), y=+(sb[i]||0);
      if(x!==y) return x<y?-1:1;
    }
    return 0;
  }
  const x=String(a==null?'':a), y=String(b==null?'':b);
  return x<y?-1:x>y?1:0;
};
const bClass = b => b==='01 Final'?'b-final':b==='02 In Progress'?'b-prog':b==='04 Excluded'?'b-exc':'b-sup';

const state = { search:'', searchRaw:'', status:'all', country:'all', milestone:'all', flagged:false, node:null, sortKey:'artNo', sortDir:1 };

// One lower-cased haystack per document, built once. Searching a TMF means searching for whatever
// the reviewer happens to remember - an artifact number, a zone, a site, a status word, a document
// ID, the milestone - so every field a reader can see is in here, and the per-keystroke filter is
// then a single string test rather than a dozen concatenations per row.
// PowerShell's ConvertTo-Json emits a one-element array as a bare value on some hosts, so never
// assume this field is an array.
const msList = d => !d.milestones ? [] : (Array.isArray(d.milestones) ? d.milestones : [d.milestones]);
DATA.forEach(d => { d._hay = [d.artNo, d.title, d.originalName, d.newName, d.artifact, d.zone,
  d.section, d.scope, d.country, d.level, d.person, d.version, d.status, d.lifecycle,
  d.versionState, d.documentId, msList(d).join(' '), d.date, d.reviewFlag, d.supersededBy,
  d.exclusionReason]
  .filter(Boolean).join(' ').toLowerCase(); });

// ---- header ----
$('#studyName').textContent = META.study || META.studyId || 'eTMF Archive';
// A merged view spans every milestone export in the archive, so naming a single event would be a
// lie; name the span and the latest snapshot instead. (Per-milestone filtering: see #2.)
const scopeTxt = META.merged
  ? `<b>${esc(META.milestoneCount)}</b> milestones merged · latest <b>${esc(META.milestone)}</b>`
  : `Milestone <b>${esc(META.milestone)}</b> · Transfer <b>${esc(META.transferId)}</b>`;
$('#studySub').innerHTML =
  `${scopeTxt} · TMF-RM <b>${esc(META.tmfrm)}</b> · Source <b>${esc(META.source)}</b>` +
  (META.finalOnly ? ` · <b>current approved documents only</b>` : '');
const intTxt = META.integrityChecked ? `${META.integrityOk}/${META.integrityTotal}` : 'n/a';
const countDefs = [['Total',META.total]];
if (!META.finalOnly) countDefs.push(['Final',META.final],['In&nbsp;Progress',META.inProgress],['Superseded',META.superseded]);
countDefs.push(['E-signed',META.signed],['Integrity',intTxt]);
if (META.missing)  countDefs.push(['Missing',META.missing]);
if (META.flagged)  countDefs.push(['Flagged',META.flagged]);
if (META.excluded) countDefs.push(['Excluded',META.excluded]);
$('#counts').innerHTML = countDefs
  .map(([l,n])=>`<div class="count"><div class="n">${n}</div><div class="l">${l}</div></div>`).join('');
$('#footer').innerHTML = `Generated ${esc(META.generated)} from <b>${esc(META.manifest)}</b> · ${META.total} documents · eTMF Archive Humanizer <b>v${esc(META.toolVersion)}</b>`;

// ---- status chips ----
// "All" means all of the TMF - excluded documents are reached through their own chip, so the folder
// a client is handed and the list they browse agree with each other. They are never hidden outright:
// the header carries an Excluded count whenever there are any.
// Only offer a filter for a bucket this copy actually contains: in a -FinalOnly copy an "In
// Progress" chip that always lists nothing invites the reader to wonder what is being hidden.
const bucketsPresent = new Set(DATA.map(d=>d.bucket));
const chips = [['all','All']].concat(
  [['01 Final','Final'],['02 In Progress','In Progress'],['03 Superseded','Superseded']]
    .filter(([b])=>bucketsPresent.has(b)));
if (META.excluded) chips.push(['04 Excluded','Excluded']);
// One bucket and nothing excluded (a -FinalOnly copy): "All" and "Final" would be the same filter.
if (chips.length < 3) { $('#statusChips').hidden = true; }
$('#statusChips').innerHTML = chips.map(([s,l])=>`<span class="chip${s==='all'?' on':''}" data-s="${s}">${l}</span>`).join('');
$('#statusChips').addEventListener('click',e=>{ const c=e.target.closest('.chip'); if(!c) return;
  state.status=c.dataset.s;
  // Scoped: #flagToggle is also a .chip, and a document-wide selector de-highlighted it while its
  // filter stayed on - leaving the control showing the opposite of its state.
  document.querySelectorAll('#statusChips .chip').forEach(x=>x.classList.toggle('on',x===c)); render(); });

// ---- milestone filter ----
// Filtering on the record's OWN milestone would be near-useless: after the merge that is the newest
// snapshot for almost every document. The question a reviewer actually asks is "what did the TMF
// hold at milestone N", so this filters on MEMBERSHIP - was the document in that export at all -
// and the options come from the archive's milestone list, not from the surviving records.
const allMilestones = Array.isArray(META.milestones) ? META.milestones : [META.milestones];
if (META.milestoneCount > 1) {
  $('#milestoneFilter').innerHTML = '<option value="all">All milestones</option>' +
    allMilestones.map(m=>`<option value="${esc(m)}">${esc(m)}</option>`).join('');
  $('#milestoneFilter').hidden = false;
} else {
  $('#docTable').classList.add('no-ms');   // one snapshot: the column would repeat on every row
}
$('#milestoneFilter').addEventListener('change',e=>{ state.milestone=e.target.value; render(); });

// ---- country filter ----
const countries = [...new Set(DATA.map(d=>d.country).filter(Boolean))].sort();
// value= is required: option.value falls back to option.text, which the DOM returns with internal
// whitespace collapsed - so any label containing a double space would filter to nothing.
$('#countryFilter').innerHTML = '<option value="all">All countries</option>' + countries.map(c=>`<option value="${esc(c)}">${esc(c)}</option>`).join('');
$('#countryFilter').addEventListener('change',e=>{ state.country=e.target.value; render(); });

// ---- flagged-for-review toggle ----
if (META.flagged) {
  const t = $('#flagToggle');
  t.hidden = false;
  t.textContent = `\u2691 Flagged for review (${META.flagged})`;
  t.addEventListener('click',()=>{ state.flagged=!state.flagged; t.classList.toggle('on',state.flagged); render(); });
}

// ---- search ----
$('#search').addEventListener('input',e=>{ state.searchRaw=e.target.value; state.search=e.target.value.toLowerCase(); render(); });

// ---- tree ----
function buildTree(){
  const t={};
  DATA.filter(d=>!d.excluded).forEach(d=>{
    const z=t[d.zone]??=({c:0,f:0,s:{}}); z.c++; if(d.bucket!=='01 Final')z.f++;
    const se=z.s[d.section]??=({c:0,f:0,a:{}}); se.c++; if(d.bucket!=='01 Final')se.f++;
    const a=se.a[d.artifact]??=({c:0,f:0}); a.c++; if(d.bucket!=='01 Final')a.f++;
  });
  const dot=o=>`<span class="dot" style="background:${o.f?'var(--prog)':'var(--final)'}"></span>`;
  let h='<ul><li><div class="node active" data-type="all">All documents<span class="cnt">'+DATA.filter(d=>!d.excluded).length+'</span></div></li>';
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
// The user's own filters, with the exclusion rule left out - so render() can tell the difference
// between "your filter hid it" and "it is excluded from the TMF", which are not the same message.
function matchesCore(d){
  if(state.status!=='all' && d.bucket!==state.status) return false;
  if(state.flagged && !d.reviewFlag) return false;
  if(state.country!=='all' && d.country!==state.country) return false;
  if(state.milestone!=='all' && msList(d).indexOf(state.milestone)<0) return false;
  if(state.node){ if(state.node.type==='zone'&&d.zone!==state.node.v) return false;
    if(state.node.type==='section'&&d.section!==state.node.v) return false;
    if(state.node.type==='artifact'&&d.artifact!==state.node.v) return false; }
  if(state.search && !d._hay.includes(state.search)) return false;
  return true;
}
function matches(d){
  // Excluded documents are reachable only under their own chip: the client-facing tree is the
  // product, so hiding them is the default - but never SILENTLY (see renderFilterBar).
  if(d.excluded && state.status!=='04 Excluded') return false;
  return matchesCore(d);
}

// ---- active filters: say what is being hidden, and offer one click to undo it ----
const chipLabel = s => (chips.find(c=>c[0]===s)||[,s])[1];
function activeFilters(){
  const a=[];
  if(state.searchRaw)          a.push(['Search', '\u201c'+state.searchRaw+'\u201d']);
  if(state.status!=='all')     a.push(['Status', chipLabel(state.status)]);
  if(state.country!=='all')    a.push(['Country', state.country]);
  if(state.milestone!=='all')  a.push(['In milestone', state.milestone]);
  if(state.flagged)            a.push(['Only', 'flagged for review']);
  if(state.node)               a.push([state.node.type.charAt(0).toUpperCase()+state.node.type.slice(1), state.node.v]);
  return a;
}
// How many EXCLUDED documents the current filters would otherwise have shown. Only then is the
// exclusion worth mentioning: a permanent notice would turn the bar into furniture and destroy its
// one-glance "something is narrowed" signal, but staying silent produced a search that dead-ends on
// a document the archive does contain - and, when everything is excluded, a viewer that just looks
// broken.
function hiddenExcludedCount(){
  if(state.status==='04 Excluded') return 0;
  return DATA.filter(d=>d.excluded && matchesCore(d)).length;
}
function renderFilterBar(){
  const a=activeFilters();
  const hidden=hiddenExcludedCount();
  $('#filterBar').hidden = (a.length===0 && hidden===0);
  const chips=a.map(([k,v])=>`<span class="fchip">${esc(k)}: ${esc(v)}</span>`);
  if(hidden>0) chips.push(`<span class="fchip">${hidden} excluded from the TMF, not shown <button type="button" class="fb-show" id="showExcluded">show</button></span>`);
  $('#filterList').innerHTML = chips.join('');
  const btn=$('#showExcluded');
  if(btn) btn.addEventListener('click',()=>{
    const chip=document.querySelector('.chip[data-s="04 Excluded"]');
    if(chip) chip.click();
  });
}
$('#clearFilters').addEventListener('click',()=>{
  state.search=''; state.status='all'; state.country='all'; state.milestone='all'; state.node=null;
  state.flagged=false; const ft=$('#flagToggle'); if(ft) ft.classList.remove('on');
  $('#search').value=''; $('#countryFilter').value='all'; $('#milestoneFilter').value='all';
  document.querySelectorAll('.chip').forEach(c=>c.classList.toggle('on', c.dataset.s==='all'));
  const all=document.querySelector('#tree .node[data-type="all"]');
  document.querySelectorAll('#tree .node').forEach(n=>n.classList.toggle('active', n===all));
  render();
});
function render(){
  let rows=DATA.filter(matches);
  const k=state.sortKey, dir=state.sortDir;
  rows.sort((a,b)=>{ let x=a[k],y=b[k];
    if(k==='version'){ const c=cmpVersion(a.version,b.version); return c*dir; }
    else if(k==='date'){x=a.dateIso||'';y=b.dateIso||'';}        // sort by ISO date, not the display string
    else {x=(''+x).toLowerCase();y=(''+y).toLowerCase();}
    return x<y?-1*dir:x>y?dir:0; });
  const tb=$('#tbody');
  tb.innerHTML=rows.map((d,i)=>{
    const idx=DATA.indexOf(d);
    const sig=d.signatures&&d.signatures.length?' <span class="sig" title="Electronically signed">✎</span>':'';
    const restr=d.restricted?' <span title="Restricted artifact" style="color:#b23b3b">🔒</span>':'';
    const flag=d.reviewFlag?` <span class="flag" title="${esc(d.reviewReason)}">\u2691</span>`:'';
    const integ = !d.present ? '<span class="no" title="Source file missing">⚠</span>'
                : d.copyOk ? '<span class="ok" title="MD5 verified">✓</span>'
                : (META.integrityChecked?'<span class="no" title="Integrity mismatch">✗</span>':'');
    return `<tr data-i="${idx}">
      <td class="art">${esc(d.artNo)}</td>
      <td>${esc(d.title)}${sig}${restr}${flag}</td>
      <td class="doc" title="${esc(d.originalName)}">${esc(d.originalName)}</td>
      <td>${esc(d.scope||(d.level==='Trial'?'—':d.country))}</td>
      <td class="ms col-ms" title="Exported in: ${esc(msList(d).join('; '))}">${esc(d.milestone)}</td>
      <td>${esc(d.version)}</td>
      <td><span class="badge ${bClass(d.bucket)}">${esc(d.status)}</span></td>
      <td>${esc(d.date)}</td>
      <td>${integ}</td></tr>`;
  }).join('');
  $('#empty').style.display = rows.length?'none':'block';
  // Count against the reachable set, not DATA: "6 of 7" while All shows 6 sends a reader hunting
  // for a document the view was never going to list.
  const reachable = DATA.filter(d=> state.status==='04 Excluded' ? d.excluded : !d.excluded).length;
  $('#resultCount').textContent = rows.length+' of '+reachable+' documents';
  renderFilterBar();
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
    .sort((a,b)=>cmpVersion(a.x.version,b.x.version));
  const vh = versions.length>1 ? `<div class="sec">Version history</div><div class="vh">`+
    versions.map(o=>{
      const label=`<span class="badge ${bClass(o.x.bucket)}">${esc(o.x.status)}</span> v${esc(o.x.version)} <span style="color:var(--muted)">${esc(o.x.originalName)}</span>`;
      // href="#" with target=_blank opened a second copy of the viewer; a version with no file is
      // not a link at all.
      return (o.x.present && o.x.relPath)
        ? `<a class="${o.ix===i?'cur':''}" href="${href(o.x.relPath)}"${linkTgt}>${label}</a>`
        : `<span class="vh-none ${o.ix===i?'cur':''}" title="Source file not available">${label} \u26a0</span>`;
    }).join('')+`</div>` : '';
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
      ${d.excluded?`<div class="excbox"><b>Excluded from the TMF tree.</b><br>${esc(d.exclusionReason)}<br>
         <span style="color:var(--muted)">Filed under _Excluded, not deleted. This came from the exclusion list you supplied.</span></div>`:''}
      ${d.reviewFlag?`<div class="flagbox"><b>\u2691 Flagged for review: ${esc(d.reviewFlag)}</b><br>${esc(d.reviewReason)}<br>
         <span style="color:var(--muted)">A flag is a prompt to look, not a verdict - the tool cannot know whether this document belongs here.</span></div>`:''}
      ${d.present && d.relPath ? `<a class="openbtn" href="${href(d.relPath)}"${linkTgt}>Open document ↗</a>`
        : `<div class="openbtn" style="background:var(--sup);cursor:default">⚠ Source file not available</div>
           ${d.gapReason?`<div class="gapwhy">Why: ${esc(d.gapReason)}.<br>The record is kept so the gap is visible; the file itself is not in the export.</div>`:''}`}
      <div class="sec">Document</div>
      <div class="kv">
        ${kv('Artifact',d.artifact)}${kv('Sub-artifact / title',d.title)}${kv('TMF level',d.level)}
        ${kv('Country',d.country)}${kv('Site / org',d.scope)}${kv('Person',d.person)}
        ${kv('Version',d.version)}${kv('Version state',d.versionState)}${kv('Lifecycle',d.lifecycle)}
        ${kv('Superseded by',d.supersededBy)}
        ${kv('Restricted',d.restricted?'Yes':'')}${kv('Language',d.language)}${kv('Translation',d.translation)}
        ${kv(d.dateDesc||'Artifact date',d.date)}${kv('Expires',d.expiry)}${kv('Document ID',d.documentId)}
        ${kv('Last export',d.milestone)}${META.merged?kv('Exported in',msList(d).join('; ')):''}
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
$script:LogFallbackDir = $null
$script:TempDir    = $null      # set as soon as the temp dir exists, so cleanup survives a throw
$script:Demoted    = 0          # older versions filed under 03 Superseded by version resolution
$script:Flagged    = 0          # records raised for a human to look at
$script:Excluded   = 0          # records the reviewer's exclusion list took out of the TMF tree
$source = $null

# No archive supplied -> show friendly instructions and stop (instead of prompting).
if ([string]::IsNullOrWhiteSpace($ArchivePath)) {
    Show-Usage
    return
}

# Where the run log goes if the run dies before there is an output folder - a crash log is the one
# we most need. Set only once the input is known to exist: a run that fails because the path was
# wrong has nothing in its log the user did not just read on screen, and writing it would drop a
# file into whatever directory they happened to be standing in.

try {
    Show-RunHeader -ArchivePath $ArchivePath -BoundParameters $PSBoundParameters

    # Resolve where the output will live (sibling of the input by default).
    $inputItem = Get-Item -LiteralPath $ArchivePath
    # Beside the input, never INSIDE it: for folder input the archive itself is what the user handed
    # us, and the tool promises not to modify it.
    $script:LogFallbackDir = if ($inputItem.PSIsContainer) {
                                 if ($inputItem.Parent) { $inputItem.Parent.FullName } else { $inputItem.FullName }
                             } else { $inputItem.DirectoryName }
    $baseDir   = if ($inputItem.PSIsContainer) {
                     if ($inputItem.Parent) { $inputItem.Parent.FullName } else { $inputItem.FullName }  # drive/fs root has no parent
                 } else { $inputItem.DirectoryName }

    $source = Resolve-EtmfSource -ArchivePath $ArchivePath
    # For a zip the extractor already told us where the manifests are (and what archive prefix each
    # one's CONTENTURLs are relative to); for a folder we go and find them.
    if ($null -ne $source.Manifests) {
        $manifests = @($source.Manifests | ForEach-Object {
            [pscustomobject]@{ File = (Get-Item -LiteralPath $_.Path); Prefix = $_.Prefix }
        })
        if ($manifests.Count -eq 0) { throw "No '*exchange.xml' manifest found in: $ArchivePath" }
    }
    else {
        $manifests = @(Find-ExchangeManifests -Root $source.Root |
            ForEach-Object { [pscustomobject]@{ File = $_; Prefix = '' } })
    }
    Write-Log ("Found {0} manifest file(s)." -f $manifests.Count) 'OK'

    # ---- Phase 4: parse + classify + verify every manifest (one per study event/milestone) ----
    $events = foreach ($m in $manifests) {
        $mf = $m.File
        Write-Log "Parsing manifest: $($mf.Name)" 'STEP'
        $model = Import-EtmfManifest -ManifestFile $mf -ContentMap $source.ContentMap -ContentPrefix $m.Prefix
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

    # Read and validate the exclusion list up front: it must fail fast on a bad path or a missing
    # column whichever output mode is chosen, and it is applied to every model below.
    $exclusions = @()
    if ($ExclusionList) {
        $exclusions = Import-EtmfExclusions -Path $ExclusionList
        Write-Log ("Exclusion list: {0} row(s) read from {1}" -f @($exclusions).Count, $ExclusionList) 'INFO'
    }

    if ($DryRun) {
        # Preview the decisions, not just the parse: a reviewer uses -DryRun to see what WOULD be
        # merged, superseded, flagged and excluded before committing to an output folder.
        if (-not $SplitByMilestone -or $events.Count -le 1) {
            $preview = Merge-EtmfEvents -Events $events
            $pv = Resolve-EtmfVersions -Model $preview
            $pf = Set-EtmfReviewFlags -Model $preview
            $px = if (@($exclusions).Count -gt 0) { Set-EtmfExclusions -Model $preview -Exclusions $exclusions } else { 0 }
            Write-Log ("DryRun: {0} document(s) after merge ({1} duplicate(s) collapsed); {2} older version(s) would be superseded; {3} flagged for review; {4} excluded." -f `
                $preview.Documents.Count, $preview.Collapsed, $pv.Demoted, $pf, $px) 'INFO'
        }
        Write-Log "DryRun: nothing was written." 'INFO'
        return
    }

    # ---- output root (computed once from the study) ----
    $b0 = $events[0].Model.Batch
    $studyName = ConvertTo-SafeName $b0.SpecificationId
    if ([string]::IsNullOrWhiteSpace($studyName)) { $studyName = ConvertTo-SafeName $b0.StudyId }
    # A final-only copy gets its own default folder name: it must never be mistaken for, or
    # overwrite, the complete rendering of the same study.
    $nameTemplate = if ($FinalOnly) { $script:FinalOnlyOutputNameTemplate } else { $script:DefaultOutputNameTemplate }
    $outputRoot = if ($OutputPath) { $OutputPath }
                  else { Join-Path $baseDir (ConvertTo-SafeName ($nameTemplate -f $studyName)) }
    $script:OutputRoot = $outputRoot

    if (Test-Path -LiteralPath $outputRoot) {
        if ($Force) {
            # Safety: only ever auto-delete a directory that looks like our own prior output,
            # so a mistyped -OutputPath + -Force cannot wipe an unrelated folder.
            # The run log is dated (_log_<stamp>.txt), so match it by prefix - a run that died after
            # creating the output root leaves only that file, and an exact '_log.txt' test would
            # refuse the very re-run the user needs.
            $looksLikeOurs = @(
                @('index.html', 'inventory.csv', 'Documents', '_Excluded') |
                    Where-Object { Test-Path -LiteralPath (Join-Path $outputRoot $_) }
                @(Get-ChildItem -LiteralPath $outputRoot -Filter '_log_*.txt' -File -ErrorAction SilentlyContinue)
            )
            if (-not $looksLikeOurs) {
                throw "Refusing to -Force delete '$outputRoot': it does not look like a previous run's output " +
                      "(no index.html / inventory.csv / _log_<date>.txt / Documents). Remove it manually or choose another -OutputPath."
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
    Write-Log "Output folder: $outputRoot" 'INFO'

    # ---- Phases 5 + 6: build tree + inventory + viewer ----
    # DEFAULT: ONE merged tree across every milestone export. A milestone export is a snapshot of
    # the whole TMF, so a folder per milestone hands the reader N partial TMFs, N indexes and no
    # way to search across them - which is exactly the state the archive was already in.
    # -SplitByMilestone restores the per-snapshot view for anyone who wants it.
    $written = [System.Collections.Generic.List[object]]::new()
    $split   = $SplitByMilestone -and $events.Count -gt 1
    $plan    = [System.Collections.Generic.List[object]]::new()

    if ($split) {
        $usedLabels = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($ev in $events) {
            $label = ConvertTo-SafeName $ev.Model.Batch.Milestone
            if (-not $usedLabels.Add($label)) {                       # two manifests, one label
                $n = 2
                while (-not $usedLabels.Add("$label ($n)")) { $n++ }
                $label = "$label ($n)"
                Write-Log "Duplicate milestone label; disambiguated to '$label'" 'WARN'
            }
            $eventDir = Join-Path $outputRoot $label
            New-Item -ItemType Directory -Path $eventDir -Force | Out-Null
            $plan.Add([pscustomobject]@{ Model = $ev.Model; Integrity = $ev.Integrity; Dir = $eventDir; Label = $label; Folder = $label })
        }
    }
    else {
        $merged = Merge-EtmfEvents -Events $events
        if ($events.Count -gt 1) {
            Write-Log ("Merged {0} milestone export(s) into one TMF: {1} document(s) after collapsing {2} duplicate(s)." -f `
                $events.Count, $merged.Documents.Count, $merged.Collapsed) 'OK'
        }
        elseif ($merged.Collapsed -gt 0) {
            # One manifest can still list the same file twice under two OBJECTs.
            Write-Log ("Collapsed {0} duplicate record(s) pointing at the same file." -f $merged.Collapsed) 'INFO'
        }
        # Now that every snapshot is in one model, decide which VERSION of each document is the
        # current one - the per-snapshot OBJECTVERSIONSTATE cannot know about a later replacement.
        $resolution = Resolve-EtmfVersions -Model $merged
        if ($resolution.Demoted -gt 0) {
            Write-Log ("Version resolution: {0} older version(s) across {1} document(s) moved to '03 Superseded' - a newer version exists in the archive." -f `
                $resolution.Demoted, $resolution.Groups) 'OK'
        }

        # Integrity was verified per manifest; restate it over the merged set so the viewer's
        # header counts what this view actually holds rather than one snapshot's worth.
        $mergedIntegrity = $null
        if (-not $SkipIntegrityCheck) {
            $okCount      = @($merged.Documents | Where-Object { $_.IntegrityOk -eq $true }).Count
            $missingCount = @($merged.Documents | Where-Object { -not $_.FileExists }).Count
            $mergedIntegrity = [pscustomobject]@{
                Ok      = $okCount
                Bad     = ($merged.Documents.Count - $okCount - $missingCount)
                Missing = $missingCount
            }
        }
        $label = if ($merged.Batch.IsMerged) { "All {0} milestones (merged)" -f @($merged.Batch.Milestones).Count }
                 else                        { $merged.Batch.Milestone }
        $plan.Add([pscustomobject]@{ Model = $merged; Integrity = $mergedIntegrity; Dir = $outputRoot; Label = $label; Folder = '' })
        $script:Demoted = $resolution.Demoted
    }

    # Flag what a person should look at (nothing moves), then apply the decisions a person has
    # already made. BOTH run for every model, in either output mode: a reviewer's written decision
    # that a document does not belong in this TMF must not depend on which layout was chosen - it
    # used to be dropped in silence under -SplitByMilestone, publishing the document to the client.
    foreach ($p in $plan) {
        $script:Flagged += Set-EtmfReviewFlags -Model $p.Model
        if (@($exclusions).Count -gt 0) {
            $script:Excluded += Set-EtmfExclusions -Model $p.Model -Exclusions $exclusions
        }
    }
    if (@($exclusions).Count -gt 0) {
        foreach ($p in $plan) { Update-EtmfSupersededReferences -Model $p.Model }
        Write-EtmfExclusionOutcome -Exclusions $exclusions
        Write-Log ("{0} document(s) excluded from the TMF tree (filed under _Excluded)." -f $script:Excluded) 'OK'
    }
    # A flag on a document the reviewer has already excluded is settled: count only what is still
    # in the tree and still awaiting a decision.
    $script:Flagged = @($plan | ForEach-Object { $_.Model.Documents } |
                        Where-Object { $_.ReviewFlag -and -not $_.Excluded }).Count
    if ($script:Flagged -gt 0) {
        Write-Log ("{0} document(s) flagged for review (duplicate content and/or a name mentioning another study's code)." -f $script:Flagged) 'WARN'
    }

    # -FinalOnly: reduce every model to the current approved documents, and remember what that
    # left out so the copy can state its own scope.
    $withheld = $null
    if ($FinalOnly) {
        foreach ($p in $plan) {
            $w = Select-EtmfFinalDocuments -Model $p.Model
            if ($null -eq $withheld) { $withheld = $w }
            else {
                foreach ($f in @('Superseded', 'InProgress', 'Excluded', 'NoFile', 'Flagged')) {
                    $withheld.$f = $withheld.$f + $w.$f
                }
            }
        }
        $keptTotal = @($plan | ForEach-Object { $_.Model.Documents }).Count
        Write-Log ("-FinalOnly: keeping {0} current approved document(s); leaving out {1} superseded, {2} in progress, {3} excluded." -f `
            $keptTotal, $withheld.Superseded, $withheld.InProgress, $withheld.Excluded) 'OK'
        if ($withheld.Flagged -gt 0) {
            Write-Log ("{0} document(s) in this copy are flagged for review - they are approved in the eTMF, but the tool could not tell whether they belong to this study. Check them before handing the folder over." -f $withheld.Flagged) 'WARN'
        }
    }

    foreach ($p in $plan) {
        $model = $p.Model; $integrity = $p.Integrity

        Write-Log ("Humanizing '{0}' -> {1}" -f $p.Label, $p.Dir) 'STEP'
        $result = Invoke-Humanize -Model $model -OutputRoot $p.Dir `
                    -KeepOriginalNames:$KeepOriginalNames -VerifyCopy:(-not $SkipIntegrityCheck) `
                    -OmitBucketFolder:$FinalOnly
        Write-Log ("Copied {0} file(s); {1} missing; {2} copy mismatch(es)." -f `
            $result.Copied, $result.Missing, $result.CopyBad) ($(if ($result.CopyBad -eq 0 -and $result.Missing -eq 0) { 'OK' } else { 'WARN' }))

        $invPath = Join-Path $p.Dir 'inventory.csv'
        $result.Inventory | Export-Csv -LiteralPath $invPath -NoTypeInformation -Encoding UTF8
        Write-Log "Inventory written: $invPath" 'OK'

        $indexPath = New-HtmlViewer -Model $model -OutputRoot $p.Dir -Integrity $integrity -FinalOnlyCopy:$FinalOnly
        Write-Log "Viewer written: $indexPath" 'OK'

        if ($FinalOnly) {
            $notePath = Join-Path $p.Dir 'About this folder.txt'
            Write-EtmfScopeNote -Path $notePath -Batch $model.Batch -Included $model.Documents.Count `
                -Withheld $withheld -SourceInput $ArchivePath -ExclusionList $ExclusionList
            Write-Log "Scope note written: $notePath" 'OK'
        }

        $docs = $model.Documents
        $written.Add([pscustomobject]@{
            Event = $p.Label; EventLabel = $p.Folder; Dir = $p.Dir; Index = $indexPath
            Docs = $docs.Count; Copied = $result.Copied; Missing = $result.Missing; CopyBad = $result.CopyBad
            Final      = @($docs | Where-Object { $_.Bucket -eq '01 Final' }).Count
            InProgress = @($docs | Where-Object { $_.Bucket -eq '02 In Progress' }).Count
            Superseded = @($docs | Where-Object { $_.Bucket -eq '03 Superseded' }).Count
            Excluded   = @($docs | Where-Object { $_.Excluded }).Count
        })
    }

    $landing = $null
    if ($split) {
        $landing = New-LandingPage -OutputRoot $outputRoot -Events $written -Study $b0.SpecificationId
        Write-Log "Landing page written: $landing" 'OK'
    }

    # Gap causes, most frequent first (empty when everything bound).
    $gaps = @($plan | ForEach-Object { $_.Model.Documents } |
              Where-Object { -not $_.FileExists } |
              Group-Object { if ($_.GapReason) { $_.GapReason } else { 'no reason recorded' } } |
              Sort-Object Count -Descending |
              ForEach-Object { [pscustomobject]@{ Count = $_.Count; Reason = $_.Name } })

    Show-Recap -OutputRoot $outputRoot -Written $written -Landing $landing -Gaps $gaps `
               -Demoted $script:Demoted -Flagged $script:Flagged -Excluded $script:Excluded -Withheld $withheld

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
    # Persist the run log. DATED, so re-running never overwrites the record of the previous run, and
    # so a log that arrives by mail identifies itself. It lands next to the output when there is one,
    # and beside the input otherwise - a run that failed early is exactly when the log matters most.
    $logDir = $null
    if ($script:OutputRoot -and (Test-Path -LiteralPath $script:OutputRoot)) { $logDir = $script:OutputRoot }
    elseif ($script:LogFallbackDir -and (Test-Path -LiteralPath $script:LogFallbackDir)) { $logDir = $script:LogFallbackDir }
    if ($logDir) {
        $logPath = Join-Path $logDir ("_log_{0}.txt" -f $script:RunStamp)
        try {
            $script:LogLines | Set-Content -LiteralPath $logPath -Encoding UTF8 -ErrorAction Stop
            Write-Host ("  Run log       : {0}" -f $logPath) -ForegroundColor Green
        }
        catch { Write-Host ("  (could not write the run log to {0}: {1})" -f $logPath, $_.Exception.Message) -ForegroundColor Yellow }
    }
    # Remove the temp extraction (files have been copied out by now). Keyed off $script:TempDir,
    # not $source, so a throw *during* extraction still cleans up rather than leaving clinical
    # documents behind in the temp folder.
    if ($script:TempDir -and (Test-Path -LiteralPath $script:TempDir)) {
        Remove-Item -LiteralPath $script:TempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

