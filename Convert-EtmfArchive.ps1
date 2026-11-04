<#
.SYNOPSIS
    Turns a Viedoc eTMF-EMS archive (.zip) into a human-friendly
    rendering: a clean eTMF folder tree with readable file names and the export's own status on
    every file, an inventory.csv, plus a self-contained HTML mini-viewer.

.DESCRIPTION
    Reads the eTMF Exchange Mechanism Standard (eTMF-EMS v1.0) manifest (*_exchange.xml)
    and the referenced files, then:
      * parses every OBJECT / FILE / AUDITRECORD / SIGNATURE / METADATA into an object model,
      * classifies each document (per object) by lifecycle (audit events) AND version state,
      * copies + renames files into a readable tree, verifying MD5 on every copy,
      * writes inventory.csv, a run log, and a self-contained index.html viewer.

    The eTMF tree is the backbone: Zone > Section > Artifact > <Country/Site> > files, one
    folder per artifact with every version in it. The status of each file (Finalized, Awaiting
    review, Superseded, ...) is in its file name, in inventory.csv and in the viewer.
    -SeparateNotFinalized (opt-in) moves the records that are not Finalized into a sibling
    "Not finalized" tree.
    -FinalOnly produces the client-facing counterpart: the latest Finalized version of each
    document only, plus a note, and withheld.csv (beside the folder) recording what was left out.
    One Viedoc export holds one manifest per milestone; they are combined into that one tree and
    the milestone is recorded as data, in inventory.csv and in the viewer, instead of as a folder
    level. The original archive is never modified.

.PARAMETER ArchivePath
    Path to the eTMF archive .zip, exactly as exported from Viedoc (do not unzip it first).

.PARAMETER OutputPath
    Destination folder for the humanized output. Defaults to "<input directory>\<StudyName>",
    where the folder name is built from $DefaultOutputNameTemplate near the top of the script.

.PARAMETER FinalOnly
    Produce a copy containing ONLY the latest Finalized version of each document - the rows whose
    last lifecycle event in the export is Finalized and that no newer Finalized version replaces - for
    handing to a client or sponsor. Older versions and records not Finalized are left out, and the status
    subfolders are dropped (everything in this copy is final). An "About this folder.txt" states what the
    copy contains, what was left out and how many of each, and withheld.csv, saved beside the folder with the run log, lists every row left out with
    the reason, so the scope of the handover is on record.
    This is a convenience copy, NOT the complete trial master file - run without the switch for that.

.PARAMETER SeparateNotFinalized
    Put the records whose last lifecycle event in the export is NOT Finalized (Awaiting review,
    Flagged by QC Comment, or none) in a sibling folder, "Not finalized", that mirrors the
    Zone/Section/Artifact structure. "Documents" then holds only Finalized records - including
    older versions that were Finalized before a newer version replaced them. Nothing is deleted:
    the complete TMF is both folders together, and "About this folder.txt" says so.
    Ignored with -FinalOnly (which already leaves those records out).

.PARAMETER SkipIntegrityCheck
    Skip MD5 integrity verification of referenced files (verification is on by default).

.PARAMETER Force
    Overwrite the output folder if it already exists. Only a folder made by this tool is ever
    replaced (it holds a .etmf-humanizer-output marker, or index.html and inventory.csv from an
    older version), and never one that is, or contains, the input - however the path is spelled
    or linked - or a filesystem root.

.PARAMETER DryRun
    Parse, classify and print the summary only — write nothing.

.EXAMPLE
    .\Convert-EtmfArchive.ps1 -ArchivePath ".\2022 - Demo Study_eTMFArchive_20260626085449.zip"

.EXAMPLE
    .\Convert-EtmfArchive.ps1 -ArchivePath ".\archive" -Force

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
#   2. #requires ABOVE the help block: PowerShell then stops associating the comment-based help
#      with the script. Keep #requires below it.
#requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Position = 0)]   # optional: with no value, the script prints friendly usage and exits
    [string]$ArchivePath,

    [Parameter()]
    [string]$OutputPath,

    [Parameter()]
    [switch]$FinalOnly,

    [Parameter()]
    [switch]$SeparateNotFinalized,

    [Parameter()]
    [switch]$SkipIntegrityCheck,

    [Parameter()]
    [switch]$Force,

    [Parameter()]
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ----------------------------------------------------------------------------------------
#  Constants
# ----------------------------------------------------------------------------------------
$script:EmsNamespace = 'https://tmfrefmodel.com/ems'
# Stamped into the run log, the viewer and inventory.csv. A support mail that says "it still fails"
# is unanswerable without it: we could not tell whether a reported log came from the build that
# contains a given fix. Bump it with any change to behaviour or output.
$script:ToolVersion  = '2.1.2'
# Characters assumed for the output root when budgeting file names (e.g. C:\Users\name\Documents\Study),
# and for the tree folder under it ('Documents' or 'Not finalized', the longer one counts).
$script:PathBudgetRoot = 60
$script:PathBudgetTree = 13
$script:OutputMarkerName = '.etmf-humanizer-output'    # -Force only deletes a folder that holds this
$script:RunStamp     = (Get-Date).ToString('yyyyMMdd-HHmmss', [System.Globalization.CultureInfo]::InvariantCulture)   # names the log; one file per run
$script:LogLines     = [System.Collections.Generic.List[string]]::new()
$script:LifecycleAudit = $null   # filled while the manifests are read: what the lifecycle decisions relied on
$script:Reconcile = @{           # what the manifests held versus what the tool wrote (Write-EtmfReconciliation)
    Objects = 0; Files = 0; ListedOnce = 0; MergedInto = 0; KeptDistinct = 0; NoFileObjects = 0
    Referenced = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
    Unread = (New-Object System.Collections.Hashtable ([System.StringComparer]::Ordinal))
}

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
    $Message = $Message -replace '[\x00-\x1F\x7F-\x9F]', '?'
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
    # Manifest-derived text can carry control characters (a lone CR would overwrite the screen line
    # or forge a log line). Keep only the line breaks the caller wrote; everything else becomes '?'.
    $Message = ($Message -replace "`r`n", "`n") -replace '[\x00-\x09\x0B-\x1F\x7F-\x9F]', '?'
    foreach ($line in ($Message -split "`n")) { $script:LogLines.Add($line.TrimEnd()) }
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

# One place that decides how untrusted XML is read: no DTD at all (an EMS manifest has none, and a
# DTD is how entity-expansion bombs and external entities arrive - on .NET Framework / PowerShell 5.1
# XmlDocument.Load(path) imposes no limit on entity expansion), no resolver, and a hard size cap.
$script:EtmfMaxManifestBytes = 256MB
function New-EtmfXmlReader {
    param([string]$Path)
    $len = (Get-Item -LiteralPath $Path).Length
    if ($len -gt $script:EtmfMaxManifestBytes) {
        throw ("'{0}' is {1:N0} bytes, larger than the {2:N0}-byte limit for a manifest." -f (Split-Path $Path -Leaf), $len, $script:EtmfMaxManifestBytes)
    }
    $settings = New-Object System.Xml.XmlReaderSettings
    $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
    $settings.XmlResolver   = $null
    # From a stream, not a path: XmlReader.Create(string) reads its argument as a URI, so a folder name
    # holding '%41' is a different file. The reader owns (and closes) the stream.
    $settings.CloseInput = $true
    $stream = New-Object System.IO.FileStream ($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
    try { return [System.Xml.XmlReader]::Create($stream, $settings) }
    catch { $stream.Dispose(); throw }
}

# Does the start of this file mention a BATCH element (as ASCII, UTF-8 or UTF-16 text)? Used only
# to tell an unreadable MANIFEST (stop) from an unrelated file that happens to be named like one.
function Test-EtmfLooksLikeManifest {
    param([string]$Path)
    try {
        $stream = [System.IO.File]::OpenRead($Path)
        try {
            $buffer = New-Object byte[] 65536
            $count  = $stream.Read($buffer, 0, $buffer.Length)
        }
        finally { $stream.Dispose() }
        $head = New-Object byte[] $count
        [Array]::Copy($buffer, $head, $count)
        $ascii = [System.Text.Encoding]::GetEncoding(28591).GetString($head)
        $wide  = [System.Text.Encoding]::Unicode.GetString($head) + [System.Text.Encoding]::BigEndianUnicode.GetString($head)
        return (($ascii + $wide) -cmatch '<([A-Za-z_][\w.-]*:)?BATCH[\s>/]')
    }
    catch { return $true }      # cannot even read it: do not wave it through
}

# What is this file named like a manifest? Three answers, because they must be treated differently:
#   Manifest    - well-formed XML, no DTD, root element BATCH.
#   NotManifest - well-formed XML whose root is something else: an ordinary document that merely has
#                 such a name (or something planted); it must not be able to stop the run.
#   Broken      - could not be read as XML at all (truncated, a DTD, over the size cap, locked,
#                 wrong encoding declaration). This may well be a REAL manifest: skipping it would
#                 silently drop a milestone from the TMF, so the caller must stop and say why.
function Test-EtmfManifestFile {
    param([string]$Path)
    $reader = $null
    try {
        $reader = New-EtmfXmlReader $Path
        while ($reader.Read()) {
            if ($reader.NodeType -eq [System.Xml.XmlNodeType]::Element) {
                $state = if ($reader.LocalName -ceq 'BATCH') { 'Manifest' } else { 'NotManifest' }
                return [pscustomobject]@{ State = $state; Reason = '' }
            }
        }
        return [pscustomobject]@{ State = 'Broken'; Reason = 'the file has no root element' }
    }
    catch { return [pscustomobject]@{ State = 'Broken'; Reason = $_.Exception.Message } }
    finally { if ($null -ne $reader) { $reader.Dispose() } }
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
# ONE rule, shared by the extractor (Expand-EtmfZip) and the lookup (Get-ContentUrlTail). If the
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
    # Both separators: Get-ContentPathSegments already accepts backslashes,
    # and splitting on '/' alone made a backslash-separated CONTENTURL report itself as "malformed",
    # which tells a customer their export is broken when it is merely written the other way round.
    $segments = $ContentUrl -split '[\\/]+'
    if ($segments.Count -lt 2) { return $null }
    $parts = @(Get-ContentPathSegments (($segments | Select-Object -Skip 1) -join '/'))
    if ($parts.Count -eq 0 -or ($parts -contains '..')) { return $null }
    return ($parts -join '/')
}

# The same key with leading/trailing white space (a space, TAB, CR/LF ...) removed from the FILE NAME, the
# last segment only; $null when there is nothing to remove. Viedoc's archive writer names a zip entry from
# the file name TRIMMED, but writes the manifest's CONTENTURL from the same name untrimmed, so a document
# whose stored name begins with a space is listed as 'folder/ name.pdf' and sits in the zip as
# 'folder/name.pdf' (seen in a real export: 27 of 9,969 documents). Import-EtmfManifest tries this key
# only after the exact one missed, and binds only when the file's checksum equals the manifest's.
function Get-EtmfTrimmedNameKey {
    param([string]$Key)
    $slash   = $Key.LastIndexOf('/')
    $name    = $Key.Substring($slash + 1)
    $trimmed = $name.Trim()
    if ($trimmed.Length -eq 0 -or $trimmed -ceq $name) { return $null }
    return $Key.Substring(0, $slash + 1) + $trimmed
}

# A size for a message a person reads: "20.2 GB", not "21,727,903,744 bytes". Invariant, so the
# text does not depend on the machine's language.
function Format-EtmfSize {
    param([double]$Bytes)
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    if ($Bytes -ge 1GB) { return ($Bytes / 1GB).ToString('N1', $inv) + ' GB' }
    if ($Bytes -ge 1MB) { return ($Bytes / 1MB).ToString('N0', $inv) + ' MB' }
    return ([Math]::Ceiling($Bytes / 1KB)).ToString('N0', $inv) + ' KB'
}

# Free bytes on the volume that holds $Path, or $null when that cannot be told (a network share, a
# volume the platform will not report). $null means "not checked", never "no space": a check that
# cannot measure must not refuse a run.
function Get-EtmfFreeBytes {
    param([string]$Path)
    try {
        $full = [System.IO.Path]::GetFullPath($Path)
        if ([System.IO.Path]::DirectorySeparatorChar -eq '\') {
            $root = [System.IO.Path]::GetPathRoot($full)
            if ($root -notmatch '^[A-Za-z]:\\') { return $null }        # UNC share
            return [long]([System.IO.DriveInfo]::new($root)).AvailableFreeSpace
        }
        # Elsewhere the volume is found from the path itself, which must exist: walk up to what does.
        $probe = $full
        while ($probe -and -not (Test-Path -LiteralPath $probe)) { $probe = [System.IO.Path]::GetDirectoryName($probe) }
        if (-not $probe) { return $null }
        return [long]([System.IO.DriveInfo]::new($probe)).AvailableFreeSpace
    }
    catch { return $null }
}

# Would $Bytes fit on the volume holding $Path? Returns $null when that is not known; otherwise
# Free, Need (the bytes plus headroom for the inventory, the viewer, the log and the filesystem's own
# overhead) and Short ($true when it does not fit).
function Get-EtmfSpaceCheck {
    param([string]$Path, [double]$Bytes)
    $free = Get-EtmfFreeBytes $Path
    if ($null -eq $free) { return $null }
    $need = $Bytes + 256MB
    return [pscustomobject]@{ Free = $free; Need = $need; Short = ($free -lt $need) }
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
    # Invisible and direction-changing characters (right-to-left override and friends) let a name read
    # as something it is not - 'report<U+202E>fdp.exe' shows as 'reportexe.pdf' - so they go.
    $clean = ($clean -replace '[\u200B-\u200F\u202A-\u202E\u2060-\u2069\uFEFF\x7F-\x9F]', '_')
    $clean = $clean.Trim().TrimEnd('.', ' ')   # Windows dislikes trailing dot/space
    if ([string]::IsNullOrWhiteSpace($clean)) { $clean = 'file' }
    # Windows device names (CON, NUL, COM1 ...) are not file names there, with or without an extension.
    if ($clean -match '^(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(\..*)?$') { $clean = '_' + $clean }
    return $clean
}

# Write rows as a CSV that Excel opens correctly and safely: UTF-8 WITH a byte-order mark (without it
# Excel reads non-ASCII names as mojibake, and PowerShell 7 writes none), and any text cell that
# starts with = + - @ TAB or CR gets a leading apostrophe, because Excel evaluates such a cell as a
# formula (a title like =HYPERLINK(...) in an untrusted manifest). Only the CSV is changed; the
# viewer and the manifest carry the value as written.
function Write-EtmfCsv {
    param([string]$Path, $Rows)
    $safe = foreach ($r in @($Rows)) {
        $o = [ordered]@{}
        foreach ($p in $r.PSObject.Properties) {
            $v = $p.Value
            # The checksum column is exempt when it has exactly the shape of a base64 MD5 (22 characters
            # and '=='): it can legitimately start with '+', must stay equal to the manifest, and has no
            # parentheses, quotes or separators, so it cannot call anything in a spreadsheet.
            $isChecksum = ($p.Name -ceq 'Integrity' -and $v -is [string] -and $v -cmatch '^[A-Za-z0-9+/]{22}==$')
            if (-not $isChecksum -and $v -is [string] -and $v.Length -gt 0 -and "=+-@`t`r".IndexOf($v[0]) -ge 0) { $v = "'" + $v }
            $o[$p.Name] = $v
        }
        [pscustomobject]$o
    }
    $lines = @($safe | ConvertTo-Csv -NoTypeInformation)
    [System.IO.File]::WriteAllText($Path, (($lines -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
}

# "01.01.trial oversight" -> "01.01 Trial Oversight"  (numeric prefix + Title Case name)
function Get-PrettyTmfSegment {
    param([string]$Segment)
    # Invariant on purpose: under a Turkish culture "izmir" would become "İzmir", so the same export
    # would be filed under different folder names depending on the machine's language settings.
    $ti = [System.Globalization.CultureInfo]::InvariantCulture.TextInfo
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
    if ($Segment.Length -le 3)  { return $Segment.ToUpperInvariant() }
    return [System.Globalization.CultureInfo]::InvariantCulture.TextInfo.ToTitleCase($Segment)
}

# Split a CONTENTURL into TMF path parts using the TransferID as anchor.
# Bounds-safe: a short/malformed CONTENTURL yields empty parts (never throws under StrictMode),
# so the referenced document degrades to a recorded gap rather than aborting the whole run.
function Get-ContentPathParts {
    param([string]$ContentUrl, [string]$TransferId)
    $segs = $ContentUrl -split '[\\/]'      # both separators, as the lookup does (Get-ContentUrlTail)
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


# ----------------------------------------------------------------------------------------
#  Lifecycle (what the record's own audit trail says about its approval)
# ----------------------------------------------------------------------------------------
# The lifecycle is READ from the audit events, never inferred from words inside them and never
# given a name of our own. An event's NAME is its first line up to the first colon; everything after
# it is user text (a QC comment can say "Please finalize" or "Blocked by sponsor", and used to make
# the record Finalized / Locked because of it). Viedoc writes a small fixed vocabulary, and the tool
# knows exactly these names, compared exactly (case included), never as "contains":
#
#   Awaiting review | Flagged by QC Comment | Finalized     -> the record's lifecycle
#   Document locked | Document unlocked                     -> its lock state (a separate fact)
#
# The Lifecycle is the export's own name for the LAST of the first three, in time order, and the
# Lock is the export's own name for the last lock event. A record with none of them has an empty
# Lifecycle: the export did not say, and nothing is made up to fill the gap.
#
# Every other event name is inert ("Previewed", "Uploaded as Unpublished", "Saved as Unpublished:",
# ...). A name that merely LOOKS like a lifecycle event (a renamed event, a different spelling, a
# future wording) is never acted on, but it is reported as a near miss: ignoring it silently could
# leave an approved document outside the approved set. The report is a warning, not a classification.
#
# Events are pooled across a record's files, so they are ordered by timestamp, not by position.
# Within one timestamp the natural order applies (review, approval, lock); a pair that contradicts
# itself (Finalized vs Flagged by QC Comment, locked vs unlocked) resolves towards the reading that
# claims LESS approval, and is reported.

# A timestamp from the export. ISO-8601 only (what Viedoc writes: 2026-11-01T18:02:18+00:00): a
# lenient parse would read 'Jan 5' or '10:00' against TODAY's date and '03/04/2024' month-first, so
# the same export would classify differently depending on the day it was run. Anything else is
# $null, which callers treat as "no usable timestamp" and report. No offset means UTC.
$script:EtmfIsoFormats = [string[]]@(
    'yyyy-MM-dd\THH:mm:ssK', 'yyyy-MM-dd\THH:mm:ss.FFFFFFFK', 'yyyy-MM-dd\THH:mmK',
    'yyyy-MM-dd\THH:mm:ss', 'yyyy-MM-dd\THH:mm:ss.FFFFFFF', 'yyyy-MM-dd\THH:mm')
function ConvertFrom-EtmfTimestamp {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $when = [DateTimeOffset]::MinValue
    if ([DateTimeOffset]::TryParseExact($Text.Trim(), $script:EtmfIsoFormats, [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::AssumeUniversal, [ref]$when)) { return $when }
    return $null
}

# The name of an audit event: its first line, up to the first colon.
function Get-EtmfEventHead {
    param([string]$Event)
    $first = (([string]$Event) -split "`r?`n")[0]
    $colon = $first.IndexOf(':')
    if ($colon -ge 0) { $first = $first.Substring(0, $colon) }
    return $first.Trim()
}

# Decide the lifecycle and say how. $AuditEvents holds event texts (strings), or objects with Event
# and Timestamp. Returns the Lifecycle and Lock (the export's own words, '' when absent) plus what
# the decision relied on, so a run can report it.
function Get-EtmfLifecycleDetail {
    param([object[]]$AuditEvents)

    # Case-SENSITIVE and exact: the names are Viedoc's, and a different spelling is a different thing.
    $kinds = New-Object System.Collections.Hashtable ([System.StringComparer]::Ordinal)
    $kinds['Awaiting review']        = 'Lifecycle'
    $kinds['Flagged by QC Comment']  = 'Lifecycle'
    $kinds['Finalized']              = 'Lifecycle'
    $kinds['Document locked']        = 'Lock'
    $kinds['Document unlocked']      = 'Lock'
    # Within ONE timestamp the reading that claims LESS approval comes last and so wins: a tie can
    # only err towards "not approved" (and is reported). Finalized < Awaiting review < Flagged by QC
    # Comment; Document locked < Document unlocked.
    $rank = New-Object System.Collections.Hashtable ([System.StringComparer]::Ordinal)
    $rank['Finalized'] = 0; $rank['Awaiting review'] = 1; $rank['Flagged by QC Comment'] = 2
    $rank['Document locked'] = 3; $rank['Document unlocked'] = 4

    $relevant = [System.Collections.Generic.List[object]]::new()
    $used     = [System.Collections.Generic.List[string]]::new()
    $ignored  = [System.Collections.Generic.List[string]]::new()
    $nearMiss = [System.Collections.Generic.List[string]]::new()
    $notes    = [System.Collections.Generic.List[string]]::new()
    $inv      = [System.Globalization.CultureInfo]::InvariantCulture
    $index    = 0

    foreach ($e in @($AuditEvents)) {
        $text = ''; $stamp = ''
        if ($e -is [string]) { $text = $e }   # plain text carries no time: list order
        elseif ($null -ne $e) {
            $text = [string]$e.Event
            if ($e.PSObject.Properties['Timestamp']) { $stamp = [string]$e.Timestamp }
        }
        $head = Get-EtmfEventHead $text
        $index++
        if (-not $kinds.ContainsKey($head)) {
            # Only a warning aid: this never classifies anything.
            if ($head -match '(?i)finali[sz]|lock|awaiting|flag|approv|reject|revok|withdr|\breview') {
                if (-not $nearMiss.Contains($head)) { $nearMiss.Add($head) }
            }
            elseif ($head -and -not $ignored.Contains($head)) { $ignored.Add($head) }
            continue
        }
        if (-not $used.Contains($head)) { $used.Add($head) }
        $when  = [DateTimeOffset]::MinValue
        $parsed = ConvertFrom-EtmfTimestamp $stamp
        $timed = ($null -ne $parsed)
        if ($timed) { $when = $parsed }
        $hasOffset = ($stamp -match '(Z|[+-]\d\d:\d\d)\s*$')
        $relevant.Add([pscustomobject]@{ Head = $head; Group = $kinds[$head]; When = $when; Timed = $timed; ListOnly = ($e -is [string]); HasOffset = $hasOffset; Index = $index; Rank = $rank[$head] })
    }

    # Lifecycle and lock are separate facts, so each is ordered on its own: a lock event without a
    # time must not cost the lifecycle its (fully timed) order. Time order when every event of the
    # group has a usable timestamp; otherwise the order they were listed in, and say so - a guess
    # about order must never be silent, and must never claim approval.
    $lifecycle = ''; $lock = ''
    foreach ($kind in @('Lifecycle', 'Lock')) {
        $mine = @($relevant | Where-Object { $_.Group -eq $kind })
        if ($mine.Count -eq 0) { continue }
        $ordered = $mine
        if ($mine.Count -gt 1) {
            $listOnly = (@($mine | Where-Object { $_.ListOnly }).Count -gt 0)
            $why = ''
            if (-not $listOnly) {
                if (@($mine | Where-Object { -not $_.Timed }).Count -gt 0) {
                    $why = 'an audit event has no usable (ISO-8601) timestamp'
                }
                elseif (@($mine | Where-Object { $_.HasOffset }).Count -gt 0 -and @($mine | Where-Object { -not $_.HasOffset }).Count -gt 0) {
                    $why = 'audit timestamps with and without a time zone offset are mixed'
                }
            }
            if ($why) {
                # The order is a guess. A guess must not claim approval: if the events disagree, the
                # reading that claims less is used (not Finalized; not locked).
                $heads = @($mine | ForEach-Object { $_.Head } | Select-Object -Unique)
                $claimsMost = if ($kind -eq 'Lifecycle') { 'Finalized' } else { 'Document locked' }
                if ($heads -contains $claimsMost -and $heads.Count -gt 1) {
                    $notes.Add("$why and the $($kind.ToLowerInvariant()) events disagree, so the reading that claims less was used")
                    $ordered = @($mine | Where-Object { $_.Head -cne $claimsMost })
                }
                else {
                    $notes.Add("$why, so its events were read in the order the manifest lists them")
                }
            }
            elseif (-not $listOnly) {
                $ordered = @($mine | Sort-Object -Property @{ Expression = { $_.When.UtcTicks } },
                                                           @{ Expression = { $_.Rank } },
                                                           @{ Expression = { $_.Index } })
                foreach ($t in @($ordered | Group-Object -Property { $_.When.UtcTicks } | Where-Object { $_.Count -gt 1 })) {
                    $tied = @($t.Group | ForEach-Object { $_.Head } | Select-Object -Unique)
                    if ($tied.Count -gt 1) {
                        $notes.Add('audit events with the same timestamp disagree (' + ($tied -join ' / ') + '); the reading that claims less approval was used')
                    }
                }
            }
        }
        $last = $ordered[$ordered.Count - 1].Head
        if ($kind -eq 'Lifecycle') { $lifecycle = $last } else { $lock = $last }
    }
    if ($lock -and $lifecycle -ne 'Finalized') {
        $notes.Add('a lock event is not backed by a Finalized lifecycle; the record was not treated as approved')
    }

    return [pscustomobject]@{
        Lifecycle = $lifecycle
        Lock      = $lock
        Used      = @($used)
        Ignored   = @($ignored)
        NearMiss  = @($nearMiss)
        Notes     = @($notes)
    }
}

function Get-EtmfLifecycle {
    param([object[]]$AuditEvents)
    return (Get-EtmfLifecycleDetail -AuditEvents @($AuditEvents)).Lifecycle
}

# What a run's lifecycle decisions relied on: which event names counted, which were ignored, which
# looked like lifecycle events but were not known, and every judgement call. Collected here while
# the manifests are read and written to the log afterwards (Write-EtmfLifecycleAudit), so a reviewer
# can see the vocabulary an approval decision rested on instead of taking it on trust.
function Add-EtmfLifecycleAudit {
    param($Detail)
    if ($null -eq $script:LifecycleAudit) {
        # Ordinal: 'Finalized' and 'finalized' are different names and must not merge in the log.
        $table = { New-Object System.Collections.Hashtable ([System.StringComparer]::Ordinal) }
        $script:LifecycleAudit = @{ Records = 0; Unknown = 0; Used = (& $table); Ignored = (& $table); NearMiss = (& $table); Notes = (& $table) }
    }
    $a = $script:LifecycleAudit
    $a.Records++
    if (-not $Detail.Lifecycle) { $a.Unknown++ }
    foreach ($pair in @(@('Used', $Detail.Used), @('Ignored', $Detail.Ignored), @('NearMiss', $Detail.NearMiss), @('Notes', $Detail.Notes))) {
        foreach ($h in @($pair[1])) {
            if ($a[$pair[0]].ContainsKey($h)) { $a[$pair[0]][$h]++ } else { $a[$pair[0]][$h] = 1 }
        }
    }
}

# Bucket = the folder a document lands in (two-axis: version state + lifecycle).
function Get-EtmfBucket {
    param([string]$VersionState, [string]$Lifecycle)
    if ($VersionState -cin @('Superseded', 'Obsolete')) { return '03 Superseded' }
    if ($Lifecycle    -ceq 'Finalized')                { return '01 Finalized' }
    return '02 Not finalized'
}

# Status token shown in the file name / badge: the record's LAST status in the export's own words -
# its lifecycle (Awaiting review / Flagged by QC Comment / Finalized), or "Document locked" while the
# record is Finalized and locked (a lock always follows approval, so a locked record is Finalized);
# a lock the export gives with no lifecycle event is shown as it is, and a lock beside another
# lifecycle (not backed by approval) never replaces that lifecycle word.
# Superseded is NOT part of it: when a document has several versions the version number in the name
# already says which are older, and VersionState / SupersededBy say it in the inventory. Obsolete is
# kept (it is the export's own word and changes what the file is). Empty when the export says nothing.
function Get-EtmfStatusToken {
    param([string]$VersionState, [string]$Lifecycle, [string]$Lock)
    $status = if ($Lock -ceq 'Document locked' -and ($Lifecycle -ceq 'Finalized' -or -not $Lifecycle)) { $Lock } else { $Lifecycle }
    if ($VersionState -ceq 'Obsolete') {
        if ($status) { return ('Obsolete; ' + $status) }
        return 'Obsolete'
    }
    return $status
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
    foreach ($seg in ($sa + $sb)) { if ($seg -notmatch '^[0-9]+$') { $numeric = $false; break } }
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
    # Not plain dotted numbers ('v9' / 'v10', '9a' / '10'): compare runs of digits as numbers and
    # everything else as text, so the number inside a label decides and the order does not flip
    # at 10. Still ordinal and culture-free.
    $ta = [regex]::Matches(([string]$A).Trim(), '[0-9]+|[^0-9]+') | ForEach-Object { $_.Value }
    $tb = [regex]::Matches(([string]$B).Trim(), '[0-9]+|[^0-9]+') | ForEach-Object { $_.Value }
    $ta = @($ta); $tb = @($tb)
    $max = [Math]::Max($ta.Count, $tb.Count)
    for ($i = 0; $i -lt $max; $i++) {
        if ($i -ge $ta.Count) { return -1 }
        if ($i -ge $tb.Count) { return 1 }
        $x = $ta[$i]; $y = $tb[$i]
        if ($x -match '^[0-9]+$' -and $y -match '^[0-9]+$') {
            $x = $x.TrimStart('0'); $y = $y.TrimStart('0')
            if ($x.Length -ne $y.Length) { if ($x.Length -lt $y.Length) { return -1 } else { return 1 } }
        }
        $c = [string]::CompareOrdinal($x, $y)
        if ($c -lt 0) { return -1 }
        if ($c -gt 0) { return 1 }
    }
    return 0
}


# ----------------------------------------------------------------------------------------
#  Input resolution: the archive is a .zip, extracted entry by entry to a temp folder.
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
    $entryTotal = 0
    try {
        # Say so BEFORE the extraction if the temp drive cannot hold it: a 20 GB export that fills the disk
        # fails twenty minutes in, with an error ("out of memory", "disk full") that names no cause.
        $unpacked = 0L
        foreach ($entry in $zip.Entries) { $unpacked += $entry.Length }
        # Deflate cannot shrink data more than about 1000:1. A total far beyond that is a table of contents .NET
        # misread (the same thing that gives a 5.1 read of some zips no entries at all), not a size to refuse on.
        $space = $null
        if ($unpacked -le ((Get-Item -LiteralPath $ZipPath).Length * 1100)) { $space = Get-EtmfSpaceCheck -Path $Destination -Bytes $unpacked }
        else { Write-Log "The sizes this zip lists add up to an implausible total, so the free-space check is skipped." 'WARN' }
        if ($null -ne $space) {
            $driveName = [System.IO.Path]::GetPathRoot([System.IO.Path]::GetFullPath($Destination))
            Write-Log ("Unpacking needs about {0} of temporary space on {1}; {2} is free." -f (Format-EtmfSize $unpacked), $driveName, (Format-EtmfSize $space.Free)) 'INFO'
            if ($space.Short) {
                throw ("Not enough free space on {0} to unpack this archive: it unpacks to about {1} in the temporary folder, and {2} is free. " -f $driveName, (Format-EtmfSize $unpacked), (Format-EtmfSize $space.Free)) +
                      'Free up space there, or point the temporary folder at a drive with room before running again (in PowerShell, after creating the folder: $env:TEMP = ''D:\tmp''; $env:TMP = $env:TEMP).'
            }
        }
        foreach ($entry in $zip.Entries) {
            $entryTotal++                   # every entry .NET could read, directories and manifests included (counted here: @($empty).Count is 1 and $empty.Count throws under StrictMode)
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
                $why = $_.Exception.Message
                Write-Log "Could not extract '$($entry.FullName)': $why" 'WARN'
                if ($isManifest) {
                    # Skipping a manifest would drop a whole milestone from a run that then reports success.
                    throw "The manifest '$rawRel' could not be extracted from the archive ($why). Not skipping it: that would silently drop a milestone from the TMF."
                }
                # A partly written file must not be mistaken for the document; the key stays on record so
                # the gap says the entry exists but could not be extracted (not that the archive lacks it).
                if (Test-Path -LiteralPath $dest) { Remove-Item -LiteralPath $dest -Force -ErrorAction SilentlyContinue }
                if ($files.ContainsKey($rawRel)) { $files[$rawRel] = '' } else { $files[$rawRel] = $dest }
                $skipped++; continue
            }

            if ($isManifest) {
                $sniff = Test-EtmfManifestFile $dest
                if ($sniff.State -eq 'Broken' -and -not (Test-EtmfLooksLikeManifest $dest)) {
                    Write-Log "'$rawRel' is named like a manifest but is not readable XML and does not mention a BATCH element; it is treated as an ordinary file." 'WARN'
                    $sniff = [pscustomobject]@{ State = 'NotManifest'; Reason = 'warned' }
                }
                if ($sniff.State -eq 'Broken') {
                    throw "'$rawRel' is named like a manifest but cannot be read as XML ($($sniff.Reason)). Not skipping it: that could silently drop a milestone from the TMF. Fix or remove the file and run again."
                }
                if ($sniff.State -eq 'Manifest') {
                    # Prefix = where this manifest sits in the archive. A CONTENTURL is relative to it,
                    # so prefix + tail reconstructs the entry path exactly.
                    $prefix = ''
                    if ($parts.Count -gt 1) { $prefix = (@($parts[0..($parts.Count - 2)]) -join '/') }
                    $manifests.Add([pscustomobject]@{ Path = $dest; Prefix = $prefix })
                    continue
                }
                # Named like a manifest, but well-formed XML with another root element: an ordinary
                # file of the archive; it must not stop the conversion.
                if ($sniff.Reason -ne 'warned') { Write-Log "'$rawRel' is named like a manifest but its root element is not BATCH; it is treated as an ordinary file." 'WARN' }
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
    return [pscustomobject]@{ FileCount = $files.Count; EntryCount = $entryTotal; Skipped = $skipped; Files = $files; Manifests = $manifests }
}

function Resolve-EtmfSource {
    param([string]$ArchivePath)

    if (-not (Test-Path -LiteralPath $ArchivePath)) {
        throw "ArchivePath not found: $ArchivePath"
    }
    $item = Get-Item -LiteralPath $ArchivePath

    if (-not $item.PSIsContainer -and $item.Extension -ieq '.zip') {
        $temp = Join-Path ([IO.Path]::GetTempPath()) ("etmf_" + [Guid]::NewGuid().ToString('N').Substring(0, 8))
        New-Item -ItemType Directory -Path $temp -Force | Out-Null
        $script:TempDir = $temp        # recorded now so the finally in main.ps1 can clean up
        Write-Log "Extracting archive to temp: $temp" 'STEP'                # even if extraction throws
        $x = Expand-EtmfZip -ZipPath $item.FullName -Destination $temp
        Write-Log "Extracted $($x.FileCount) file(s)" 'INFO'
        if ($x.FileCount -eq 0) {
            # Nothing to convert, so stop here and say why rather than fall through to "no manifest found",
            # which reads as a defect in the export. A zip that opens but yields no entries at all is what
            # .NET does with a table of contents it cannot follow (seen in a real 20 GB export read by Windows
            # PowerShell 5.1, which PowerShell 7 read fine); no code of ours drops an entry silently.
            $zipBytes = $item.Length
            if ($x.Manifests.Count -eq 0) {
                # Not a single file, manifest or document. How many entries .NET did find is worth saying: a damaged
                # table of contents can yield none, or a bogus one whose name is empty.
                $why = ("No files could be read from '{0}' ({1} bytes): this PowerShell found {2} entr{3} in it and none is a file. " -f `
                    $item.Name, $zipBytes.ToString('N0', [System.Globalization.CultureInfo]::InvariantCulture), $x.EntryCount, $(if ($x.EntryCount -eq 1) { 'y' } else { 'ies' }))
                if ($PSVersionTable.PSVersion.Major -lt 6) {
                    $why += ("This is Windows PowerShell {0}; it cannot read some large zips that PowerShell 7 reads. Run the script with PowerShell 7 (pwsh). If PowerShell 7 reads nothing either, the zip is incomplete or damaged: download or export it again." -f $PSVersionTable.PSVersion)
                }
                else {
                    $why += 'The zip is incomplete or damaged (for example a partial download or a copy that was cut short): download or export it again.'
                }
            }
            else {
                $why = ("'{0}' holds {1} manifest file(s) but no documents, so there is nothing to convert. Check that this is the complete export zip." -f $item.Name, $x.Manifests.Count)
            }
            throw $why
        }
        # Manifests come from the extractor, not from a directory scan: their temp names are
        # neutral, and each carries the archive prefix its CONTENTURLs are relative to.
        return [pscustomobject]@{ ContentMap = $x.Files; Manifests = $x.Manifests }
    }

    throw "ArchivePath must be the .zip file of the export. Got: $($item.FullName)"
}


# ----------------------------------------------------------------------------------------
#  Manifest parsing -> object model
# ----------------------------------------------------------------------------------------
# The audit trail in time order, keeping the manifest's own order when a timestamp cannot be read.
function Sort-EtmfAuditTrail {
    param($Trail)
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $i = 0
    $keyed = @(foreach ($r in @($Trail)) {
        $i++
        $w = [DateTimeOffset]::MinValue
        $parsed = ConvertFrom-EtmfTimestamp ([string]$r.Timestamp)
        $ok = ($null -ne $parsed)
        if ($ok) { $w = $parsed }
        [pscustomobject]@{ R = $r; Ok = $ok; W = $w; I = $i }
    })
    if (@($keyed | Where-Object { -not $_.Ok }).Count -gt 0) { return @($Trail) }
    return @($keyed | Sort-Object -Property @{ Expression = { $_.W.UtcTicks } }, @{ Expression = { $_.I } } | ForEach-Object { $_.R })
}

# ----------------------------------------------------------------------------------------
#  Fields the manifest carries that the tool does not read
# ----------------------------------------------------------------------------------------
# Every element and METADATA name the tool reads or reproduces. Anything else in a manifest (a
# newer EMS version, a custom field) is not lost from the archive, but it is not in the inventory
# either - so it is counted here and reported in the run log instead of being ignored in silence.
$script:EtmfKnownFields = @{
    'BATCH'       = @('OBJECT')
    'OBJECT'      = @('OBJECTID', 'OBJECTLEVEL', 'COUNTRYID', 'SITESYSTEMID', 'SITEID', 'ORGANIZATIONNAME', 'PERSONNAME', 'UNIQUEID',
                      'ARTIFACTNUMBER', 'OBJECTVERSION', 'OBJECTVERSIONSTATE', 'OBJECTTITLE', 'SUBARTIFACT', 'OBJECTCOPY',
                      'ARTIFACTDATE', 'DATEDESCRIPTION', 'RESTRICTED', 'OBJECTLANGUAGE', 'TRANSLATION', 'OBJECTEXPIRYDATE', 'FILE', 'METADATA')
    'FILE'        = @('INTEGRITY', 'FILENAME', 'CONTENTURL', 'AUDITRECORD', 'SIGNATURE')
    'AUDITRECORD' = @('AUDITID', 'DATETIMESTAMP', 'USERREF', 'AUDITENTRYTYPE', 'AUDITEVENT')
    'SIGNATURE'   = @('SIGNATUREMETHODOLOGY', 'USERID', 'SIGNATURENAME', 'SIGNATUREDATETIME', 'SIGNATUREREASON')
    'METADATA'    = @('DocumentId', 'Version')
}
# Attributes the tool reads (BATCH) or expects (METADATA NAME); any other attribute is reported.
$script:EtmfKnownAttributes = @{
    'BATCH'    = @('STUDYSYSTEMID', 'STUDYID', 'EVENTID', 'TRANSFERID', 'TRANSFERSOURCEID', 'SPECIFICATIONID', 'TMFRMVERSION')
    'METADATA' = @('NAME')
}
# Children that may legitimately repeat; a repeat of any other child is reported because the tool
# reads only the first.
$script:EtmfRepeatable = @{
    'BATCH'  = @('OBJECT')
    'OBJECT' = @('FILE', 'METADATA', 'PERSONNAME')
    'FILE'   = @('AUDITRECORD', 'SIGNATURE')
}
function Add-EtmfUnreadFields {
    param($Element, [string]$Path)
    $note = {
        param([string]$Label)
        if ($script:Reconcile.Unread.ContainsKey($Label)) { $script:Reconcile.Unread[$Label]++ } else { $script:Reconcile.Unread[$Label] = 1 }
    }
    $knownAttr = $script:EtmfKnownAttributes[$Path]
    foreach ($attr in $Element.Attributes) {
        if (-not ($knownAttr -ccontains $attr.LocalName) -and $attr.NamespaceURI -ne 'http://www.w3.org/2000/xmlns/' -and $attr.Name -ne 'xmlns') {
            & $note ($Path + '@' + $attr.LocalName)
        }
    }
    $known    = $script:EtmfKnownFields[$Path]
    $multi    = $script:EtmfRepeatable[$Path]
    $counts   = New-Object System.Collections.Hashtable ([System.StringComparer]::Ordinal)
    foreach ($child in $Element.ChildNodes) {
        if ($child.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
        $name = $child.LocalName
        if ($known -ccontains $name) {
            if ($name -ceq 'METADATA') {
                $metaName = $child.GetAttribute('NAME')
                if (-not ($script:EtmfKnownFields['METADATA'] -ccontains $metaName)) { & $note ($Path + '/METADATA NAME="' + $metaName + '"') }
                $key = 'METADATA NAME="' + $metaName + '"'
                foreach ($attr in $child.Attributes) { if ($attr.LocalName -cne 'NAME') { & $note ($Path + '/METADATA@' + $attr.LocalName) } }
            }
            else {
                $key = $name
                foreach ($attr in $child.Attributes) {
                    if ($attr.NamespaceURI -ne 'http://www.w3.org/2000/xmlns/' -and $attr.Name -ne 'xmlns') { & $note ($Path + '/' + $name + '@' + $attr.LocalName) }
                }
            }
            if ($counts.ContainsKey($key)) { $counts[$key]++ } else { $counts[$key] = 1 }
            continue
        }
        & $note ($Path + '/' + $name)
    }
    foreach ($k in $counts.Keys) {
        if ($counts[$k] -gt 1 -and -not ($multi -ccontains $k) -and ($k -cne 'METADATA' -or $false)) {
            & $note ($Path + '/' + $k + $(if ($k -like 'METADATA*') { ' (repeated; the tool uses the last)' } else { ' (repeated; the tool reads only the first)' }))
        }
    }
}

# ----------------------------------------------------------------------------------------
#  Manifest parsing -> object model
# ----------------------------------------------------------------------------------------
function Import-EtmfManifest {
    # -ContentMap: raw archive path -> extracted temp file, from Expand-EtmfZip. -ContentPrefix:
    # where this manifest sits in the archive, since a CONTENTURL is relative to the manifest but
    # the index is keyed from the archive root.
    param([System.IO.FileInfo]$ManifestFile, $ContentMap, [string]$ContentPrefix = "")

    $xml = New-Object System.Xml.XmlDocument
    # Untrusted manifest: disable DTD / external-entity resolution so a malicious manifest can't
    # mount an XXE attack. .NET Core (PowerShell 7) already defaults to this; .NET Framework
    # (Windows PowerShell 5.1) does NOT, so setting it explicitly is required for a safe 5.1 run.
    $xml.XmlResolver = $null
    $xmlReader = New-EtmfXmlReader $ManifestFile.FullName       # no DTD, no resolver, size-capped
    try { $xml.Load($xmlReader) }
    catch { throw "The manifest '$($ManifestFile.Name)' cannot be read as XML: $($_.Exception.Message)" }
    finally { $xmlReader.Dispose() }
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

    Add-EtmfUnreadFields -Element $batchEl -Path 'BATCH'
    $documents = [System.Collections.Generic.List[object]]::new()
    $objectNodes = @($batchEl.SelectNodes("*[local-name()='OBJECT']"))

    $objectNo = 0
    $script:Reconcile.Objects += $objectNodes.Count
    foreach ($obj in $objectNodes) {
        $objectNo++
        Add-EtmfUnreadFields -Element $obj -Path 'OBJECT'
        # ---- object-level (artifact instance) fields ----
        $artifactNumber = Get-NodeText $obj 'ARTIFACTNUMBER'
        $versionState   = Get-NodeText $obj 'OBJECTVERSIONSTATE'
        $country        = Get-NodeText $obj 'COUNTRYID'
        $org            = Get-NodeText $obj 'ORGANIZATIONNAME'
        # an OBJECT may relate to multiple people; capture them all
        $person         = (@($obj.SelectNodes("*[local-name()='PERSONNAME']") |
                             ForEach-Object { $_.InnerText.Trim() } | Where-Object { $_ }) -join '; ')
        # The title falls back to the sub-artifact, then a placeholder, so folders and names never go
        # bare. TitleSource says which one it was, so a placeholder is never mistaken for the export.
        $artifactTitle  = Get-NodeText $obj 'OBJECTTITLE'
        $titleSource    = 'OBJECTTITLE'
        if (-not $artifactTitle) { $artifactTitle = Get-NodeText $obj 'SUBARTIFACT'; $titleSource = 'SUBARTIFACT' }
        if (-not $artifactTitle) { $artifactTitle = 'Untitled Document'; $titleSource = 'none (placeholder)' }
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
        # (e.g. "Finalized") is logged on just one file, so we aggregate events across them. The
        # events keep their timestamps: the lifecycle is decided by the LAST relevant event, and
        # events of different files are only comparable by time. ----
        $fileNodes = @($obj.SelectNodes("*[local-name()='FILE']"))
        if ($fileNodes.Count -eq 0) {
            $script:Reconcile.NoFileObjects++
            Write-Log ("Record {0} (artifact {1}, '{2}') lists no FILE, so there is nothing to file and it has no row in the inventory." -f (Get-NodeText $obj 'OBJECTID'), $artifactNumber, $artifactTitle) 'WARN'
        }
        $objEvents = foreach ($fn in $fileNodes) {
            foreach ($a in $fn.SelectNodes("*[local-name()='AUDITRECORD']")) {
                [pscustomobject]@{
                    Event     = (Get-NodeText $a 'AUDITEVENT')
                    Timestamp = (Get-NodeText $a 'DATETIMESTAMP')
                }
            }
        }
        $lifecycleDetail = Get-EtmfLifecycleDetail @($objEvents)
        $lifecycle = $lifecycleDetail.Lifecycle
        Add-EtmfLifecycleAudit -Detail $lifecycleDetail
        $bucket    = Get-EtmfBucket $versionState $lifecycle
        $status    = Get-EtmfStatusToken $versionState $lifecycle $lifecycleDetail.Lock

        # ---- files (one humanized copy each), de-duplicated by CONTENTURL ----
        $seen = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::Ordinal)
        $fileNo = 0
        # Sets, not scans: a hostile record with thousands of repeated entries must stay linear.
        $seenFp    = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.HashSet[string]]' ([System.StringComparer]::Ordinal)
        $auditKeys = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.HashSet[string]]' ([System.StringComparer]::Ordinal)
        $sigKeys   = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.HashSet[string]]' ([System.StringComparer]::Ordinal)
        $script:Reconcile.Files += $fileNodes.Count
        foreach ($file in $fileNodes) {
            $fileNo++
            Add-EtmfUnreadFields -Element $file -Path 'FILE'
            foreach ($child in $file.ChildNodes) {
                if ($child.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
                if ($child.LocalName -in @('AUDITRECORD', 'SIGNATURE')) { Add-EtmfUnreadFields -Element $child -Path $child.LocalName }
            }
            $contentUrl = Get-NodeText $file 'CONTENTURL'
            # A record can list one file twice (the doubled signed file). Nothing is dropped on a
            # guess, and nothing the second entry says is lost:
            #   - identical in every detail                       -> filed once, counted;
            #   - same checksum and name, but more audit events or signatures
            #                                                       -> filed once; those events and
            #                                                          signatures are added to that row,
            #                                                          the row's Repeats column and the
            #                                                          run say so;
            #   - a different checksum or name                     -> a different file: kept as its own row.
            # Entries with no CONTENTURL have no path to be "the same file" at, so each stands alone.
            $fingerprint   = $file.InnerXml
            $integrityText = Get-NodeText $file 'INTEGRITY'
            $fileNameText  = Get-NodeText $file 'FILENAME'
            if ($contentUrl) {
                if ($seen.ContainsKey($contentUrl)) {
                    $entries = $seen[$contentUrl]
                    if ($seenFp[$contentUrl].Contains($fingerprint)) {
                        $script:Reconcile.ListedOnce++
                        Write-Log ("'{0}' is listed twice in one record with identical details; it is filed once." -f $fileNameText) 'INFO'
                        continue
                    }
                    $twin = @($entries | Where-Object { $_.Integrity -ceq $integrityText -and $_.FileName -ceq $fileNameText -and $null -ne $_.Doc })
                    if ($twin.Count -gt 0) {
                        $doc = $twin[0].Doc
                        $docKey = [string]$doc.FileNo
                        if (-not $auditKeys.ContainsKey($docKey)) {
                            $set = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
                            foreach ($x in @($doc.AuditTrail)) { [void]$set.Add((@($x.AuditId, $x.Timestamp, $x.User, $x.EntryType, $x.Event) -join [string][char]31)) }
                            $auditKeys[$docKey] = $set
                            $set2 = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
                            foreach ($x in @($doc.Signatures)) { [void]$set2.Add((@($x.Methodology, $x.UserId, $x.SignerName, $x.DateTime, $x.Reason) -join [string][char]31)) }
                            $sigKeys[$docKey] = $set2
                        }
                        $addedAudit = 0; $addedSig = 0
                        foreach ($a in $file.SelectNodes("*[local-name()='AUDITRECORD']")) {
                            $rec = [pscustomobject]@{
                                AuditId   = Get-NodeText $a 'AUDITID'
                                Timestamp = Get-NodeText $a 'DATETIMESTAMP'
                                User      = Get-NodeText $a 'USERREF'
                                EntryType = Get-NodeText $a 'AUDITENTRYTYPE'
                                Event     = Get-NodeText $a 'AUDITEVENT'
                            }
                            if ($auditKeys[$docKey].Add((@($rec.AuditId, $rec.Timestamp, $rec.User, $rec.EntryType, $rec.Event) -join [string][char]31))) { $doc.AuditTrail = @($doc.AuditTrail) + $rec; $addedAudit++ }
                        }
                        foreach ($sg in $file.SelectNodes("*[local-name()='SIGNATURE']")) {
                            $rec = [pscustomobject]@{
                                Methodology = Get-NodeText $sg 'SIGNATUREMETHODOLOGY'
                                UserId      = Get-NodeText $sg 'USERID'
                                SignerName  = Get-NodeText $sg 'SIGNATURENAME'
                                DateTime    = Get-NodeText $sg 'SIGNATUREDATETIME'
                                Reason      = Get-NodeText $sg 'SIGNATUREREASON'
                            }
                            if ($sigKeys[$docKey].Add((@($rec.Methodology, $rec.UserId, $rec.SignerName, $rec.DateTime, $rec.Reason) -join [string][char]31))) { $doc.Signatures = @($doc.Signatures) + $rec; $addedSig++ }
                        }
                        # remember this exact repeat too, so a further identical one is recognised
                        $entries.Add([pscustomobject]@{ Fingerprint = $fingerprint; Integrity = $integrityText; FileName = $fileNameText; Doc = $null })
                        [void]$seenFp[$contentUrl].Add($fingerprint)
                        if (($addedAudit + $addedSig) -eq 0) {
                            $script:Reconcile.ListedOnce++
                            Write-Log ("'{0}' is listed twice in one record; the second entry adds no audit events or signatures, so it is filed once." -f $fileNameText) 'INFO'
                        }
                        else {
                            $script:Reconcile.MergedInto++
                            $doc.AuditTrail = @(Sort-EtmfAuditTrail $doc.AuditTrail)
                            $doc.Repeats = (@($doc.Repeats, ("FILE entry {0} of this record is the same file; it added {1} audit event(s) and {2} signature(s)" -f $fileNo, $addedAudit, $addedSig)) | Where-Object { $_ }) -join '; '
                            Write-Log ("'{0}' is listed twice in one record with the same checksum and name but more audit events or signatures; it is filed once and the second entry's {1} audit event(s) and {2} signature(s) were added to that row." -f $fileNameText, $addedAudit, $addedSig) 'INFO'
                        }
                        continue
                    }
                    $script:Reconcile.KeptDistinct++
                    Write-Log ("'{0}' is listed twice in one record with the same content path but a different checksum or name; both entries are kept." -f $fileNameText) 'WARN'
                }
                else {
                    $seen[$contentUrl]   = [System.Collections.Generic.List[object]]::new()
                    $seenFp[$contentUrl] = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
                }
            }
            $entry = [pscustomobject]@{ Fingerprint = $fingerprint; Integrity = $integrityText; FileName = $fileNameText; Doc = $null }
            if ($contentUrl) { $seen[$contentUrl].Add($entry); [void]$seenFp[$contentUrl].Add($fingerprint) }

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
            $bindNote  = ''     # set when the file was bound by the trimmed-name second chance (a Note in the inventory)

            # We extracted this archive, so the index knows exactly which file this record is. The key
            # is the entry's VERBATIM archive path, so nothing is derived - there is nothing to collide
            # and no guess to make. Missing key, or '' (an ambiguous archive), means the document is
            # honestly recorded as a gap.
            $tail = Get-ContentUrlTail $contentUrl
            if ($null -eq $tail) {
                $gapReason = 'the CONTENTURL is malformed or points outside the archive'
            }
            else {
                $key = $tail
                if ($ContentPrefix) { $key = $ContentPrefix + '/' + $tail }
                $lookupKey = $key
                $mapped = $null
                $found = $ContentMap.TryGetValue($key, [ref]$mapped)
                $nearMiss = ''
                if (-not $found) {
                    # One narrow second chance, never a guess: the export's CONTENTURL may keep white space
                    # at the start of the file name that the zip entry has lost (see Get-EtmfTrimmedNameKey).
                    # Used only when exactly that one entry exists AND its bytes hash to the manifest's
                    # checksum - measured here, whatever -SkipIntegrityCheck says.
                    $trimKey = Get-EtmfTrimmedNameKey $key
                    if ($null -ne $trimKey) {
                        $trimMapped = $null
                        if (-not $ContentMap.TryGetValue($trimKey, [ref]$trimMapped)) { }      # no such entry: nothing to say
                        elseif ([string]::IsNullOrEmpty($trimMapped)) {
                            $nearMiss = 'an entry with this name without its leading/trailing white space exists, but two entries share that path, so it was not used'
                        }
                        elseif (-not (Test-Path -LiteralPath $trimMapped -PathType Leaf)) {
                            $nearMiss = 'an entry with this name without its leading/trailing white space exists, but it could not be extracted, so it was not used'
                        }
                        else {
                            $trimOk = $false
                            try { $trimOk = (-not [string]::IsNullOrWhiteSpace($integrity)) -and ((Get-FileMd5Base64 $trimMapped) -ceq $integrity) } catch { $trimOk = $false }
                            if ($trimOk) {
                                $found    = $true
                                $mapped   = $trimMapped
                                $key      = $trimKey
                                $lookupKey = $trimKey
                                $bindNote = 'The export lists this file name with leading white space that the zip entry does not have; matched without it, and the checksum equals the manifest''s'
                                Write-Log ("'{0}': {1}." -f (Get-NodeText $file 'FILENAME'), $bindNote) 'INFO'
                            }
                            else {
                                $nearMiss = 'an entry with this name without its leading/trailing white space exists, but its checksum is not the manifest''s (or none is stated), so it was not used'
                            }
                        }
                    }
                }
                if ($found) { [void]$script:Reconcile.Referenced.Add($key) }   # for the reconciliation
                if (-not $found) {
                    $gapReason = 'the archive contains no entry with this exact path'
                    if ($nearMiss) { $gapReason += '; ' + $nearMiss }
                }
                elseif ([string]::IsNullOrEmpty($mapped)) {
                    $gapReason = 'two or more archive entries share this exact path, so neither can be identified'
                }
                elseif (-not (Test-Path -LiteralPath $mapped -PathType Leaf)) {
                    $gapReason = 'the archive entry could not be extracted (see the extraction warning above)'
                }
                else { $diskPath = $mapped; $exists = $true }
            }

            $documents.Add([pscustomobject]@{
                Milestone      = $milestone
                AllMilestones  = @($milestone)   # every snapshot this exact file appeared in (merge)
                TransferId     = $transferId     # anchors this record's CONTENTURL, even once merged
                ArtifactNumber = $artifactNumber
                ObjectLevel    = Get-NodeText $obj 'OBJECTLEVEL'
                ManifestFile   = $ManifestFile.Name        # where this record sits in the export...
                ObjectNo       = $objectNo                 # ...which OBJECT (1-based) of that manifest
                FileNo         = $fileNo                   # ...and which FILE of that OBJECT
                SiteSystemId   = Get-NodeText $obj 'SITESYSTEMID'
                TitleSource    = $titleSource
                Repeats        = ''                        # filled when a repeated FILE entry added audit events/signatures to this row
                EffectiveCurrent      = ''                 # version resolution: True for the version treated as current ('' = not decided)
                LatestApprovedVersion = ''                 # ...the highest version of this document with a Finalized event
                NewerVersion          = ''                 # ...a newer version than the latest approved one, if any
                NewerVersionLifecycle = ''                 # ...and that version's own lifecycle (the export's words)
                Placement             = ''                 # which tree the file was written to (set when the files are copied)
                VersionResolution     = ''                 # ...in words, when resolution changed this row's bucket or status
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
                Lifecycle      = $lifecycle      # the export's own word for the record's last lifecycle event
                Lock           = $lifecycleDetail.Lock   # 'Document locked' / 'Document unlocked' / ''
                LifecycleNote  = (@($lifecycleDetail.Notes) -join '; ')   # set when the reading was a conservative one (tie, unusable time, ...)
                EarlierFileVersion = $false                               # set by Set-EtmfEarlierFileVersionStatus
                Bucket         = $bucket
                Status         = $status
                OriginalName   = Get-NodeText $file 'FILENAME'
                ContentUrl     = $contentUrl
                Integrity      = $integrity
                DiskPath       = $diskPath
                FileExists     = $exists
                GapReason      = $gapReason   # why it could not be bound (empty when it was)
                LookupKey      = $lookupKey   # the exact path/key we looked for
                BindNote       = $bindNote    # how the file was found, when not by the exact path
                SizeBytes      = if ($exists) { (Get-Item -LiteralPath $diskPath).Length } else { $null }      # no file, no size
                IntegrityOk    = $null   # filled by integrity pass (source)
                AuditTrail     = $auditTrail
                Signatures     = $signatures
                Metadata       = $metadata
                SupersededBy   = ''      # filled by Resolve-EtmfVersions when a newer version exists
                NewName        = ''      # filled by Phase 5 (copy/rename)
                RelPath        = ''      # filled by Phase 5 (path relative to output root)
                CopyIntegrityOk = $null  # filled by Phase 5 (verify the copy)
            })
            $entry.Doc = $documents[$documents.Count - 1]
        }
    }

    Set-EtmfEarlierFileVersionStatus -Documents $documents
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
        # order decide whether a document came out "Finalized / 01 Finalized" or "Awaiting review".
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
                    # would be the D6 wrong-bytes bug wearing a new hat. Checked always.
                    $donorOk = $false
                    try { $donorOk = ((Get-FileMd5Base64 $prev.DiskPath) -ceq $d.Integrity) } catch { $donorOk = $false }
                    if ($donorOk) {
                        $d.DiskPath    = $prev.DiskPath
                        $d.FileExists  = $true
                        $d.IntegrityOk = $true
                        $d.SizeBytes   = $prev.SizeBytes
                        $d.LookupKey   = $prev.LookupKey
                        $d.BindNote    = $prev.BindNote
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
#  Version resolution
# ----------------------------------------------------------------------------------------
# THE reported problem: "we have different version numbers of the same document, without knowing
# which one is the correct final and approved version."
#
# Two facts disagree and neither is wrong:
#   - OBJECTVERSIONSTATE is Viedoc's own and per SNAPSHOT: a v1 exported at one milestone says
#     "Current" in that manifest for ever, even when a later one carries a v2.
#   - Viedoc marks an approved v1 "Superseded" as soon as a v2 exists - even while v2 is only
#     Awaiting review and has never been approved. Filing v1 as superseded would then leave a
#     document that WAS approved with no approved version in the TMF, and a hand-over copy
#     (-FinalOnly) would lose the document altogether.
#
# So, inside one version group (DocumentId + artifact + country + site), the rule is:
#   * The EFFECTIVE CURRENT version is the highest version whose lifecycle is Finalized (a
#     version the export itself marks Obsolete is never chosen).
#   * Every LOWER version is filed under "03 Superseded" and records what replaced it - whether it
#     was approved or never was; Lifecycle says which (the status in the name is that lifecycle).
#   * A HIGHER version is left as the export has it. If it is not Finalized it stays "Not
#     finalized" and the approved version below it stays "Finalized"; the approved row says that a
#     newer version exists (NewerVersion, NewerVersionLifecycle).
#   * When Viedoc itself marks the effective current version "Superseded" only because that newer,
#     unapproved version exists, the version is kept as current (promoted), and the row says so in
#     VersionResolution. The export's own state stays in VersionState, untouched.
#   * A group in which no version is Finalized keeps the older rule: the highest version the
#     archive still calls Current stands, and the versions below it are superseded.
# Versions are compared segment-wise (v10 > v9), same-version records are one document's files and
# are left alone, a record without a DocumentId is never grouped, and the export's raw fields
# (VersionState, Lifecycle, Lock) are never overwritten - Bucket, Status and the columns above are
# the tool's derived values, each with its reason.
function Resolve-EtmfVersions {
    param($Model)

    $groups = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[object]]' ([System.StringComparer]::Ordinal)
    foreach ($d in $Model.Documents) {
        $key = Get-EtmfVersionGroupKey $d
        if ($null -eq $key) { continue }             # no document identity -> never grouped
        if (-not $groups.ContainsKey($key)) { $groups[$key] = [System.Collections.Generic.List[object]]::new() }
        $groups[$key].Add($d)
    }

    $demoted = 0; $promoted = 0; $resolved = 0; $demotedGroups = 0
    foreach ($key in @($groups.Keys)) {
        $members = $groups[$key]
        # More than one DISTINCT version? Several records of the same version are one document's files.
        $distinct = @($members | Group-Object Version).Count
        if ($distinct -lt 2) { continue }
        $resolved++
        $demotedBefore = $demoted

        $approved = @($members | Where-Object { $_.Lifecycle -ceq 'Finalized' -and $_.VersionState -cne 'Obsolete' })
        if ($approved.Count -gt 0) {
            $winner = $approved[0]
            foreach ($m in $approved) { if ((Compare-EtmfVersion $m.Version $winner.Version) -gt 0) { $winner = $m } }
            if (-not $winner.FileExists) {
                $samePresent = @($approved | Where-Object { $_.FileExists -and (Compare-EtmfVersion $_.Version $winner.Version) -eq 0 })
                if ($samePresent.Count -gt 0) { $winner = $samePresent[0] }
            }
        }
        else {
            # Nothing approved: the highest version the archive still calls Current (else the highest).
            $currents = @($members | Where-Object { $_.VersionState -ceq 'Current' })
            $pool     = @(if ($currents.Count -gt 0) { $currents } else { $members })
            $winner   = $pool[0]
            foreach ($m in $pool) { if ((Compare-EtmfVersion $m.Version $winner.Version) -gt 0) { $winner = $m } }
        }
        $hasApproved = ($approved.Count -gt 0)

        # Versions above the winner (only possible when the winner is the highest Finalized one).
        # A newer version the export itself marks Obsolete replaces nothing.
        $newer = @($members | Where-Object { $_.VersionState -cne 'Obsolete' -and (Compare-EtmfVersion $_.Version $winner.Version) -gt 0 })
        $newest = $null
        foreach ($m in $newer) { if ($null -eq $newest -or (Compare-EtmfVersion $m.Version $newest.Version) -gt 0) { $newest = $m } }

        foreach ($m in $members) {
            if ($hasApproved) { $m.LatestApprovedVersion = [string]$winner.Version }
            if ((Compare-EtmfVersion $m.Version $winner.Version) -ne 0) { continue }
            if ($hasApproved -and $m.Lifecycle -ceq 'Finalized') {
                if ($null -ne $newest) {
                    $m.NewerVersion          = [string]$newest.Version
                    $m.NewerVersionLifecycle = [string]$newest.Lifecycle
                    if ($m.VersionState -ceq 'Superseded' -and $m.Bucket -ceq '03 Superseded') {
                        $m.Bucket = '01 Finalized'
                        $m.Status = Get-EtmfStatusToken '' $m.Lifecycle $m.Lock
                        $m.VersionResolution = ("Kept as the current approved version: the export marks it 'Superseded', but the newer version v{0} {1}." -f `
                            $newest.Version, $(if ($newest.Lifecycle) { "is '" + $newest.Lifecycle + "', not Finalized" } else { 'has no lifecycle event' }))
                        $promoted++
                    }
                }
                if ($m.Bucket -ceq '01 Finalized') { $m.EffectiveCurrent = 'True' }
                elseif (-not $m.VersionResolution) {
                    $m.VersionResolution = ("Finalized, but the export marks it '{0}' and no newer version is in this export to explain it; left as the export has it." -f $m.VersionState)
                }
            }
        }

        if (-not $winner.FileExists) {
            Write-Log ("The current version of '{0}' (v{1}) has no file in the archive; older versions are still filed as superseded." -f `
                $winner.OriginalName, $winner.Version) 'WARN'
        }
        $sharedChecksum = @($members | Group-Object Integrity | Where-Object { $_.Count -gt 1 -and $_.Name })
        if ($sharedChecksum.Count -gt 0) {
            Write-Log ("'{0}': versions {1} share one file and checksum - normal for this export format (the version history of one document), recorded here because it is also what a version RELABELLED between exports would look like." -f `
                $winner.OriginalName, (@($members | ForEach-Object { 'v' + $_.Version } | Select-Object -Unique) -join ', ')) 'INFO'
        }
        # Only what the export gave: a winner with no lifecycle event has no status to quote.
        $replacedBy = 'v{0} ({1})' -f $winner.Version, ((@($winner.Status, $winner.Milestone) | Where-Object { $_ }) -join ', ')
        foreach ($m in $members) {
            if ((Compare-EtmfVersion $m.Version $winner.Version) -ge 0) { continue }
            $m.SupersededBy = $replacedBy
            if ($m.Bucket -ne '03 Superseded') {
                # It called itself Current in its own snapshot; a newer version says otherwise.
                $m.Bucket = '03 Superseded'
                # an earlier file version with no lifecycle of its own keeps the word it was given
                $word = if ($m.EarlierFileVersion -and -not $m.Lifecycle) { 'Unpublished' } else { $m.Lifecycle }
                $m.Status = Get-EtmfStatusToken 'Superseded' $word $m.Lock
                $m.VersionResolution = ("Moved to Superseded by the tool: v{0} of this document is newer{1}; the export marks this row '{2}'." -f `
                    $winner.Version, $(if ($hasApproved) { ' and Finalized' } else { '' }), $m.VersionState)
                $demoted++
            }
        }
        if ($demoted -gt $demotedBefore) { $demotedGroups++ }
    }

    # A Finalized row the export marks Superseded, replaced by nothing visible in this export (alone
    # in its group, or its successor is elsewhere / has no file): the tool cannot say what replaced
    # it, so it stays as the export has it - and says so, because such a row is not in 01 Finalized
    # and not in a -FinalOnly copy.
    $unexplained = 0
    foreach ($d in $Model.Documents) {
        if ($d.Lifecycle -ceq 'Finalized' -and $d.VersionState -ceq 'Superseded' -and $d.Bucket -ceq '03 Superseded' -and -not $d.SupersededBy) {
            $unexplained++
            if (-not $d.VersionResolution) {
                $d.VersionResolution = "Finalized, but the export marks it 'Superseded' and no newer version is in this export to explain it; left as the export has it."
            }
        }
    }

    return [pscustomobject]@{ Demoted = $demoted; Promoted = $promoted; Groups = $demotedGroups; Unexplained = $unexplained }
}

# The EARLIER FILE VERSIONS of a record. A record (one document version) can hold several files:
# each time a new file is uploaded to it - before approval, or after publishing and before approval -
# the previous file stays in the export as an extra FILE of the same record. Only the file that
# carries the record's Finalized event is the approved one; every other file of that record has no
# Finalized event of its own, so it was never approved. Not matched: a signed file (it is the signed
# rendering of the approved one), or a file with the same checksum as the approved one (the same
# bytes). Returns {Document; ReplacedBy (names of the approved files)}; it changes nothing.
function Get-EtmfEarlierFileVersions {
    param($Documents)
    $found = [System.Collections.Generic.List[object]]::new()
    $byRecord = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[object]]' ([System.StringComparer]::Ordinal)
    foreach ($d in $Documents) {
        if ([string]::IsNullOrWhiteSpace($d.ManifestFile)) { continue }
        $rk = '{0}|{1}|{2}' -f $d.TransferId, $d.ManifestFile, $d.ObjectNo
        if (-not $byRecord.ContainsKey($rk)) { $byRecord[$rk] = [System.Collections.Generic.List[object]]::new() }
        $byRecord[$rk].Add($d)
    }
    foreach ($rk in @($byRecord.Keys)) {
        $files = $byRecord[$rk]
        if ($files.Count -lt 2) { continue }
        $carriers = @($files | Where-Object { @($_.AuditTrail | Where-Object { (Get-EtmfEventHead $_.Event) -ceq 'Finalized' }).Count -gt 0 })
        if ($carriers.Count -eq 0) { continue }
        foreach ($f in $files) {
            if ($carriers -contains $f) { continue }
            if (@($f.Signatures).Count -gt 0) { continue }
            $sameBytes = @($carriers | Where-Object { $_.Integrity -and $_.Integrity -ceq $f.Integrity })
            if ($sameBytes.Count -gt 0) { continue }
            $found.Add([pscustomobject]@{ Document = $f; ReplacedBy = @($carriers | ForEach-Object { $_.OriginalName }) })
        }
    }
    return @($found)
}

# An earlier file version was never approved: it has no Finalized event of its own. Its lifecycle is
# what ITS OWN audit trail says (the last Awaiting review / Flagged by QC Comment event, or none),
# so it lands in '02 Not finalized' and - under -SeparateNotFinalized - in Not finalized/, and is
# left out of -FinalOnly. The record's approved file is untouched. Nothing is invented: the export's
# words for this file, and a note saying why it is not read with the record.
function Set-EtmfEarlierFileVersionStatus {
    param($Documents)
    $earlier = @(Get-EtmfEarlierFileVersions -Documents $Documents)
    foreach ($r in $earlier) {
        $d = $r.Document
        $d.EarlierFileVersion = $true
        $own = Get-EtmfLifecycleDetail @($d.AuditTrail)
        $d.Lifecycle     = $own.Lifecycle
        $d.Lock          = ''
        $d.Bucket        = Get-EtmfBucket $d.VersionState $d.Lifecycle
        # No event of its own shows it was ever published or finalized: its status is Unpublished (a
        # labelled derivation - Viedoc's status for a record that was not published).
        $word = if ($own.Lifecycle) { $own.Lifecycle } else { 'Unpublished' }
        $d.Status        = Get-EtmfStatusToken $d.VersionState $word ''
        $why = 'this is an earlier file version of the record: it has no Finalized event of its own, while another file of the record carries it (' + (@($r.ReplacedBy | Select-Object -First 2) -join ', ') + '), so it is not treated as approved'
        if (-not $own.Lifecycle) { $why += '; nothing in its own history shows it was ever published or finalized, so its status is Unpublished' }
        # keep what the record-level reading already said (a tie, an unusable time, ...) and add this
        $d.LifecycleNote = (@($d.LifecycleNote, $why) | Where-Object { $_ }) -join '; '
    }
    if ($earlier.Count -gt 0) {
        Write-Log ("{0} file(s) are earlier file versions of a record (no Finalized event of their own while another file of the record has one); they are not treated as approved - see LifecycleNote in inventory.csv." -f $earlier.Count) 'INFO'
    }
}

# Reduce a model to the documents a hand-over copy may contain: the rows filed under '01 Finalized' -
# the latest Finalized version of a document, not replaced by a newer Finalized one. FAIL CLOSED:
# a row that is not positively '01 Finalized' is left out, whatever else is true of it, and every row
# left out is itemised (Items -> withheld.csv) with the reason, because a partial TMF that does not
# say what it left out is the problem this tool exists to fix.
function Select-EtmfFinalDocuments {
    param($Model)

    $kept = [System.Collections.Generic.List[object]]::new()
    $items = [System.Collections.Generic.List[object]]::new()
    $withheld = [pscustomobject]@{ Superseded = 0; InProgress = 0; NoFile = 0; NewerPending = 0; IntegrityProblems = 0; Items = $items }
    foreach ($d in $Model.Documents) {
        if ($d.Bucket -ne '01 Finalized') {
            if ($d.Bucket -eq '03 Superseded') {
                $withheld.Superseded = $withheld.Superseded + 1
                $reason = if ($d.SupersededBy) { 'Superseded by ' + $d.SupersededBy } else { 'Superseded (the export marks this version ' + $d.VersionState + $(if ($d.Lifecycle) { '; its last lifecycle event is ' + $d.Lifecycle } else { '' }) + ')' }
            }
            else {
                $withheld.InProgress = $withheld.InProgress + 1
                $reason = 'Not Finalized: ' + $(if ($d.Lifecycle) { 'last lifecycle event is ' + $d.Lifecycle } else { 'no lifecycle event in the export' })
                if ($d.LifecycleNote) { $reason = $reason + ' - ' + $d.LifecycleNote }
            }
            $items.Add([pscustomobject][ordered]@{
                ArtifactNumber = $d.ArtifactNumber; ArtifactTitle = $d.ArtifactTitle; Country = $d.Country; SiteId = $d.SiteId
                Version = $d.Version; VersionState = $d.VersionState; Lifecycle = $d.Lifecycle; Lock = $d.Lock
                Status = $d.Status; Reason = $reason; SupersededBy = $d.SupersededBy
                LatestApprovedVersion = $d.LatestApprovedVersion; OriginalName = $d.OriginalName
                Manifest = $d.ManifestFile; ObjectNo = $d.ObjectNo; FileNo = $d.FileNo
            })
            continue
        }
        if (-not $d.FileExists)              { $withheld.NoFile       = $withheld.NoFile + 1 }
        if ($d.NewerVersion)                 { $withheld.NewerPending = $withheld.NewerPending + 1 }
        if ($d.FileExists -and ($d.IntegrityOk -eq $false)) { $withheld.IntegrityProblems = $withheld.IntegrityProblems + 1 }
        $kept.Add($d)
    }
    $Model.Documents = $kept
    return $withheld
}

# ----------------------------------------------------------------------------------------
#  Integrity verification (base64 MD5)
# ----------------------------------------------------------------------------------------
# The integrity tally of a set of documents - what a viewer header and a summary should count. It is
# computed from the documents a view actually holds, so a -FinalOnly copy does not claim to have
# checked the documents it left out.
function Get-EtmfIntegritySummary {
    param($Documents)
    $docs = @($Documents)
    return [pscustomobject]@{
        Ok         = @($docs | Where-Object { $_.FileExists -and $_.IntegrityOk -eq $true }).Count
        Bad        = @($docs | Where-Object { $_.FileExists -and $_.IntegrityOk -eq $false }).Count
        Missing    = @($docs | Where-Object { -not $_.FileExists }).Count
        Unverified = @($docs | Where-Object { $_.FileExists -and $null -eq $_.IntegrityOk }).Count
    }
}

function Invoke-IntegrityCheck {
    param([System.Collections.Generic.List[object]]$Documents)
    $ok = 0; $bad = 0; $missing = 0; $unverified = 0
    foreach ($d in $Documents) {
        if (-not $d.FileExists) { $missing++; $d.IntegrityOk = $null; continue }   # no file, so nothing was checked
        if ([string]::IsNullOrWhiteSpace($d.Integrity)) {
            # No checksum stated is not corruption, and not "verified" either: the result stays empty
            # (neither True nor False) and is not counted as a mismatch.
            $d.IntegrityOk = $null
            $unverified++
            Write-Log "No checksum stated in the manifest for '$($d.OriginalName)' - the copy could not be verified." 'WARN'
            continue
        }
        try { $actual = Get-FileMd5Base64 $d.DiskPath }
        catch {
            # One unreadable file must not stop the whole run; it is unverified, and says why.
            $d.IntegrityOk = $null
            $unverified++
            Write-Log "Could not read '$($d.OriginalName)' to verify it: $($_.Exception.Message)" 'WARN'
            continue
        }
        $match  = ($actual -ceq $d.Integrity)   # base64 is case-sensitive
        $d.IntegrityOk = $match
        if ($match) { $ok++ } else { $bad++; Write-Log "Integrity MISMATCH: $($d.OriginalName)" 'WARN' }
    }
    return [pscustomobject]@{ Ok = $ok; Bad = $bad; Missing = $missing; Unverified = $unverified }
}


# Human-friendly destination filename.
# The folder already encodes Zone/Section/Artifact, so the file name only needs the document's
# own name + version + status -- this keeps the full path within the Windows MAX_PATH (260) limit.
# -MaxLength (when > 0) caps the whole file name, truncating the document-name part to fit.
function Get-HumanFileName {
    param($Doc, [int]$MaxLength = 0)
    # Sanitize BEFORE any path API sees it: OriginalName is the verbatim manifest <FILENAME>, and
    # on .NET Framework (5.1) GetExtension/GetFileNameWithoutExtension throw "Illegal characters
    # in path." for " < > | and control chars - the same crash this tool now handles at extraction.
    # Doing it here also stops an unsafe extension surviving into the destination via "+ $ext".
    $safeOriginal = ConvertTo-SafeName $Doc.OriginalName
    $ext = [System.IO.Path]::GetExtension($safeOriginal)
    $stem   = [System.IO.Path]::GetFileNameWithoutExtension($safeOriginal)
    # No status in the export means no status in the name: nothing is invented to fill the brackets.
    $suffix = if ($Doc.Status) { " v{0} [{1}]" -f $Doc.Version, $Doc.Status } else { " v{0}" -f $Doc.Version }
    if ($MaxLength -gt 0) {
        $room = $MaxLength - $suffix.Length - $ext.Length
        if ($room -lt 8) { $room = 8 }
        if ($stem.Length -gt $room) {
            # Keep the START **and the END** of the name, with '~' marking the cut. Clinical file
            # names are distinguished by their tail ("...-01" / "...-02", a date, a version), so
            # head-only truncation collapses different documents onto one name; the collision guard
            # then files them as " (2)", which reads to a reviewer like a duplicate of the first.
            [int]$tailRoom = [Math]::Min(12, [int][Math]::Floor(($room - 1) / 2))
            # never cut between the two halves of a surrogate pair (an emoji): the half left behind is
            # not a character, and the file name and the inventory would then disagree about it
            $headLen = $room - 1 - $tailRoom
            if ($headLen -gt 0 -and [char]::IsHighSurrogate($stem[$headLen - 1])) { $headLen-- }
            $tailStart = $stem.Length - $tailRoom
            if ($tailStart -lt $stem.Length -and [char]::IsLowSurrogate($stem[$tailStart])) { $tailStart++ }
            $head = $stem.Substring(0, $headLen).TrimEnd()
            $stem = $head + '~' + $stem.Substring($tailStart)
        }
    }
    return (ConvertTo-SafeName ($stem + $suffix)) + $ext
}

# How many bytes the output will add on its volume. Several records can share ONE extracted file (every
# version of a document has the same file); each record still gets its own output file. When the
# extracted files are moved rather than copied AND sit on the output's volume, a move is a rename and
# costs nothing, so only the extra copies for shared files count; otherwise every file is written in full.
function Get-EtmfOutputBytes {
    param($Documents, [bool]$MovesAreFree)
    $total  = 0L
    $unique = 0L
    $seen   = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($d in $Documents) {
        if (-not $d.FileExists -or $null -eq $d.SizeBytes) { continue }
        $total += [long]$d.SizeBytes
        if ($seen.Add([string]$d.DiskPath)) { $unique += [long]$d.SizeBytes }
    }
    if ($MovesAreFree) { return ($total - $unique) }
    return $total
}

# Copy every document into  <out>/Documents/<Zone>/<Section>/<Artifact>/<Scope>/<file>,
# re-verify each copy, and return the inventory rows. One opt-in layout:
#   -SeparateNotFinalized  records whose last lifecycle event is not Finalized go to a sibling tree,
#                          <out>/Not finalized/<Zone>/<Section>/<Artifact>/..., mirroring the first.
# Placement is decided by the record's OWN last lifecycle event, never by the tool's version
# resolution, so a superseded version that was approved stays with the approved records.
function Invoke-Humanize {
    param(
        $Model,
        [string]$OutputRoot,
        [switch]$VerifyCopy,
        [switch]$SeparateNotFinalized,
        # The source files are the run's own temporary extraction, deleted afterwards: move each one into
        # the tree instead of copying it, so the output does not need a second copy's worth of disk.
        [switch]$MoveSources
    )
    # A trailing separator would make the root look one character longer than it is, and the file
    # names are budgeted from the root's length, so the same export would get different names.
    if ($OutputRoot.Length -gt 1 -and $OutputRoot -notmatch '^[A-Za-z]:[\\/]$') { $OutputRoot = $OutputRoot.TrimEnd('\', '/') }
    $treeRoot = Join-Path $OutputRoot 'Documents'
    New-Item -ItemType Directory -Path $treeRoot -Force | Out-Null
    $notFinalizedRoot = Join-Path $OutputRoot 'Not finalized'

    $used      = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $nextSuffix = New-Object System.Collections.Hashtable ([System.StringComparer]::OrdinalIgnoreCase)
    $inventory = [System.Collections.Generic.List[object]]::new()
    # A source file shared by several records (every version of a document has the same file) is copied for
    # all but the LAST record that uses it, and moved for that one. Counted up front, so "last" is known.
    $usesLeft = New-Object 'System.Collections.Generic.Dictionary[string,int]' ([System.StringComparer]::OrdinalIgnoreCase)
    if ($MoveSources) {
        foreach ($u in $Model.Documents) {
            if ($u.FileExists -and $u.DiskPath) {
                if ($usesLeft.ContainsKey($u.DiskPath)) { $usesLeft[$u.DiskPath] = $usesLeft[$u.DiskPath] + 1 } else { $usesLeft[$u.DiskPath] = 1 }
            }
        }
    }
    $moved = 0
    $copied = 0; $missing = 0; $copyBad = 0; $shortened = 0; $tooLong = 0; $writeFailed = 0
    $targetMaxPath = 250   # keep generated full paths comfortably under the Windows MAX_PATH (260)

    foreach ($d in $Model.Documents) {
        # Anchor on the record's OWN TransferID: after a merge the documents come from several
        # snapshots, so the model's batch is not the one this CONTENTURL was written against.
        $pp   = Get-ContentPathParts $d.ContentUrl $d.TransferId
        $note = ''
        # Count this record off its source file NOW, before anything below can throw, so a record that
        # fails early does not keep the file from moving with the last record that uses it.
        $lastUse = $false
        if ($MoveSources -and $d.FileExists -and $d.DiskPath -and $usesLeft.ContainsKey($d.DiskPath)) {
            $usesLeft[$d.DiskPath] = $usesLeft[$d.DiskPath] - 1
            $lastUse = ($usesLeft[$d.DiskPath] -le 0)
        }
        $d.Placement = ''      # set only once the file is actually written there
        $intendedPlacement = if ($SeparateNotFinalized -and $d.Lifecycle -cne 'Finalized') { 'Not finalized' } else { 'Documents' }

        if ($d.FileExists) {
            $note = $d.BindNote      # empty unless the file was found by the trimmed-name second chance
          try {
            # The folder segments below the output root: Zone / Section / Artifact [/ Country / Site].
            $segments = @(
                (ConvertTo-SafeName (Get-PrettyTmfSegment $pp.Zone)),
                (ConvertTo-SafeName (Get-PrettyTmfSegment $pp.Section)),
                (ConvertTo-SafeName ("{0} {1}" -f $d.ArtifactNumber, $d.ArtifactTitle))
            )
            foreach ($s in $pp.Scope) { $segments += (ConvertTo-SafeName (Get-PrettyScope $s)) }
            $root = if ($intendedPlacement -ceq 'Not finalized') { $notFinalizedRoot } else { $treeRoot }
            $folder = $root
            foreach ($seg in $segments) { $folder = Join-Path $folder $seg }
            New-Item -ItemType Directory -Path $folder -Force | Out-Null

            # Budget the file name from the artifact folder's depth alone: the Zone/Section/Artifact
            # [/Country/Site] part, plus a fixed allowance for the output root and one for the tree
            # folder (the longer of the two names, 'Not finalized'). Neither the root nor the tree the
            # record sits in enters the sum, so one export gets the same file
            # names wherever it is written and however it is laid out (a budget taken from the real
            # root made 38 of 57 names differ between a short and a long -OutputPath). A root longer
            # than the allowance can still push a path over the Windows limit; that is counted and
            # warned about at the end instead of silently changing names.
            $relFolder = [Math]::Max(0, $folder.Length - $root.Length)
            $budget   = $targetMaxPath - ($script:PathBudgetRoot + 1 + $script:PathBudgetTree + $relFolder + 1)
            if ($budget -lt 30) { $budget = 30 }
            $fileName = Get-HumanFileName $d -MaxLength $budget
            if ($fileName -ne (Get-HumanFileName $d)) { $shortened++ }
            $dest     = Join-Path $folder $fileName
            if ($used.Contains($dest) -or (Test-Path -LiteralPath $dest)) {    # collision guard
                $stem = [IO.Path]::GetFileNameWithoutExtension($fileName)
                $ext  = [IO.Path]::GetExtension($fileName)
                $baseKey = $stem + '|' + $folder + '|' + $ext
                $n = if ($nextSuffix.ContainsKey($baseKey)) { $nextSuffix[$baseKey] } else { 2 }       # resume: never re-walk 2, 3, 4 ...
                do   { $dest = Join-Path $folder ("{0} ({1}){2}" -f $stem, $n, $ext); $n++ }
                while ($used.Contains($dest) -or (Test-Path -LiteralPath $dest))
                $nextSuffix[$baseKey] = $n
            }
            [void]$used.Add($dest)
            if ($dest.Length -ge 260) { $tooLong++ }

            # Move only on the last use of the file; if that record is the one that fails, the file is copied for
            # the others and simply stays in temp (the safe direction).
            $moveNow = ($MoveSources -and $lastUse)
            if ($moveNow) { Move-Item -LiteralPath $d.DiskPath -Destination $dest -Force; $moved++ }
            else          { Copy-Item -LiteralPath $d.DiskPath -Destination $dest -Force }
            $copyOk = $null
            if ($VerifyCopy) {
                if (-not [string]::IsNullOrWhiteSpace($d.Integrity)) {
                    $copyOk = ((Get-FileMd5Base64 $dest) -ceq $d.Integrity)   # base64 is case-sensitive
                    # Name the file we actually WROTE: the collision guard above may have suffixed it,
                    # and a QA reviewer reconciling the log against the tree must land on the real one.
                    if (-not $copyOk) { $copyBad++; Write-Log "Checksum MISMATCH after writing: $(Split-Path $dest -Leaf)" 'WARN' }
                }   # no checksum stated: nothing to compare against, so the copy stays unverified (empty), not a mismatch
            }
            $copied++

            $d.NewName         = Split-Path $dest -Leaf
            $d.RelPath         = ($dest.Substring($OutputRoot.Length)).TrimStart('/', '\') -replace '\\', '/'
            $d.CopyIntegrityOk = $copyOk
            $d.Placement       = $intendedPlacement
          }
          catch {
            # Windows PowerShell 5.1 without long-path support throws on a path of 260+ characters.
            # One file that cannot be written must not abort the run and lose every row after it.
            $writeFailed++
            Write-Log ("Could not write '{0}': {1} (choose a shorter -OutputPath if the path is too long)" -f $d.OriginalName, $_.Exception.Message) 'WARN'
            $note = 'NOT WRITTEN: ' + $_.Exception.Message
            $d.NewName = ''; $d.RelPath = ''; $d.CopyIntegrityOk = $null
          }
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
            Placement      = $d.Placement          # the tree the file sits in: Documents or Not finalized
            Status         = $d.Status
            SupersededBy   = $d.SupersededBy      # the version that replaced this one, if any
            EffectiveCurrent      = $d.EffectiveCurrent       # True: the version treated as the current approved one
            LatestApprovedVersion = $d.LatestApprovedVersion  # highest version of this document with a Finalized event
            NewerVersion          = $d.NewerVersion           # a newer version than that, if one exists...
            NewerVersionLifecycle = $d.NewerVersionLifecycle  # ...and its own lifecycle, in the export's words
            VersionResolution     = $d.VersionResolution      # why the tool changed Bucket/Status, when it did
            DocumentId     = $d.DocumentId
            ObjectId       = $d.ObjectId
            ArtifactDate   = $d.ArtifactDate
            DateDescription = $d.DateDesc            # what ArtifactDate is (Version date, Report Date, ...)
            SubArtifact    = $d.SubArtifact
            Translation    = $d.Translation
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
            LifecycleNote  = $d.LifecycleNote     # why the lifecycle was read conservatively, when it was
            Lock           = $d.Lock               # last lock event the export names ('Document locked' / 'Document unlocked')
            # Where this row sits in the export, and fields the export has that the columns above omit:
            SiteId         = $d.SiteId
            SiteSystemId   = $d.SiteSystemId
            UniqueId       = $d.UniqueId
            ObjectCopy     = $d.IsCopy
            MetadataVersion = $d.MetaVersion       # the user's own version label (METADATA Version); OBJECTVERSION is Version
            TitleSource    = $d.TitleSource        # which export field ArtifactTitle came from
            Manifest       = $d.ManifestFile
            ObjectNo       = $d.ObjectNo           # 1-based position of the OBJECT in that manifest
            FileNo         = $d.FileNo             # 1-based position of the FILE in that OBJECT
            ContentUrl     = $d.ContentUrl
            Repeats        = $d.Repeats            # set when a repeated FILE entry of the record added audit events or signatures to this row
            ToolVersion    = $script:ToolVersion   # so a CSV separated from its folder still says what made it
        })
    }

    if ($moved -gt 0) {
        $sharedCopies = $copied - $moved
        $tail = if ($sharedCopies -gt 0) { " (a further $sharedCopies copied: records that share a file get a copy each, and the last one moves it)" } else { '' }
        Write-Log ("{0} file(s) were moved out of the temporary folder instead of copied{1}." -f $moved, $tail) 'INFO'
    }
    if ($tooLong -gt 0) {
        Write-Log ("{0} file(s) have a full path of 260 characters or more, which Windows Explorer and many programs cannot open. Choose a shorter -OutputPath (the file names themselves were not changed for this)." -f $tooLong) 'WARN'
    }
    if ($shortened -gt 0) {
        Write-Log ("{0} long file name(s) shortened (the ~ marks where text was cut) to stay within the Windows 260-char path limit; originals are in inventory.csv." -f $shortened) 'INFO'
    }

    return [pscustomobject]@{
        Inventory = $inventory; TreeRoot = $treeRoot
        Copied = $copied; Moved = $moved; Missing = $missing; CopyBad = $copyBad; WriteFailed = $writeFailed
    }
}


function New-HtmlViewer {
    param($Model, [string]$OutputRoot, $Integrity, [switch]$FinalOnlyCopy, [switch]$SeparatedCopy)

    $batch = $Model.Batch


    # ---- per-document records for the front-end ----
    $records = foreach ($d in $Model.Documents) {
        # Anchor on the record's own TransferID - after a merge the batch belongs to one snapshot.
        $pp    = Get-ContentPathParts $d.ContentUrl $d.TransferId
        $scope = @($pp.Scope | ForEach-Object { Get-PrettyScope $_ })
        $iso   = ConvertFrom-EmsDate $d.ArtifactDate
        [pscustomobject]@{
            milestone    = $d.Milestone
            milestones   = @($d.AllMilestones)      # every snapshot it appeared in (viewer filters on this)
            artNo        = $d.ArtifactNumber
            zone         = (Get-PrettyTmfSegment $pp.Zone)
            section      = (Get-PrettyTmfSegment $pp.Section)
            artifact     = ("{0} {1}" -f $d.ArtifactNumber, $d.ArtifactTitle)
            title        = $d.ArtifactTitle
            level        = $d.ObjectLevel
            siteId       = $d.SiteId
            country      = $d.Country                 # the export's COUNTRYID, as written
            organization = $d.Organization            # the export's ORGANIZATIONNAME, as written (scope below is the folder name)
            scope        = ($scope -join ' / ')
            person       = $d.Person
            version      = $d.Version
            versionState = $d.VersionState
            lifecycle    = $d.Lifecycle
            lock         = $d.Lock
            bucket       = $d.Bucket
            placement    = $d.Placement
            status       = $d.Status
            supersededBy = $d.SupersededBy
            latestApproved = $d.LatestApprovedVersion
            newerVersion = $d.NewerVersion
            newerLifecycle = $d.NewerVersionLifecycle
            resolution   = $d.VersionResolution
            lifecycleNote = $d.LifecycleNote
            restricted   = ($d.Restricted -ieq 'Yes')
            restrictedRaw = $d.Restricted
            language     = $d.Language
            translation  = $d.Translation
            expiry       = $d.ExpiryDate
            date         = $d.ArtifactDate
            dateIso      = if ($iso) { $iso.ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture) } else { '' }
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
            integrityStated = [bool]$d.Integrity
            sourceOk     = [bool]$d.IntegrityOk
            copyOk       = [bool]$d.CopyIntegrityOk
            repeats      = $d.Repeats
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
        separated      = [bool]$SeparatedCopy
        generated      = (Get-Date).ToString('yyyy-MM-dd HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)
        total          = $Model.Documents.Count
        final          = @($Model.Documents | Where-Object { $_.Bucket -eq '01 Finalized' }).Count
        inProgress     = @($Model.Documents | Where-Object { $_.Bucket -eq '02 Not finalized' }).Count
        superseded     = @($Model.Documents | Where-Object { $_.Bucket -eq '03 Superseded' }).Count
        signed         = @($Model.Documents | Where-Object { $_.Signatures.Count -gt 0 }).Count
        missing        = @($Model.Documents | Where-Object { -not $_.FileExists }).Count
        integrityOk    = $intOk
        integrityTotal = $intTotal
        integrityUnverified = $(if ($Integrity) { $Integrity.Unverified } else { 0 })
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
    # Both tokens in ONE pass: replacing them one after the other would also replace a token that
    # sits inside the first replacement's manifest text (a study id containing /*__DATA__*/).
    $html = [regex]::Replace($template, '/\*__(META|DATA)__\*/', [System.Text.RegularExpressions.MatchEvaluator]{
        param($m)
        if ($m.Groups[1].Value -ceq 'META') { $metaJson } else { $dataJson }
    })

    $indexPath = Join-Path $OutputRoot 'index.html'
    # write UTF-8 (no BOM) so the browser reads it cleanly
    [System.IO.File]::WriteAllText($indexPath, $html, (New-Object System.Text.UTF8Encoding($false)))
    return $indexPath
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
    Write-Host "    * tidy folders you can browse - one per TMF artifact, with readable" -ForegroundColor $g
    Write-Host "      file names that carry the status the export gives each document," -ForegroundColor $g
    Write-Host "      and" -ForegroundColor $g
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
    Write-Host "  On a locked-down PC, or sharing this with a colleague?  This longer" -ForegroundColor $g
    Write-Host "  form starts on any Windows machine and lets this one script run" -ForegroundColor $g
    Write-Host "  even if Windows would otherwise block it as unsigned:" -ForegroundColor $g
    Write-Host '      powershell -ExecutionPolicy Bypass -File .\Convert-EtmfArchive.ps1 -ArchivePath "...your.zip"' -ForegroundColor $gr
    Write-Host ""
    Write-Host "  To hand a client only the finalized documents, add -FinalOnly; to" -ForegroundColor $g
    Write-Host "  keep what is not finalized in its own folder, add -SeparateNotFinalized." -ForegroundColor $g
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
# The layout note every normal run writes: what the folders are, what the statuses mean in the export's
# words, and - when records were put in a second tree - that BOTH trees are the trial master file.
# The milestone line of a note: every milestone of a merged tree, not only the newest.
function Get-EtmfMilestoneLine {
    param($Batch)
    $ms = @($Batch.Milestones | Where-Object { $_ })
    if ($ms.Count -gt 1) { return ("  Milestones    : {0} ({1})" -f ($ms -join '; '), $ms.Count) }
    return ("  Milestone     : {0}" -f $Batch.Milestone)
}

function Write-EtmfLayoutNote {
    param([string]$Path, $Batch, $Documents, [string]$SourceInput, [switch]$SeparateNotFinalized)

    $docs = @($Documents)
    $nFinal = @($docs | Where-Object { $_.Bucket -ceq '01 Finalized' }).Count
    $nNot   = @($docs | Where-Object { $_.Bucket -ceq '02 Not finalized' }).Count
    $nSup   = @($docs | Where-Object { $_.Bucket -ceq '03 Superseded' }).Count
    $nSepNot = @($docs | Where-Object { $_.Placement -ceq 'Not finalized' }).Count
    $nDocs   = @($docs | Where-Object { $_.Placement -ceq 'Documents' }).Count
    $lines = [System.Collections.Generic.List[string]]::new()
    $add = { param([string]$Text = '') $lines.Add($Text) }

    & $add "ABOUT THIS FOLDER"
    & $add "================="
    & $add ""
    & $add ("The trial master file for {0}, taken from the eTMF export named below." -f $Batch.SpecificationId)
    & $add ""
    & $add ("  Source export : {0}" -f (Split-Path $SourceInput -Leaf))
    & $add (Get-EtmfMilestoneLine $Batch)
    & $add ("  Prepared      : {0}" -f (Get-Date).ToString('yyyy-MM-dd HH:mm', [System.Globalization.CultureInfo]::InvariantCulture))
    & $add ("  Prepared with : eTMF Archive Humanizer v{0}" -f $script:ToolVersion)
    & $add ""
    & $add "HOW IT IS ORGANISED"
    if ($SeparateNotFinalized) {
        & $add "  Documents\            records whose last lifecycle event in the export is Finalized"
        & $add "  Not finalized\        every other record (awaiting review, flagged by QC, no event)"
        & $add ""
        if ($nSepNot -gt 0) {
            & $add "  Both folders use the same structure: Zone\Section\Artifact\Country or Site."
            & $add ("  {0} file(s) are in 'Documents' and {1} in 'Not finalized'. THE COMPLETE TRIAL MASTER FILE IS BOTH FOLDERS." -f $nDocs, $nSepNot)
            & $add "  Hand over 'Documents' alone only if you mean to leave those records out."
        }
        else { & $add "  No file is in 'Not finalized' in this run: every file that was written is in 'Documents'." }
    }
    else {
        & $add "  Documents\Zone\Section\Artifact\Country or Site\file"
        & $add "  Every version and every record of an artifact is in its artifact folder."
    }
    & $add "  The status of each file is in its name, in brackets, and in inventory.csv."
    & $add ""
    & $add "WHAT THE STATUSES MEAN (the export's own words)"
    & $add ("  Finalized       {0,4}  the last lifecycle event in the export is Finalized" -f $nFinal)
    & $add ("  Not finalized   {0,4}  the last lifecycle event is Awaiting review or Flagged by QC Comment," -f $nNot)
    & $add "                        or the export has none"
    & $add ("  Superseded      {0,4}  marked Superseded/Obsolete by the export, or replaced by a newer Finalized" -f $nSup)
    & $add "                        version (inventory.csv: VersionState, SupersededBy)"
    & $add "  A Finalized version that the export marks Superseded only because a newer version is not"
    & $add "  Finalized is counted Finalized; inventory.csv (VersionResolution) says so."
    if ($SeparateNotFinalized) {
        & $add "  These counts are by status; the folders are split by each record's own last lifecycle"
        & $add "  event, so a Superseded file is in whichever folder that event puts it."
    }
    & $add "  'Document locked' means the record is Finalized and locked. 'Unpublished' on an earlier file"
    & $add "  version means nothing in its own history shows it was ever published or finalized."
    & $add "  Viedoc's own version state is kept unchanged in inventory.csv (VersionState)."
    & $add ""
    & $add "HOW TO USE IT"
    & $add "  Open index.html in any web browser - no internet needed - to search every"
    & $add "  document and read its details, history and signatures."
    & $add "  inventory.csv lists the same documents as a spreadsheet."
    & $add ""
    [System.IO.File]::WriteAllLines($Path, $lines, (New-Object System.Text.UTF8Encoding($false)))
}

function Write-EtmfNotFinalizedReadme {
    param([string]$Path, $Documents)
    $docs = @($Documents | Where-Object { $_.Placement -ceq 'Not finalized' })
    $lines = @(
        'NOT FINALIZED'
        '============='
        ''
        ("This folder holds {0} file(s) whose record's last lifecycle event in the Viedoc export is" -f $docs.Count)
        'not Finalized: Awaiting review, Flagged by QC Comment, or no lifecycle event.'
        ''
        'They are part of the trial master file. They sit here, with the same Zone\Section\Artifact'
        "structure as 'Documents', so that 'Documents' holds only Finalized records."
        ''
        "The status of every file is in inventory.csv (columns Lifecycle and Placement) and in index.html."
    )
    [System.IO.File]::WriteAllLines($Path, $lines, (New-Object System.Text.UTF8Encoding($false)))
}

function Write-EtmfScopeNote {
    param([string]$Path, $Batch, [int]$Included, $Withheld, [string]$SourceInput, [int]$Unwritten = 0)

    $lines = [System.Collections.Generic.List[string]]::new()
    $add = { param([string]$Text = '') $lines.Add($Text) }

    & $add "ABOUT THIS FOLDER"
    & $add "================="
    & $add ""
    & $add "This folder holds the latest FINALIZED version of each document of the trial master"
    & $add ("file for {0}, taken from the eTMF export named below." -f $Batch.SpecificationId)
    & $add ""
    & $add ("  Source export : {0}" -f (Split-Path $SourceInput -Leaf))
    & $add (Get-EtmfMilestoneLine $Batch)
    & $add ("  Prepared      : {0}" -f (Get-Date).ToString('yyyy-MM-dd HH:mm', [System.Globalization.CultureInfo]::InvariantCulture))
    & $add ("  Prepared with : eTMF Archive Humanizer v{0}" -f $script:ToolVersion)
    & $add ""
    & $add "WHAT IS IN HERE"
    & $add ("  {0} file(s) whose last lifecycle event in the export is Finalized and that no" -f $Included)
    & $add "  newer Finalized version replaces. They are filed the familiar TMF way:"
    & $add "  Documents\Zone\Section\Artifact."
    if ($Withheld.NewerPending -gt 0) {
        & $add ("  {0} of them have a newer version in the eTMF that is NOT Finalized; this copy holds" -f $Withheld.NewerPending)
        & $add "  the Finalized one (inventory.csv: NewerVersion, NewerVersionLifecycle)."
    }
    & $add ""
    & $add "WHAT IS NOT IN HERE"
    & $add ("  {0} older version(s) replaced by a newer version, or marked Superseded by the export" -f $Withheld.Superseded)
    & $add ("  {0} file(s) whose last lifecycle event in the export is not Finalized" -f $Withheld.InProgress)
    & $add "  (awaiting review, flagged by QC, or no lifecycle event)"
    if (@($Withheld.Items).Count -gt 0) {
        & $add "  A list of every one of them, with the reason, is kept by whoever prepared this copy"
        & $add "  ('<this folder's name> - withheld.csv', stored beside this folder, not in it)."
    }
    if ($Withheld.NoFile -gt 0) {
        & $add ("  {0} of the included file(s) are listed but their file is not present in the" -f $Withheld.NoFile)
        & $add '  export; they appear in index.html marked "source file not available".'
    }
    if ($Unwritten -gt 0) {
        & $add ("  {0} of the included file(s) could NOT be written to this folder (for example a path too long" -f $Unwritten)
        & $add "  for Windows); they are listed in inventory.csv with the reason in Note and are not in here."
    }
    if ($Withheld.IntegrityProblems -gt 0) {
        & $add ("  {0} of the included file(s) FAILED the integrity check (their MD5 does not match the" -f $Withheld.IntegrityProblems)
        & $add "  manifest); they are marked in inventory.csv (SourceMd5Ok) and in index.html."
    }
    & $add ""
    & $add "THIS IS NOT THE COMPLETE TRIAL MASTER FILE."
    & $add "Superseded versions, records that are not Finalized and their audit trails are"
    & $add "not part of this copy. It exists so the documents the export marks Finalized can"
    & $add "be read and searched without the rest."
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
    Write-Out ("  Run at     : {0}" -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz', [System.Globalization.CultureInfo]::InvariantCulture))
    Write-Out ("  Input      : {0}" -f $inputDesc)
    Write-Out ("  Options    : {0}" -f $optText)
    Write-Out ("  Running on : PowerShell {0} / {1} / machine {2}" -f $PSVersionTable.PSVersion, $os, [Environment]::MachineName)
    Write-Out "===============================================================" 'White'
}

# ----------------------------------------------------------------------------------------
#  Dry-run summary
# ----------------------------------------------------------------------------------------
# Log what the lifecycle decisions relied on (see Get-EtmfLifecycleDetail). The names that counted
# as approval are listed so the run record shows the vocabulary it trusted; a name that looks like
# a lifecycle event but is not one we know is a WARN, because ignoring it could hide an approval.
function Write-EtmfLifecycleAudit {
    if (-not (Test-Path -Path 'variable:script:LifecycleAudit')) { return }
    $a = $script:LifecycleAudit
    if ($null -eq $a) { return }       # a manifest with no records reads no events
    $fmt = { param($t) (($t.GetEnumerator() | Sort-Object Name | ForEach-Object { '{0} x{1}' -f $_.Name, $_.Value }) -join '; ') }
    if ($a.Used.Count -gt 0) {
        Write-Log ("Audit events used for the lifecycle and lock state (the last of each in time order decides; counts are events): {0}" -f (& $fmt $a.Used)) 'INFO'
    }
    if ($a.Ignored.Count -gt 0) {
        Write-Log ("Audit events this tool does not use, ignored: {0}" -f (& $fmt $a.Ignored)) 'INFO'
    }
    foreach ($h in ($a.NearMiss.GetEnumerator() | Sort-Object Name)) {
        Write-Log ("Audit event '{0}' (x{1}) looks like a lifecycle event but is not one this tool knows, so it was IGNORED. Check that no approval is being missed." -f $h.Name, $h.Value) 'WARN'
    }
    foreach ($n in ($a.Notes.GetEnumerator() | Sort-Object Name)) {
        Write-Log ("{0} (x{1})" -f $n.Name, $n.Value) 'WARN'
    }
    if ($a.Unknown -gt 0) {
        Write-Log ("{0} record(s) carry no lifecycle event in their audit trail: the export does not say whether they were approved, so none is claimed and they are filed as not approved." -f $a.Unknown) 'WARN'
    }
}

# What went in versus what came out, so a reader can see that nothing was lost. Every FILE entry in
# the manifests is either a row of the inventory, an identical repeat inside one record that was
# filed once, or a repeat collapsed across milestone exports; and every file in a zip archive is
# either referenced by a FILE entry or named here as not referenced. Fields the manifest carries
# that the inventory has no column for are named too.
# What the version resolution did, in the log (Resolve-EtmfVersions).
function Write-EtmfResolutionLog {
    param($Resolution, [string]$Label)
    if ($Resolution.Demoted -gt 0) {
        Write-Log ("Version resolution ({0}): {1} older version(s) across {2} document(s) moved to '03 Superseded' - a newer version exists in the archive." -f `
            $Label, $Resolution.Demoted, $Resolution.Groups) 'OK'
    }
    if ($Resolution.Unexplained -gt 0) {
        Write-Log ("Version note ({0}): {1} Finalized row(s) are marked Superseded by the export but no newer version is in this export; they stay Superseded (not in -FinalOnly). See VersionResolution in inventory.csv." -f `
            $Label, $Resolution.Unexplained) 'INFO'
    }
    if ($Resolution.Promoted -gt 0) {
        Write-Log ("Version resolution ({0}): {1} Finalized version(s) kept as current although the export marks them Superseded, because the newer version is not Finalized (see VersionResolution in inventory.csv)." -f `
            $Label, $Resolution.Promoted) 'OK'
    }
}

# The OBJECTVERSIONSTATE values the export carries. The tool knows Current, Superseded and Obsolete;
# anything else (or none) is treated as "not superseded", which is a guess worth saying out loud.
function Write-EtmfVersionStateAudit {
    param($Documents)
    $counts = New-Object System.Collections.Hashtable ([System.StringComparer]::Ordinal)
    foreach ($d in @($Documents)) {
        $k = if ($d.VersionState) { [string]$d.VersionState } else { '(missing)' }
        if ($counts.ContainsKey($k)) { $counts[$k]++ } else { $counts[$k] = 1 }
    }
    if ($counts.Count -eq 0) { return }
    $text = (@($counts.GetEnumerator() | Sort-Object Name | ForEach-Object { '{0} x{1}' -f $_.Name, $_.Value }) -join ', ')
    Write-Log ("Version states in the export (OBJECTVERSIONSTATE): {0}" -f $text) 'INFO'
    foreach ($k in @($counts.Keys | Sort-Object)) {
        if ($k -cnotin @('Current', 'Superseded', 'Obsolete')) {
            Write-Log ("OBJECTVERSIONSTATE '{0}' ({1} record(s)) is not one the tool knows (Current, Superseded, Obsolete); it is treated as not superseded." -f $k, $counts[$k]) 'WARN'
        }
    }
}

# Files copied out of the archive keep their extension, and Windows does not mark them as downloaded
# (the "Protected View" / SmartScreen prompt does not follow a file that was copied by a script). A
# trial master file should hold documents; say when it holds something a double-click would RUN.
function Write-EtmfActiveContentWarning {
    param($Documents)
    $risky = @('exe','com','scr','bat','cmd','ps1','psm1','vbs','vbe','js','jse','wsf','wsh','hta','lnk','url','msi','msp','jar','iso','reg','dll',
               'docm','xlsm','pptm','dotm','xlsb','xlam','html','htm','svg','chm')
    $hits = @($Documents | Where-Object { $_.FileExists -and $risky -contains ([System.IO.Path]::GetExtension((ConvertTo-SafeName $_.OriginalName)).TrimStart('.').ToLowerInvariant()) })
    if ($hits.Count -gt 0) {
        $names = (@($hits | Select-Object -First 5 | ForEach-Object { "'" + $_.OriginalName + "'" }) -join ', ')
        Write-Log ("{0} file(s) in the archive are executable or can carry active content (for example macros or scripts): {1}{2}. They were copied unchanged and, unlike a download, are not marked as coming from the internet - do not open them without checking." -f `
            $hits.Count, $names, $(if ($hits.Count -gt 5) { ', ...' } else { '' })) 'WARN'
    }
}

function Write-EtmfReconciliation {
    param([int]$Documents, [int]$Collapsed, $ContentMap)
    $r = $script:Reconcile
    $sum = $r.Files - $r.ListedOnce - $r.MergedInto - $Collapsed
    $text = ("Reconciliation: the manifest(s) list {0} record(s) with {1} file entr{2}; {3} repeat(s) inside a record filed once ({6} of them with extra audit events added), {4} repeat(s) collapsed across milestone exports; {5} document row(s) before any -FinalOnly selection." -f `
        $r.Objects, $r.Files, $(if ($r.Files -eq 1) { 'y' } else { 'ies' }), ($r.ListedOnce + $r.MergedInto), $Collapsed, $Documents, $r.MergedInto)
    if ($r.NoFileObjects -gt 0) {
        Write-Log ("{0} record(s) list no FILE at all, so they have no row (see the warnings above)." -f $r.NoFileObjects) 'WARN'
    }
    if ($sum -eq $Documents) { Write-Log $text 'OK' }
    else { Write-Log ($text + (" These do not add up (expected {0}); a record may be missing from the output." -f $sum)) 'WARN' }
    if ($null -ne $ContentMap) {
        $unref = @($ContentMap.Keys | Where-Object { -not $r.Referenced.Contains($_) })
        if ($unref.Count -gt 0) {
            $names = (@($unref | Select-Object -First 5 | ForEach-Object { "'" + $_ + "'" }) -join ', ')
            Write-Log ("{0} file(s) in the archive are not referenced by any manifest entry, so they are not in the output: {1}{2}" -f `
                $unref.Count, $names, $(if ($unref.Count -gt 5) { ', ...' } else { '' })) 'WARN'
        }
        else { Write-Log 'Every file in the archive is referenced by a manifest entry.' 'OK' }
    }
    foreach ($u in ($r.Unread.GetEnumerator() | Sort-Object Name)) {
        Write-Log ("The manifest carries {0} (x{1}), which this tool does not read; it is not in the inventory." -f $u.Name, $u.Value) 'WARN'
    }
}

function Show-Summary {
    param($Model, $Integrity)

    $b = $Model.Batch
    $docs = $Model.Documents

    Write-Out ""
    Write-Out "============ eTMF ARCHIVE SUMMARY (as read, before version resolution) ============" -ForegroundColor White
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
        $label = if ($_.Name) { $_.Name } else { '(none in the export)' }
        Write-Out ("     {0,-22} {1,3}" -f $label, $_.Count)
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
        $noSum = if ($Integrity.Unverified -gt 0) { ", {0} with no checksum stated" -f $Integrity.Unverified } else { '' }
        Write-Out ("  Integrity (MD5)        : {0}/{1} OK, {2} mismatch, {3} missing{4}" -f `
            $Integrity.Ok, $tot, $Integrity.Bad, $Integrity.Missing, $noSum) -ForegroundColor $col
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
    param([string]$OutputRoot, $Written, $Gaps, [int]$Demoted = 0, $Withheld = $null, [int]$NotFinalizedSeparated = -1)
    $e = $Written
    $problems = [int]$e.Missing + [int]$e.CopyBad + [int]$e.WriteFailed + [int]$e.IntegrityBad
    Write-Out ""
    if ($problems -gt 0) { Write-Out "============ DONE - BUT $problems FILE PROBLEM(S): READ BELOW ============" -ForegroundColor Yellow }
    else { Write-Out "==================== DONE ====================" -ForegroundColor White }
    Write-Out ("  Output folder : {0}" -f $OutputRoot) -ForegroundColor Green
    if ($null -ne $Withheld) {
        Write-Out  "  Contents      : LATEST FINALIZED VERSIONS ONLY - a copy to hand over," -ForegroundColor Cyan
        Write-Out  "                  not the complete trial master file." -ForegroundColor Cyan
        Write-Out ("                  Left out: {0} superseded, {1} not Finalized." -f `
            $Withheld.Superseded, $Withheld.InProgress) -ForegroundColor Cyan
        Write-Out  "                  'About this folder.txt' states this for the recipient; withheld.csv (beside the folder) lists what was left out." -ForegroundColor Cyan
    }
    if ($NotFinalizedSeparated -ge 0) {
        Write-Out ("  Layout        : {0} file(s) are in 'Not finalized\'. The complete TMF is 'Documents' AND" -f $NotFinalizedSeparated) -ForegroundColor Cyan
        Write-Out  "                  'Not finalized' together; 'Documents' alone leaves them out." -ForegroundColor Cyan
    }
    $flag = if ($problems -gt 0) { '  [!]' } else { '' }
    $col  = if ($flag) { 'Yellow' } else { 'Gray' }
    # The counts must add up to the total, or the reader is left hunting for the difference.
    Write-Out ("    {0,-30} {1,3} files  ({2} finalized / {3} not finalized / {4} superseded)  written {5}{6}" -f `
        $e.Event, $e.Docs, $e.Final, $e.InProgress, $e.Superseded, $e.Copied, $flag) -ForegroundColor $col
    if ($e.Missing -gt 0) { Write-Out ("           {0} document(s) missing source files" -f $e.Missing) -ForegroundColor Yellow }
    if ($e.CopyBad -gt 0) { Write-Out ("           {0} checksum mismatch(es) after writing" -f $e.CopyBad) -ForegroundColor Yellow }
    if ($e.IntegrityBad -gt 0) { Write-Out ("           {0} source file(s) do not match the checksum in the manifest" -f $e.IntegrityBad) -ForegroundColor Yellow }
    if ($e.WriteFailed -gt 0) { Write-Out ("           {0} file(s) could not be written" -f $e.WriteFailed) -ForegroundColor Yellow }
    # The headline answer to "which one is the current approved version?".
    if ($Demoted -gt 0) {
        Write-Out ""
        Write-Out ("  {0} older version(s) were filed under '03 Superseded' because a newer" -f $Demoted) -ForegroundColor Cyan
        Write-Out  "  version of the same document exists in the archive. inventory.csv's" -ForegroundColor Cyan
        Write-Out  "  SupersededBy column names the version that replaced each one." -ForegroundColor Cyan
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
    Write-Out ""
    Write-Out ("  Open viewer   : {0}" -f $e.Index) -ForegroundColor Green
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
  .chip.on[data-s="01 Finalized"]{background:var(--final)}
  .chip.on[data-s="02 Not finalized"]{background:var(--prog)}
  .chip.on[data-s="03 Superseded"]{background:var(--sup)}
  .resultCount{margin-left:auto;color:var(--muted);font-size:12px;white-space:nowrap}
  [hidden]{display:none !important}
  .filterbar{display:flex;align-items:center;gap:8px;flex-wrap:wrap;padding:8px 16px;background:#fdf8e8;border-bottom:1px solid #ecdfae;font-size:12.5px;color:#6b5a1e}
  .fb-label{font-weight:600}
  .fchip{background:#fff;border:1px solid #e5d9ae;border-radius:20px;padding:2px 10px;white-space:nowrap}
  #filterList{display:flex;flex-wrap:wrap;gap:6px;min-width:0}
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
          <th class="col-ms" data-key="milestone" title="The milestone (manifest of the export) this record is listed under">Milestone</th>
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
// Only an ISO-8601 time is reformatted (a time with no offset is read as UTC, as the tool reads it);
// anything else is shown exactly as the export has it. new Date(s) alone would read 'Jan 5' against
// the current year and '03/04/2024' month-first, and a viewer must not invent a date.
const fmtTs = s => {
  if(!s) return '';
  const m = /^(\d{4})-(\d\d)-(\d\d)T(\d\d):(\d\d)(?::(\d\d)(?:\.\d+)?)?(Z|[+-]\d\d:\d\d)?$/.exec(String(s).trim());
  if(!m) return s;
  const d=new Date(String(s).trim()+(m[7]?'':'Z'));
  return isNaN(d)? s : d.toISOString().slice(0,19).replace('T',' ')+' UTC';
};
// Same rule as Compare-EtmfVersion for dotted-numeric versions (non-numeric labels may order differently): segment-wise, so v10 > v9 and
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
const bClass = b => b==='01 Finalized'?'b-final':b==='02 Not finalized'?'b-prog':'b-sup';

const state = { search:'', searchRaw:'', status:'all', country:'all', milestone:'all', node:null, sortKey:'artNo', sortDir:1 };

// One lower-cased haystack per document, built once. Searching a TMF means searching for whatever
// the reviewer happens to remember - an artifact number, a zone, a site, a status word, a document
// ID, the milestone - so every field a reader can see is in here, and the per-keystroke filter is
// then a single string test rather than a dozen concatenations per row.
// PowerShell's ConvertTo-Json emits a one-element array as a bare value on some hosts, so never
// assume this field is an array.
const msList = d => !d.milestones ? [] : (Array.isArray(d.milestones) ? d.milestones : [d.milestones]);
DATA.forEach(d => { d._hay = [d.artNo, d.title, d.originalName, d.newName, d.artifact, d.zone,
  d.section, d.scope, d.country, d.level, d.person, d.version, d.status, d.lifecycle, d.lock,
  d.versionState, d.documentId, msList(d).join(' '), d.date, d.supersededBy,
  d.placement, d.resolution, d.lifecycleNote, d.latestApproved, d.newerVersion, d.newerLifecycle, d.language, d.organization, d.repeats]
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
  (META.finalOnly ? ` · <b>latest Finalized versions only</b>` : '');
const intTxt = META.integrityChecked ? `${META.integrityOk}/${META.integrityTotal}${META.integrityUnverified?` (+${META.integrityUnverified} unverified)`:''}` : 'n/a';
const countDefs = [['Total',META.total]];
if (!META.finalOnly) countDefs.push(['Finalized',META.final],['Not&nbsp;finalized',META.inProgress],['Superseded',META.superseded]);
countDefs.push(['E-signed',META.signed],['Integrity',intTxt]);
if (META.missing)  countDefs.push(['Missing',META.missing]);
$('#counts').innerHTML = countDefs
  .map(([l,n])=>`<div class="count"><div class="n">${n}</div><div class="l">${l}</div></div>`).join('');
$('#footer').innerHTML = `Generated ${esc(META.generated)} from <b>${esc(META.manifest)}</b> · ${META.total} documents · eTMF Archive Humanizer <b>v${esc(META.toolVersion)}</b>`;

// ---- status chips ----
// Only offer a filter for a bucket this copy actually contains: in a -FinalOnly copy an "In
// Progress" chip that always lists nothing invites the reader to wonder what is being hidden.
const bucketsPresent = new Set(DATA.map(d=>d.bucket));
const chips = [['all','All']].concat(
  [['01 Finalized','Finalized'],['02 Not finalized','Not finalized'],['03 Superseded','Superseded']]
    .filter(([b])=>bucketsPresent.has(b)));
// One bucket only (a -FinalOnly copy): "All" and "Final" would be the same filter.
if (chips.length < 3) { $('#statusChips').hidden = true; }
$('#statusChips').innerHTML = chips.map(([s,l])=>`<span class="chip${s==='all'?' on':''}" data-s="${s}">${l}</span>`).join('');
$('#statusChips').addEventListener('click',e=>{ const c=e.target.closest('.chip'); if(!c) return;
  state.status=c.dataset.s;
  // Scoped to #statusChips: a document-wide .chip selector would also de-highlight any other chip.
  document.querySelectorAll('#statusChips .chip').forEach(x=>x.classList.toggle('on',x===c)); render(); });

// ---- milestone filter ----
// A milestone is a property of the artifact: each record sits in exactly one manifest of the export,
// the one for its artifact's milestone. This filters on the milestone(s) the record is listed under;
// the options come from the export's list of manifests.
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

// ---- search ----
$('#search').addEventListener('input',e=>{ state.searchRaw=e.target.value; state.search=e.target.value.toLowerCase(); render(); });

// ---- tree ----
function buildTree(){
  const t={};
  DATA.forEach(d=>{
    const z=t[d.zone]??=({c:0,f:0,s:{}}); z.c++; if(d.bucket!=='01 Finalized')z.f++;
    const se=z.s[d.section]??=({c:0,f:0,a:{}}); se.c++; if(d.bucket!=='01 Finalized')se.f++;
    const a=se.a[d.artifact]??=({c:0,f:0}); a.c++; if(d.bucket!=='01 Finalized')a.f++;
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
// The user's own filters.
function matches(d){
  if(state.status!=='all' && d.bucket!==state.status) return false;
  if(state.country!=='all' && d.country!==state.country) return false;
  if(state.milestone!=='all' && msList(d).indexOf(state.milestone)<0) return false;
  if(state.node){ if(state.node.type==='zone'&&d.zone!==state.node.v) return false;
    if(state.node.type==='section'&&d.section!==state.node.v) return false;
    if(state.node.type==='artifact'&&d.artifact!==state.node.v) return false; }
  if(state.search && !d._hay.includes(state.search)) return false;
  return true;
}
// ---- active filters: say what is being hidden, and offer one click to undo it ----
const chipLabel = s => (chips.find(c=>c[0]===s)||[,s])[1];
function activeFilters(){
  const a=[];
  if(state.searchRaw)          a.push(['Search', '\u201c'+state.searchRaw+'\u201d']);
  if(state.status!=='all')     a.push(['Status', chipLabel(state.status)]);
  if(state.country!=='all')    a.push(['Country', state.country]);
  if(state.milestone!=='all')  a.push(['In milestone', state.milestone]);
  if(state.node)               a.push([state.node.type.charAt(0).toUpperCase()+state.node.type.slice(1), state.node.v]);
  return a;
}
function renderFilterBar(){
  const a=activeFilters();
  $('#filterBar').hidden = (a.length===0);
  $('#filterList').innerHTML = a.map(([k,v])=>`<span class="fchip">${esc(k)}: ${esc(v)}</span>`).join('');
}
$('#clearFilters').addEventListener('click',()=>{
  state.search=''; state.searchRaw=''; state.status='all'; state.country='all'; state.milestone='all'; state.node=null;
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
    const integ = !d.present ? '<span class="no" title="Source file missing">⚠</span>'
                : !d.integrityStated ? '<span title="The manifest states no checksum for this file">–</span>'
                : d.copyOk ? '<span class="ok" title="MD5 verified">✓</span>'
                : (META.integrityChecked?'<span class="no" title="Integrity mismatch">✗</span>':'');
    return `<tr data-i="${idx}">
      <td class="art">${esc(d.artNo)}</td>
      <td>${esc(d.title)}${sig}${restr}</td>
      <td class="doc" title="${esc(d.originalName)}">${esc(d.originalName)}</td>
      <td>${esc(d.scope||(d.level==='Trial'?'—':d.country))}</td>
      <td class="ms col-ms" title="Milestone: ${esc(msList(d).join('; '))}">${esc(d.milestone)}</td>
      <td>${esc(d.version)}</td>
      <td>${d.status?`<span class="badge ${bClass(d.bucket)}">${esc(d.status)}</span>`:''}</td>
      <td>${esc(d.date)}</td>
      <td>${integ}</td></tr>`;
  }).join('');
  $('#empty').style.display = rows.length?'none':'block';
  $('#resultCount').textContent = rows.length+' of '+DATA.length+' documents';
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
  const versions=DATA.map((x,ix)=>({x,ix})).filter(o=>o.x.documentId&&o.x.documentId===d.documentId&&o.x.artNo===d.artNo&&o.x.country===d.country&&(o.x.siteId||'')===(d.siteId||''))
    .sort((a,b)=>cmpVersion(a.x.version,b.x.version));
  const vh = versions.length>1 ? `<div class="sec">Version history</div><div class="vh">`+
    versions.map(o=>{
      const label=`${o.x.status?`<span class="badge ${bClass(o.x.bucket)}">${esc(o.x.status)}</span> `:''}v${esc(o.x.version)} <span style="color:var(--muted)">${esc(o.x.originalName)}</span>`;
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
      ${d.status?`<span class="badge ${bClass(d.bucket)}">${esc(d.status)}</span>`:''}
      <span style="color:var(--muted);font-size:12px"> ${esc(d.artNo)} · v${esc(d.version)}</span>
    </div>
    <div class="d-body">
      ${d.present && d.relPath ? `<a class="openbtn" href="${href(d.relPath)}"${linkTgt}>Open document ↗</a>`
        : `<div class="openbtn" style="background:var(--sup);cursor:default">⚠ Source file not available</div>
           ${d.gapReason?`<div class="gapwhy">Why: ${esc(d.gapReason)}.<br>The record is kept so the gap is visible; the file itself is not in the export.</div>`:''}`}
      <div class="sec">Document</div>
      <div class="kv">
        ${kv('Artifact',d.artifact)}${kv('Sub-artifact / title',d.title)}${kv('TMF level',d.level)}
        ${kv('Country',d.country)}${kv('Site / org (folder name)',d.scope)}${kv('Organization (export)',d.organization)}${kv('Person',d.person)}
        ${META.separated?kv('Placement',d.placement):''}${kv('Version',d.version)}${kv('Version state',d.versionState)}${kv('Lifecycle',d.lifecycle)}${kv('Lock',d.lock)}
        ${kv('Superseded by',d.supersededBy)}${kv('Latest approved version',d.latestApproved)}${kv('Newer version',d.newerVersion?('v'+d.newerVersion+(d.newerLifecycle?' ('+d.newerLifecycle+')':' (no lifecycle event)')):'')}${kv('Version resolution',d.resolution)}${kv('Lifecycle note',d.lifecycleNote)}
        ${kv('Restricted',d.restrictedRaw)}${kv('Language',d.language)}${kv('Translation',d.translation)}
        ${kv(d.dateDesc||'Artifact date',d.date)}${kv('Expires',d.expiry)}${kv('Document ID',d.documentId)}
        ${kv('Milestone',d.milestone)}${msList(d).length>1?kv('Listed in milestones',msList(d).join('; ')):''}
        ${kv('Size',fmtSize(d.sizeBytes))}
      </div>
      <div class="sec">Files</div>
      <div class="kv">
        ${kv('Original name',d.originalName)}${kv('Stored as',d.newName)}${kv('Repeated entry',d.repeats)}
        ${kv('Integrity (MD5)',d.integrity)}
        <div class="k">Verified</div><div class="v">${!d.present?'no file to verify':(d.present&&!d.relPath)?'file not written (see inventory.csv, Note)':!d.integrityStated?'no checksum stated in the manifest':d.copyOk?'<span class="ok">✓ copy matches manifest</span>':(META.integrityChecked?'<span class="no">✗ mismatch</span>':'not checked')}</div>
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

# ========================================================================================
#  MAIN
# ========================================================================================
$script:OutputRoot = $null
$script:LogFallbackDir = $null
$script:TempDir    = $null      # set as soon as the temp dir exists, so cleanup survives a throw
$script:Demoted    = 0          # older versions filed under 03 Superseded by version resolution
$source = $null

# No archive supplied -> show friendly instructions and stop (instead of prompting).
if ([string]::IsNullOrWhiteSpace($ArchivePath)) {
    Show-Usage
    return
}

try {
    Show-RunHeader -ArchivePath $ArchivePath -BoundParameters $PSBoundParameters

    # Resolve where the output will live (sibling of the input by default).
    $inputItem = Get-Item -LiteralPath $ArchivePath
    if ($inputItem.PSIsContainer) {
        throw "ArchivePath must be the .zip file of the export, not a folder: $($inputItem.FullName)"
    }
    # Where the run log goes if the run dies before there is an output folder - a crash log is the one
    # we most need. Set only once the input is known to exist: a run that fails because the path was
    # wrong has nothing in its log the user did not just read on screen, and writing it would drop a
    # file into whatever directory they happened to be standing in.
    $script:LogFallbackDir = $inputItem.DirectoryName
    $baseDir               = $inputItem.DirectoryName

    $source = Resolve-EtmfSource -ArchivePath $ArchivePath
    # The extractor tells us where the manifests are, and what archive prefix each one's
    # CONTENTURLs are relative to.
    $manifests = @($source.Manifests | ForEach-Object {
        [pscustomobject]@{ File = (Get-Item -LiteralPath $_.Path); Prefix = $_.Prefix }
    })
    if ($manifests.Count -eq 0) { throw "No '*exchange.xml' manifest found in: $ArchivePath" }
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
        [pscustomobject]@{ Model = $model; Manifest = $mf }
    }
    $events = @($events)
    Write-EtmfLifecycleAudit
    Write-EtmfVersionStateAudit -Documents @($events | ForEach-Object { $_.Model.Documents })

    # One Viedoc export writes all its manifests under one TransferID. Different ones mean the input
    # is not one export, which this tool does not handle; say so rather than build a tree quietly.
    $transferIds = @($events | ForEach-Object { $_.Model.Batch.TransferId } | Where-Object { $_ } | Select-Object -Unique)
    if ($transferIds.Count -gt 1) {
        Write-Log ("The manifests carry {0} different TransferIDs ({1}), so they are not from one export. This tool converts one export at a time; check the input." -f `
            $transferIds.Count, ($transferIds -join ', ')) 'WARN'
    }
    if ($FinalOnly -and $SeparateNotFinalized) {
        Write-Log '-SeparateNotFinalized is ignored with -FinalOnly: that copy holds only Finalized documents.' 'WARN'
    }
    $sepNotFinalized = [bool]($SeparateNotFinalized -and -not $FinalOnly)

    if ($DryRun) {
        # Preview the decisions, not just the parse: a reviewer uses -DryRun to see what WOULD be
        # merged and superseded before committing to an output folder.
        $preview = Merge-EtmfEvents -Events $events
        $pv = Resolve-EtmfVersions -Model $preview
        Write-Log ("DryRun: {0} document(s) after merge ({1} duplicate(s) collapsed); {2} older version(s) would be superseded and {3} kept as current although the export marks them Superseded." -f `
            $preview.Documents.Count, $preview.Collapsed, $pv.Demoted, $pv.Promoted) 'INFO'
        Write-EtmfReconciliation -Documents $preview.Documents.Count -Collapsed $preview.Collapsed -ContentMap $source.ContentMap
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
    # An absolute path from here on. PowerShell cmdlets resolve a relative path against the SESSION
    # location and .NET file APIs against the PROCESS directory; after a `cd` they differ, and half of
    # the output would land in one place and half in the other.
    $targetRoot = if ($OutputPath) { $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath) }
                  else { Join-Path $baseDir (ConvertTo-SafeName ($nameTemplate -f $studyName)) }
    if ($targetRoot.Length -gt 1 -and $targetRoot -notmatch '^[A-Za-z]:[\\/]$') { $targetRoot = $targetRoot.TrimEnd('\', '/') }

    # The output folder is rebuilt (and, with -Force, emptied first): it must never be, or contain,
    # the input - that would delete the archive the user handed us.
    # Lexical first (case-insensitively everywhere: a false refusal only affects two folders that
    # differ by case alone, a missed one deletes the input), and never a filesystem root.
    $comparison = [System.StringComparison]::OrdinalIgnoreCase
    $inputFull  = $inputItem.FullName.TrimEnd('\', '/')
    $sep        = [string][System.IO.Path]::DirectorySeparatorChar
    $pathRoot   = [System.IO.Path]::GetPathRoot($targetRoot)
    if ($pathRoot -and $pathRoot.TrimEnd('\', '/') -eq $targetRoot.TrimEnd('\', '/') -or $targetRoot -eq '/') {
        throw "The output folder '$targetRoot' is a filesystem root. Choose a folder of its own."
    }
    if ($inputFull.Equals($targetRoot, $comparison) -or $inputFull.StartsWith($targetRoot + $sep, $comparison) -or $inputFull.StartsWith($targetRoot + '/', $comparison)) {
        throw "The output folder '$targetRoot' is, or contains, the input '$($inputItem.FullName)'. Choose an -OutputPath that is separate from the input."
    }

    $replaceExisting = $false
    if (Test-Path -LiteralPath $targetRoot) {
        if ($Force) {
            # Safety: only ever auto-delete a directory this tool made. Every run writes a marker file
            # as the first thing it puts in the output folder; a folder from v1.x has no marker, so
            # it is also accepted when it holds BOTH index.html and inventory.csv. A folder that
            # merely has a 'Documents' subfolder (every Windows profile does) or a log file is NOT
            # accepted: a refused or dry run leaves a log beside the input, and that must never be
            # what arms a later -Force.
            $hasMarker  = Test-Path -LiteralPath (Join-Path $targetRoot $script:OutputMarkerName) -PathType Leaf
            $hasLegacy  = (Test-Path -LiteralPath (Join-Path $targetRoot 'index.html') -PathType Leaf) -and
                          (Test-Path -LiteralPath (Join-Path $targetRoot 'inventory.csv') -PathType Leaf)
            if (-not ($hasMarker -or $hasLegacy)) {
                throw "Refusing to -Force delete '$targetRoot': it was not made by this tool " +
                      "(no $($script:OutputMarkerName) marker, and no index.html + inventory.csv pair). Remove it manually or choose another -OutputPath."
            }
            # The comparison above is by text, which a link (symlink, junction, subst/mapped drive, 8.3
            # name) or a case-insensitive filesystem defeats. So ask the filesystem: put a uniquely
            # named file in the folder about to be emptied and look for it from the input's own
            # location upwards. If it shows up there, the input lives inside the folder.
            $canaryName = '.etmf-canary-' + [guid]::NewGuid().ToString('N')
            $canaryPath = Join-Path $targetRoot $canaryName
            Set-Content -LiteralPath $canaryPath -Value 'x'
            try {
                $probe = $inputItem.DirectoryName
                while ($probe) {
                    if (Test-Path -LiteralPath (Join-Path $probe $canaryName)) {
                        throw "The output folder '$targetRoot' contains the input '$($inputItem.FullName)' (reached through a link or a differently spelled path). Choose an -OutputPath that is separate from the input."
                    }
                    $parent = [System.IO.Path]::GetDirectoryName($probe)
                    if ($parent -eq $probe) { break }
                    $probe = $parent
                }
            }
            finally { Remove-Item -LiteralPath $canaryPath -Force -ErrorAction SilentlyContinue }
            $replaceExisting = $true        # removed further down, once it is known the new output fits
        } else {
            throw "Output folder already exists: $targetRoot  (use -Force to overwrite)"
        }
    }

    # ---- Phases 5 + 6: build tree + inventory + viewer ----
    # ONE merged tree across every milestone export. A milestone export is a snapshot of the whole
    # TMF, so a folder per milestone would hand the reader N partial TMFs, N indexes and no way to
    # search across them - which is exactly the state the archive was already in.
    $model = Merge-EtmfEvents -Events $events
    if ($events.Count -gt 1) {
        Write-Log ("Merged {0} milestone export(s) into one TMF: {1} document(s) after collapsing {2} duplicate(s)." -f `
            $events.Count, $model.Documents.Count, $model.Collapsed) 'OK'
    }
    elseif ($model.Collapsed -gt 0) {
        # One manifest can still list the same file twice under two OBJECTs.
        Write-Log ("Collapsed {0} duplicate record(s) pointing at the same file." -f $model.Collapsed) 'INFO'
    }
    # Now that every snapshot is in one model, decide which VERSION of each document is the
    # current one - the per-snapshot OBJECTVERSIONSTATE cannot know about a later replacement.
    $resolution = Resolve-EtmfVersions -Model $model
    Write-EtmfResolutionLog -Resolution $resolution -Label 'merged'
    $script:Demoted = $resolution.Demoted

    # Integrity was verified per manifest; restate it over the merged set so the viewer's
    # header counts what this view actually holds rather than one snapshot's worth.
    $integrity = $null
    if (-not $SkipIntegrityCheck) { $integrity = Get-EtmfIntegritySummary $model.Documents }
    $label = if ($model.Batch.IsMerged) { "All {0} milestones (merged)" -f @($model.Batch.Milestones).Count }
             else                       { $model.Batch.Milestone }

    Write-EtmfReconciliation -Documents $model.Documents.Count -Collapsed $model.Collapsed -ContentMap $source.ContentMap

    # -FinalOnly: reduce the model to the current approved documents, and remember what that
    # left out so the copy can state its own scope.
    $withheld = $null
    if ($FinalOnly) {
        $withheld = Select-EtmfFinalDocuments -Model $model
        # A list left by an earlier run beside this folder would describe other documents.
        $handOverParent = Split-Path $targetRoot -Parent
        $withheldPath = if ($handOverParent) { Join-Path $handOverParent ('{0} - withheld.csv' -f (Split-Path $targetRoot -Leaf)) }
                        else { Join-Path $targetRoot 'withheld.csv' }
        # (the earlier list is removed and the new one written further down, after the space check: a run that
        # cannot fit must leave the earlier hand-over folder AND its list as they were)
        # the viewer's integrity tile must count the documents this copy holds, not the ones it left out
        if ($null -ne $integrity) { $integrity = Get-EtmfIntegritySummary $model.Documents }
        Write-Log ("-FinalOnly: keeping {0} document(s) whose latest version is Finalized; leaving out {1} superseded and {2} not Finalized (itemised in withheld.csv)." -f `
            $model.Documents.Count, $withheld.Superseded, $withheld.InProgress) 'OK'
        if ($withheld.NewerPending -gt 0) {
            Write-Log ("{0} of the kept document(s) have a newer version in the eTMF that is not Finalized; the copy holds the latest Finalized version (see NewerVersion in inventory.csv)." -f $withheld.NewerPending) 'INFO'
        }
    }

    # The extracted files are moved into the tree rather than copied (see Invoke-Humanize), so on the volume
    # that holds the temp folder a run adds only the copies that shared files need; on another volume the
    # whole output is written there. Say so now if it will not fit, not after the disk fills up halfway - and
    # before an earlier output is removed for -Force, so a run that cannot succeed leaves it alone.
    $tempVolume   = [System.IO.Path]::GetPathRoot([System.IO.Path]::GetFullPath($script:TempDir))
    $outputVolume = [System.IO.Path]::GetPathRoot([System.IO.Path]::GetFullPath($targetRoot))
    # Same volume = a move is a rename. Windows: drive letters. Elsewhere every path has the root '/', which proves
    # nothing, and there is no portable device id to ask for - so the volumes are assumed to be the same. That is the
    # LOWER bound on what the output needs: this check may miss a shortfall (as before it existed) but never refuses a
    # run that would have fitted, which is the one thing it must not do.
    if ([System.IO.Path]::DirectorySeparatorChar -eq '\') {
        $sameVolume = [string]::Equals($tempVolume, $outputVolume, [System.StringComparison]::OrdinalIgnoreCase)
    }
    else { $sameVolume = $true }
    $outputBytes  = Get-EtmfOutputBytes -Documents $model.Documents -MovesAreFree $sameVolume
    $outputSpace  = Get-EtmfSpaceCheck -Path $targetRoot -Bytes $outputBytes
    $reclaimed    = 0L
    if ($null -ne $outputSpace -and $outputSpace.Short -and $replaceExisting) {
        # The folder about to be replaced gives its space back, so count it before refusing.
        $reclaimed = [long](Get-ChildItem -LiteralPath $targetRoot -Recurse -File -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
        $outputSpace.Short = (($outputSpace.Free + $reclaimed) -lt $outputSpace.Need)
    }
    if ($null -ne $outputSpace) {
        $credit = if ($reclaimed -gt 0) { " (plus {0} that the folder being replaced gives back)" -f (Format-EtmfSize $reclaimed) } else { '' }
        Write-Log ("The output adds about {0} on {1}; {2} is free{3}." -f (Format-EtmfSize $outputBytes), $outputVolume, (Format-EtmfSize $outputSpace.Free), $credit) 'INFO'
        if ($outputSpace.Short) {
            throw ("Not enough free space on {0} for the output folder: about {1} will be written there, and {2} is free{3}. Choose an -OutputPath on a drive with more room, or free up space." -f `
                $outputVolume, (Format-EtmfSize $outputBytes), (Format-EtmfSize $outputSpace.Free), $credit)
        }
    }

    if ($replaceExisting) {
        Write-Log "Output exists; -Force given, removing: $targetRoot" 'WARN'
        Remove-Item -LiteralPath $targetRoot -Recurse -Force
    }
    New-Item -ItemType Directory -Path $targetRoot -Force | Out-Null
    $script:OutputRoot = $targetRoot       # (a different name on purpose: PowerShell variable names are case-insensitive, so $outputRoot WOULD BE $script:OutputRoot) only now: a folder this run did not make must never receive its log
    Set-Content -LiteralPath (Join-Path $targetRoot $script:OutputMarkerName) -Encoding UTF8 -Value @(
        "Created by eTMF Archive Humanizer v$($script:ToolVersion) on $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)).",
        'This marker lets -Force recognise the folder as output of this tool. Do not delete it unless you want -Force to refuse this folder.')
    Write-Log "Output folder: $targetRoot" 'INFO'

    if ($FinalOnly) {
        # A list left by an earlier run beside this folder would describe other documents.
        if (Test-Path -LiteralPath $withheldPath -PathType Leaf) { Remove-Item -LiteralPath $withheldPath -Force }
        if (@($withheld.Items).Count -gt 0) {
            # Beside the hand-over folder, not inside it: the recipient gets the folder, and this list
            # names the (unapproved) documents that were left out of it.
            Write-EtmfCsv -Path $withheldPath -Rows @($withheld.Items)
            Write-Log ("Left-out documents itemised in: {0}" -f $withheldPath) 'OK'
        }
    }

    Write-Log ("Humanizing '{0}' -> {1}" -f $label, $targetRoot) 'STEP'
    $result = Invoke-Humanize -Model $model -OutputRoot $targetRoot `
                -VerifyCopy:(-not $SkipIntegrityCheck) `
                -SeparateNotFinalized:$sepNotFinalized `
                -MoveSources
    Write-EtmfActiveContentWarning -Documents $model.Documents
    Write-Log ("Wrote {0} file(s) into the output; {1} missing; {2} checksum mismatch(es) after writing." -f `
        $result.Copied, $result.Missing, $result.CopyBad) ($(if ($result.CopyBad -eq 0 -and $result.Missing -eq 0 -and $result.WriteFailed -eq 0) { 'OK' } else { 'WARN' }))
    if ($result.WriteFailed -gt 0) { Write-Log ("{0} file(s) could not be written (see the warnings above); their rows are in inventory.csv with the reason in Note." -f $result.WriteFailed) 'WARN' }

    $invPath = Join-Path $targetRoot 'inventory.csv'
    Write-EtmfCsv -Path $invPath -Rows $result.Inventory
    Write-Log "Inventory written: $invPath" 'OK'

    $indexPath = New-HtmlViewer -Model $model -OutputRoot $targetRoot -Integrity $integrity -FinalOnlyCopy:$FinalOnly -SeparatedCopy:$sepNotFinalized
    Write-Log "Viewer written: $indexPath" 'OK'

    $notePath = Join-Path $targetRoot 'About this folder.txt'
    if ($FinalOnly) {
        Write-EtmfScopeNote -Path $notePath -Batch $model.Batch -Included $model.Documents.Count `
            -Withheld $withheld -SourceInput $ArchivePath -Unwritten $result.WriteFailed
        Write-Log "Scope note written: $notePath" 'OK'
    }
    else {
        Write-EtmfLayoutNote -Path $notePath -Batch $model.Batch -Documents $model.Documents -SourceInput $ArchivePath `
            -SeparateNotFinalized:$sepNotFinalized
        Write-Log "Layout note written: $notePath" 'OK'
        $nfDir = Join-Path $targetRoot 'Not finalized'
        if ($sepNotFinalized -and (Test-Path -LiteralPath $nfDir -PathType Container)) {
            Write-EtmfNotFinalizedReadme -Path (Join-Path $nfDir 'README.txt') -Documents $model.Documents
        }
    }

    $docs = $model.Documents
    $written = [pscustomobject]@{
        Event = $label; Index = $indexPath
        Docs = $docs.Count; Copied = $result.Copied; Missing = $result.Missing; CopyBad = $result.CopyBad; WriteFailed = $result.WriteFailed; IntegrityBad = @($docs | Where-Object { $_.FileExists -and $_.IntegrityOk -eq $false }).Count
        Final      = @($docs | Where-Object { $_.Bucket -eq '01 Finalized' }).Count
        InProgress = @($docs | Where-Object { $_.Bucket -eq '02 Not finalized' }).Count
        Superseded = @($docs | Where-Object { $_.Bucket -eq '03 Superseded' }).Count
    }

    # Gap causes, most frequent first (empty when everything bound).
    $gaps = @($docs |
              Where-Object { -not $_.FileExists } |
              Group-Object { if ($_.GapReason) { $_.GapReason } else { 'no reason recorded' } } |
              Sort-Object Count -Descending |
              ForEach-Object { [pscustomobject]@{ Count = $_.Count; Reason = $_.Name } })

    Show-Recap -OutputRoot $targetRoot -Written $written -Gaps $gaps `
               -Demoted $script:Demoted -Withheld $withheld `
               -NotFinalizedSeparated $(if ($sepNotFinalized) { @($docs | Where-Object { $_.Placement -ceq 'Not finalized' }).Count } else { -1 })
}
catch {
    Write-Log $_.Exception.Message 'ERROR'
    # A folder that is only partly built must say so to whoever opens it cold.
    if ($script:OutputRoot -and (Test-Path -LiteralPath $script:OutputRoot)) {
        try {
            Set-Content -LiteralPath (Join-Path $script:OutputRoot 'INCOMPLETE - this run failed.txt') -Encoding UTF8 -Value @(
                'This run stopped with an error, so this folder is incomplete and must not be used as a trial master file.',
                "Error: $($_.Exception.Message)",
                'See the run log (a _log_*.txt file in this folder, or - for -FinalOnly - beside it) for the full record.')
        } catch { }
    }
    throw
}
finally {
    # Persist the run log. DATED, so re-running never overwrites the record of the previous run, and
    # so a log that arrives by mail identifies itself. It lands next to the output when there is one,
    # and beside the input otherwise - a run that failed early is exactly when the log matters most.
    $logDir = $null
    if ($script:OutputRoot -and (Test-Path -LiteralPath $script:OutputRoot)) { $logDir = $script:OutputRoot }
    elseif ($script:LogFallbackDir -and (Test-Path -LiteralPath $script:LogFallbackDir)) { $logDir = $script:LogFallbackDir }
    if ($logDir -and -not $DryRun) {      # a dry run writes nothing, the log included
        $logPath = Join-Path $logDir ("_log_{0}.txt" -f $script:RunStamp)
        # A hand-over copy's log names what was left out, so it is kept beside the folder, not in it.
        if ($FinalOnly -and $logDir -eq $script:OutputRoot -and (Split-Path $logDir -Parent)) {
            $logPath = Join-Path (Split-Path $logDir -Parent) ('{0} - _log_{1}.txt' -f (Split-Path $logDir -Leaf), $script:RunStamp)
        }
        try {
            $script:LogLines | Set-Content -LiteralPath $logPath -Encoding UTF8 -ErrorAction Stop
            Write-Host ("  Run log       : {0}" -f $logPath) -ForegroundColor Green
        }
        catch { Write-Host ("  (could not write the run log to {0}: {1})" -f $logPath, $_.Exception.Message) -ForegroundColor Yellow }
    }
    # Remove the temp extraction (every file that was written has been moved or copied out of it by now). Keyed off $script:TempDir,
    # not $source, so a throw *during* extraction still cleans up rather than leaving clinical
    # documents behind in the temp folder.
    if ($script:TempDir -and (Test-Path -LiteralPath $script:TempDir)) {
        try { Remove-Item -LiteralPath $script:TempDir -Recurse -Force -ErrorAction Stop }
        catch { Write-Host ("  (could not remove the temporary folder {0}: {1} - it holds the extracted documents, delete it yourself)" -f $script:TempDir, $_.Exception.Message) -ForegroundColor Yellow }
    }
}

