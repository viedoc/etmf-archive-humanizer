# eTMF Archive Humanizer

Turn a Viedoc **eTMF export** (a `.zip` file) into something you can actually read:

- **tidy folders** you can browse — with the **final / approved** documents kept separate
  from **drafts** and **older (superseded)** versions, and
- a single **web page** you open in your browser to **search, filter and read** every
  document, along with its history, signatures, and who did what and when.

Your original `.zip` is **never changed** — the tool only creates a new, easy-to-read copy.

---

## What you need

> ✅ **No install needed.** Windows already includes **Windows PowerShell**, and this tool runs on
> it as-is. If you happen to have the newer **PowerShell 7**, that works too — but you don’t need it.

- **A Windows PC** — that’s all you need software-wise. Windows 10 and 11 already include
  **Windows PowerShell**, which is enough to run this tool. *(It also runs on **PowerShell 7**, and
  on macOS/Linux, if you have those.)*
- Your **eTMF export `.zip`** from Viedoc.

Nothing else to install.

## How to use it

1. Put `Convert-EtmfArchive.ps1` **and** your `.zip` in the **same folder** — for example a new
   folder inside Downloads. Keeping them together is the simplest setup: you can then refer to
   each file by just its name.

2. Open **PowerShell**: click **Start**, type **PowerShell**, and open **Windows PowerShell** (the
   one built into Windows — that’s all you need). *(Windows Terminal, or **PowerShell 7** if you’ve
   installed it, work just as well.)*

3. Go to the folder where you put your files. A fresh PowerShell window starts in your personal
   folder (e.g. `C:\Users\You`), so move to your folder with the `cd` command (“change directory”):

   ```powershell
   cd "C:\Users\You\Downloads\tmf"
   ```

   > **Tip:** type `cd ` (with a trailing space), then **drag the folder** from Explorer into the
   > window to fill in its path, and press **Enter**. Your prompt then shows that folder,
   > e.g. `PS C:\Users\You\Downloads\tmf>`.

4. Now run this and press **Enter**. Because both files are in the folder you’re in, the `.\`
   simply means “in this folder”:

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\Convert-EtmfArchive.ps1 -ArchivePath ".\My Study eTMF.zip"
   ```

   Replace `My Study eTMF.zip` with your file’s real name. **Tip:** instead of typing the name,
   just **drag your `.zip` into the window** to fill it in.

   > **Where do the files have to be?** The command assumes the script **and** your `.zip` are in
   > the folder you’re currently in — that’s what `.\` means. If your `.zip` is somewhere else,
   > give its **full path in quotes** instead, e.g. `-ArchivePath "D:\Exports\study.zip"` (or just
   > drag the file in, which always fills the full path). Likewise, if you didn’t `cd` into the
   > script’s folder, put the script’s full path after `-File`.

   The **`-ExecutionPolicy Bypass`** part lets this one unsigned script run this one time — it
   changes nothing on your computer (more in **[If Windows won’t run it](#if-windows-wont-run-it)**
   below). *(Have PowerShell 7? You can write `pwsh` instead of `powershell`.)*

5. When it finishes, it tells you the name of the new folder it created. Open that folder and
   **double-click `index.html`** to start browsing. *(Add `-Open` to the end of the command to
   have it open automatically.)*

If you run the tool **without** giving it a `.zip`, it simply prints these instructions.

## If Windows won’t run it

Windows is cautious about scripts it didn’t create. The command in step 4 avoids the common block,
but if you ever see one of these, here’s what it means:

- **“… cannot be loaded … is not digitally signed”**
  Windows is blocking a script it didn’t create. The **`-ExecutionPolicy Bypass`** in the command
  above allows this one script for this one run only; nothing on your computer is changed permanently.

- **Still blocked even with the full command?** Your organisation may enforce a stricter rule
  (via Group Policy) that overrides `-ExecutionPolicy Bypass`. Ask your IT department — they can
  provide a **digitally signed** copy that runs without any of this (see
  [Technical notes](#technical-notes-for-it)).

- **“… requires … PowerShell 5.1 …”** *(rare)* — your PowerShell is older than the version built
  into Windows 10/11. Update Windows, or install **PowerShell 7**
  (`winget install Microsoft.PowerShell`) and run the command with `pwsh` instead of `powershell`.

## What you get

A new folder named after your study, containing:

- **`index.html`** — the viewer. Open it in any web browser (no internet needed). It shows a
  study summary, a folder tree, a searchable list of documents, and a panel for each document
  with its details, version history, electronic signatures, and full audit trail.
  The search box matches the document details you can see: title, file name, artifact number, zone,
  section, country/site, status, version, Document ID, milestone, why it was flagged and what
  superseded it. You can also narrow by status,
  country, milestone, or by clicking a folder in the tree — and whenever the list is narrowed, a
  bar above it says so and offers **Clear all filters**, so nothing looks missing when it is only
  hidden. (It searches the document *details*, not the text inside the PDFs.)
- **`Documents/`** — the actual files, organised the familiar TMF way (Zone → Section →
  Artifact), and inside each artifact split into clearly-labelled groups:
  - **`01 Final`** — current, finalised/approved documents (inspection-ready)
  - **`02 In Progress`** — current but not yet finalised (drafts / awaiting review)
  - **`03 Superseded`** — older or replaced versions
  - **`04 Excluded`** — only inside `_Excluded/` (see below); never in the TMF tree
- **`inventory.csv`** — a simple spreadsheet listing every document (open it in Excel).
- **`_Excluded/`** — only when you use `-ExclusionList`: the documents you decided do not belong,
  kept with their reason but outside the TMF tree.
- **`_log_<date>-<time>.txt`** — the complete record of the run: the tool version, the archive it
  read, the options it was given, the machine it ran on, everything it printed on screen, and the
  reason for anything it could not do. One log per run, so a re-run never overwrites the previous
  one. **If something looks wrong, send us this file** — it is the fastest way for us to tell what
  happened. (If a run fails before the result folder is created, the log is written next to your
  `.zip` instead.)

### Exports that cover several milestones

An eTMF export usually contains one snapshot **per milestone**, and every snapshot repeats the
whole TMF. The tool merges them into the **one** tree above, so you get a single TMF to browse and
a single place to search — not one folder and one index per milestone:

- a file that appears in several milestone snapshots is copied **once**, and the record kept is the
  one from the **most recent** snapshot (so its status is the latest one the system recorded);
- the milestone is written into `inventory.csv` instead: **`Milestone`** (the snapshot the document
  was taken from) and **`AllMilestones`** (every snapshot it appeared in);
- if you would rather have the old view — one folder and one viewer per milestone, with a front
  page linking them — add **`-SplitByMilestone`**. Note that this view shows each milestone exactly
  as the system exported it, so it does **not** work out which version supersedes which (that needs
  the milestones side by side). Your flags and `-ExclusionList` decisions still apply.

### Which version is the current one

An export marks a document *Current* as of the milestone it was taken from — so version 1, exported
at an early milestone, still says "Current" there even though version 2 replaced it later. Merging
lets the tool settle that: within each document, the newest version the archive still calls Current
is the one that stays in **`01 Final`**, and every earlier version is filed under
**`03 Superseded`** with an `inventory.csv` column, **`SupersededBy`**, naming the version that
replaced it (e.g. `v2 (Finalized, 09 Database Lock)`). The viewer shows the same under
**Superseded by** in a document's details.

Nothing is deleted or hidden — an archive has to stay complete — the older versions are simply out
of the way of the current ones. Documents whose several *files* belong to the same version stay
together.

### A copy to hand to a client

The output above is the **complete** rendering: current documents, drafts and superseded versions
together, which is what you want for your own review and for an inspection. When you need to hand
something to a client or sponsor, add **`-FinalOnly`**:

```powershell
.\Convert-EtmfArchive.ps1 -ArchivePath ".\My Study eTMF.zip" -FinalOnly
```

You get a separate folder — `<Study> (final documents only)` — containing **only** the documents the
eTMF records as the current version *and* as finalised or locked. Superseded versions, drafts and
anything you excluded are left out, and the status subfolders disappear (everything in there is
final, so a `01 Final` folder under every artifact would say nothing).

It also writes **`About this folder.txt`**, which tells the recipient what the folder is, which
export and milestone it came from, when and with which version it was prepared, how many documents
of each kind were left out, and — in plain words — that this is a convenience copy and **not** the
complete trial master file. That note is as much for you as for them: it is your record of exactly
what you handed over.

Your original archive and the complete rendering are untouched; run the tool without the switch to
produce those.

### Documents that should not be there

The export contains everything ever uploaded and not deleted, so it can hold documents that do not
belong in this TMF at all. **No export field says so** — that is a judgement about the study, not a
fact in the data — so the tool does two separate things and never confuses them:

**It flags, and moves nothing.** Two kinds of flag appear in `inventory.csv`'s **ReviewFlag** column
and behind the viewer's **⚑ Flagged for review** filter:

- *Filed under two artifacts* — the identical file appears in two different places in the TMF, which
  is a filing question someone has to answer. Exact, not a guess: one checksum, two artifacts. The
  same file reused across countries under one artifact (a cover letter template) is normal and is
  **not** flagged.
- *Possible foreign study* — the file name mentions something shaped like another study's code
  (e.g. `XYZ-002` when the study is `TEST-001`). Deliberately narrow, so ordinary document
  numbering such as `Consent Form-01` or `Amendment-001` never trips it. It is a prompt to look,
  never a verdict.

**You decide, and it records your decision.** Put the documents that do not belong into a CSV and
pass it with `-ExclusionList`:

```csv
DocumentId,Version,ArtifactNumber,Country,SiteId,Reason
1234,,,,,Belongs to study XYZ-002 - uploaded to this TMF in error
5678,2,,,,Superseded scan uploaded twice
9012,,01.01.02,DEU,,German copy only - the Swedish one is correct
```

`DocumentId` is required (it is in `inventory.csv` and in the viewer); every other column is
optional and narrows the row. **Use them when it matters:** the same `DocumentId` can be reused for
the same document filed in several countries, so a row with only an id may remove more than you
meant. Every excluded document is named individually in the run log, and a row that hits several
different documents warns you. `Reason` is recorded in the output. Those documents are filed under
**`_Excluded/`** — kept, with the reason, but out of the `Documents/` tree you hand to a client —
and they appear in the viewer only under the **Excluded** filter. You are never left guessing: the
count is always in the header, and if a search or filter would have shown an excluded document, the
bar above the list says so and offers one click to show it. A row that matches no document is reported, never silently ignored. Re-running with the
same CSV reproduces the same result, so the decision is made once and stays made.

In the viewer, the **Milestone** filter answers "what did the TMF hold at milestone N" — it shows
every document that was in that export. The **Last export** column shows the most recent milestone
that still contained the document, so a document that was removed from the TMF along the way is
easy to spot: its last export is an older milestone than the others.

Milestones are named after the manifest they come from (e.g. `09 Database Lock`), which is the name
the study uses — not the internal event number in the export.

## How it decides “final” vs the rest

The tool reads each document’s history from the export:

- finalised or locked → **Final**
- still a draft or awaiting review → **In Progress**
- replaced by a newer version → **Superseded**

It also **checks every file is intact** (not corrupted or truncated) and flags anything that is
missing or doesn’t match, so you can trust the copy.

If a document listed in the export has no file behind it, the tool never hides it: the document
still appears (marked ⚠), and the reason is written in three places — `inventory.csv`’s **Note**
column, the run log (with the exact path it looked for), and a summary at the end of the run that
groups the gaps by cause, e.g. *“the archive contains no entry with this exact path”*.

## Options

| Option | What it does |
|---|---|
| `-ArchivePath` | The eTMF `.zip` (or an already-unzipped export folder). |
| `-OutputPath` | Where to put the result (defaults to a new folder next to your `.zip`). |
| `-Open` | Open the viewer in your browser when finished. |
| `-Force` | Replace a result folder from a previous run. |
| `-FinalOnly` | A copy to hand over: current approved documents only, plus a note stating its scope. |
| `-ExclusionList` | CSV of documents that do not belong in this TMF; they are filed under `_Excluded/`. |
| `-SplitByMilestone` | One folder + viewer per milestone snapshot, instead of one merged TMF. |
| `-KeepOriginalNames` | Keep the original file names instead of tidy, readable ones. |
| `-SkipIntegrityCheck` | Skip the file-integrity check (faster, less thorough). |
| `-DryRun` | Show what would happen without creating anything. |

For full help, run: `Get-Help .\Convert-EtmfArchive.ps1 -Full`

---

## Technical notes (for IT)

- **Runs on PowerShell 5.1+** (`#requires -Version 5.1`) — both the **Windows PowerShell 5.1** built
  into Windows 10/11 and **PowerShell 7**. No external modules to install. PowerShell 7 is the
  primary-tested target; 5.1 is supported (the script uses no 7-only syntax, and hardens the XML
  parser against XXE, which .NET Framework does not do by default).
- **Launching / execution policy.** Start it with `powershell` (built-in) or `pwsh` (PowerShell 7).
  The published script is **unsigned**, so on machines that block unsigned scripts add
  `-ExecutionPolicy Bypass` (process-scoped — no persistent change):
  `powershell -ExecutionPolicy Bypass -File .\Convert-EtmfArchive.ps1 -ArchivePath "…zip"`.
- **Code signing (managed machines).** Where Group Policy enforces `AllSigned`,
  `-ExecutionPolicy Bypass` is *ignored*; sign the script with your organisation’s code-signing
  certificate as the **final** step (re-signing after any rebuild, which regenerates the file):
  `Set-AuthenticodeSignature -FilePath .\Convert-EtmfArchive.ps1 -Certificate $cert -TimeStampServer http://timestamp.digicert.com`.
  Recipients can confirm it with `Get-AuthenticodeSignature .\Convert-EtmfArchive.ps1`.
- Reads the **eTMF Exchange Mechanism Standard (eTMF-EMS)** manifest (`*_exchange.xml`) and the
  referenced files; classification is **per document (object)** on two axes — lifecycle (from
  audit events) and version state.
- **File integrity** is verified with base64-encoded MD5 on both the source and the copy.
- The viewer is a **single self-contained HTML file** (inline CSS/JS, no external/CDN
  resources, opens offline). The manifest is treated as **untrusted**: paths are contained to
  prevent traversal, and all document text is escaped before display.
- An **MD5 checksum** (`Convert-EtmfArchive.ps1.md5`) ships next to the script so you can confirm it
  is intact: `Get-FileHash Convert-EtmfArchive.ps1 -Algorithm MD5` (PowerShell) or
  `md5sum -c Convert-EtmfArchive.ps1.md5` (Linux/Mac).
- The published script is **built** from modular sources — see `docs/DEVELOPMENT.md` and
  `CLAUDE.md` if you’re working on the code.

## License

Provided under the **MIT License** (see the `LICENSE` file included with this package).
It comes with **no warranty and no support — use at your own risk**, and you are responsible
for validating its output for your own purposes.

