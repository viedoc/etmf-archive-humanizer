# eTMF Archive Humanizer

Turn a Viedoc **eTMF export** (a `.zip` file) into something you can actually read:

- **tidy folders** you can browse — with the **final / approved** documents kept separate
  from **drafts** and **older (superseded)** versions, and
- a single **web page** you open in your browser to **search, filter and read** every
  document, along with its history, signatures, and who did what and when.

Your original `.zip` is **never changed** — the tool only creates a new, easy-to-read copy.

---

## What you need

> ⚠️ **You must use PowerShell 7 — *not* the “Windows PowerShell” that comes with Windows.**
> The built-in “Windows PowerShell” is version **5.1**, and it will refuse to run this tool with an
> error like *“…requires…PowerShell 7.0…currently running…5.1”*. That is expected — it simply
> means you opened the wrong one. Installing and using **PowerShell 7** (steps below) is the fix.

- **PowerShell 7** — free from Microsoft; runs on Windows, Mac and Linux. This is **not** the
  “Windows PowerShell” already built into Windows (that older 5.1 version won’t run this tool).
  Install it with `winget install Microsoft.PowerShell` or from Microsoft’s website; you then run
  it by typing **`pwsh`**.
- Your **eTMF export `.zip`** from Viedoc.

Nothing else to install.

## How to use it

1. Put `Convert-EtmfArchive.ps1` **and** your `.zip` in the **same folder** — for example a new
   folder inside Downloads. Keeping them together is the simplest setup: you can then refer to
   each file by just its name.

2. Open **PowerShell 7** — this is the app you need, **not** “Windows PowerShell” (that’s the old
   5.1). Either:
   - **Start menu:** click **Start**, type **PowerShell 7**, and open the app of that exact name.
   - **Windows Terminal:** open **Terminal**, click the small **down-arrow (⌄)** next to the **+**
     (new-tab) button at the top, and choose **PowerShell 7**.

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
   pwsh -ExecutionPolicy Bypass -File .\Convert-EtmfArchive.ps1 -ArchivePath ".\My Study eTMF.zip"
   ```

   Replace `My Study eTMF.zip` with your file’s real name. **Tip:** instead of typing the name,
   just **drag your `.zip` into the window** to fill it in.

   > **Where do the files have to be?** The command assumes the script **and** your `.zip` are in
   > the folder you’re currently in — that’s what `.\` means. If your `.zip` is somewhere else,
   > give its **full path in quotes** instead, e.g. `-ArchivePath "D:\Exports\study.zip"` (or just
   > drag the file in, which always fills the full path). Likewise, if you didn’t `cd` into the
   > script’s folder, put the script’s full path after `-File`.

   The two extra words at the start get you past Windows’ safety checks (details in
   **[If Windows won’t run it](#if-windows-wont-run-it)** below): `pwsh` runs it with
   **PowerShell 7**, and `-ExecutionPolicy Bypass` lets this one unsigned script run this one
   time — it changes nothing on your computer.

5. When it finishes, it tells you the name of the new folder it created. Open that folder and
   **double-click `index.html`** to start browsing. *(Add `-Open` to the end of the command to
   have it open automatically.)*

If you run the tool **without** giving it a `.zip`, it simply prints these instructions.

## If Windows won’t run it

Windows is cautious about scripts it didn’t create. The command in step 4 is built to avoid the
two common blocks — but if you ever see one of these errors, here’s what it means and why the
command fixes it:

- **“… requires … PowerShell 7.0 … currently running … 5.1”**
  You launched the *old* PowerShell. The **first word** of the command picks the version:
  **`pwsh`** = PowerShell 7 (what you want); **`powershell`** = the old 5.1 built into Windows.
  Make sure your command starts with `pwsh`. *(If Windows replies that `pwsh` isn’t found, then
  PowerShell 7 isn’t installed — see [What you need](#what-you-need).)*

- **“… cannot be loaded … is not digitally signed”**
  Windows is blocking a script it didn’t create. The **`-ExecutionPolicy Bypass`** part allows
  this one script for this one run only; nothing on your computer is changed permanently.

- **Still blocked even with the full command?** Your organisation may enforce a stricter rule
  (via Group Policy) that overrides `-ExecutionPolicy Bypass`. Ask your IT department — they can
  provide a **digitally signed** copy that runs without any of this (see
  [Technical notes](#technical-notes-for-it)).

## What you get

A new folder named after your study, containing:

- **`index.html`** — the viewer. Open it in any web browser (no internet needed). It shows a
  study summary, a folder tree, a searchable list of documents, and a panel for each document
  with its details, version history, electronic signatures, and full audit trail.
- **`Documents/`** — the actual files, organised the familiar TMF way (Zone → Section →
  Artifact), and inside each artifact split into three clearly-labelled groups:
  - **`01 Final`** — current, finalised/approved documents (inspection-ready)
  - **`02 In Progress`** — current but not yet finalised (drafts / awaiting review)
  - **`03 Superseded`** — older or replaced versions
- **`inventory.csv`** — a simple spreadsheet listing every document (open it in Excel).
- **`_log.txt`** — a record of what the tool did.

## How it decides “final” vs the rest

The tool reads each document’s history from the export:

- finalised or locked → **Final**
- still a draft or awaiting review → **In Progress**
- replaced by a newer version → **Superseded**

It also **checks every file is intact** (not corrupted or truncated) and flags anything that is
missing or doesn’t match, so you can trust the copy.

## Options

| Option | What it does |
|---|---|
| `-ArchivePath` | The eTMF `.zip` (or an already-unzipped export folder). |
| `-OutputPath` | Where to put the result (defaults to a new folder next to your `.zip`). |
| `-Open` | Open the viewer in your browser when finished. |
| `-Force` | Replace a result folder from a previous run. |
| `-KeepOriginalNames` | Keep the original file names instead of tidy, readable ones. |
| `-SkipIntegrityCheck` | Skip the file-integrity check (faster, less thorough). |
| `-DryRun` | Show what would happen without creating anything. |

For full help, run: `Get-Help .\Convert-EtmfArchive.ps1 -Full`

---

## Technical notes (for IT)

- **Requires PowerShell 7.0+** (`#requires -Version 7.0`); no external modules to install.
- **Launching on Windows / execution policy.** Start it with `pwsh` (PowerShell 7 — *not* the
  built-in `powershell`, which is 5.1). The published script is **unsigned**, so on machines that
  block unsigned scripts add `-ExecutionPolicy Bypass` (process-scoped — it makes no persistent
  change): `pwsh -ExecutionPolicy Bypass -File .\Convert-EtmfArchive.ps1 -ArchivePath "…zip"`.
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

