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

