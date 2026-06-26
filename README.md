# eTMF Archive Humanizer

Turn a Viedoc **eTMF export** (a `.zip` file) into something you can actually read:

- **tidy folders** you can browse — with the **final / approved** documents kept separate
  from **drafts** and **older (superseded)** versions, and
- a single **web page** you open in your browser to **search, filter and read** every
  document, along with its history, signatures, and who did what and when.

Your original `.zip` is **never changed** — the tool only creates a new, easy-to-read copy.

---

## What you need

- A computer with **PowerShell 7** installed (free from Microsoft; works on Windows, Mac and Linux).
- Your **eTMF export `.zip`** from Viedoc.

Nothing else to install.

## How to use it

1. Put `Convert-EtmfArchive.ps1` somewhere easy to find (e.g. your Desktop).
2. Open **PowerShell 7** in that folder *(in Windows Explorer: hold **Shift**, right-click the
   folder, choose “Open PowerShell window here”)*.
3. Run the tool and point it at your `.zip` file:

   ```powershell
   .\Convert-EtmfArchive.ps1 -ArchivePath "C:\Downloads\My Study eTMF.zip"
   ```

   > **Tip:** instead of typing the path, just **drag your `.zip` file into the PowerShell
   > window** — it fills in the location for you.

4. When it finishes, it tells you the name of the new folder it created. Open that folder and
   **double-click `index.html`** to start browsing. *(Add `-Open` to the command to have it
   open automatically.)*

If you run the tool **without** giving it a `.zip`, it simply prints these instructions.

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

