# eTMF Archive Humanizer

Turn a Viedoc **eTMF export** (a `.zip` file) into something you can actually read:

- **tidy folders** you can browse — one folder per artifact, with each file's status (`Finalized`,
  `Awaiting review`, `Superseded`, …) in its name, and, if you ask for it, the records that are
  **not finalized** kept in a separate tree, and
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
  provide a **digitally signed** copy that runs without any of this.

- **“… requires … PowerShell 5.1 …”** *(rare)* — your PowerShell is older than the version built
  into Windows 10/11. Update Windows, or install **PowerShell 7**
  (`winget install Microsoft.PowerShell`) and run the command with `pwsh` instead of `powershell`.

## What you get

A new folder named after your study:

- **`index.html`** — the viewer. Open it in any browser (no internet needed) to search, filter and read every
  document, with its details, version history, signatures and full audit trail. The search matches the details
  you can see (title, file name, artifact, country/site, status, version, …), not the text inside the PDFs.
- **`Documents/`** — the files, organised the TMF way: Zone → Section → Artifact → Country or Site. The status
  is in each file name, e.g. `Report v2 [Finalized].pdf`.
- **`inventory.csv`** — every document as a spreadsheet (open it in Excel).
- **`About this folder.txt`** — what the folders are and what each status means.
- **`_log_<date>-<time>.txt`** — the full record of the run, including anything the tool could not do.

## Choosing the folder structure

| You want | Add | You get |
|---|---|---|
| Everything in one tree *(default)* | nothing | One folder per artifact holding all its versions. The status is in the file name and in `inventory.csv`. |
| **Approved documents in the TMF, everything not approved outside it** | `-SeparateNotFinalized` | `Documents/` holds every record whose last lifecycle event is `Finalized` (including an older approved version that a newer one replaced). **`Not finalized/`** holds the rest (`Awaiting review`, `Flagged by QC Comment`, or no lifecycle event), with the same Zone/Section/Artifact structure. |
| Status folders as well | `-SeparateByStatus` | `01 Finalized`, `02 Not finalized` and `03 Superseded` folders inside each artifact. Works on its own or with `-SeparateNotFinalized`. |
| **A copy to hand to someone else** | `-FinalOnly` | A separate folder, `<Study> (final documents only)`, with only the latest `Finalized` version of each document. `About this folder.txt` states that it is not the complete TMF. The list of what was left out (`<folder> - withheld.csv`) and the run log are saved **beside** the folder, not in it. |

Example — approved documents in the TMF, the rest outside:

```powershell
powershell -ExecutionPolicy Bypass -File .\Convert-EtmfArchive.ps1 -ArchivePath ".\My Study eTMF.zip" -SeparateNotFinalized
```

With `-SeparateNotFinalized` the complete trial master file is **both folders together**; the tool says so in
`About this folder.txt`, in `Not finalized/README.txt` and at the end of the run. Nothing is ever deleted.
`-FinalOnly` ignores the two layout switches.

## How statuses are decided

Statuses are the export's own words — the tool does not invent any.

- **Lifecycle** is the last of the export's audit events `Awaiting review`, `Flagged by QC Comment` or `Finalized`
  (matched exactly; words inside a comment never count). A record with none has no lifecycle. `Document locked` /
  `Document unlocked` is kept separately in the `Lock` column.
- **Current version:** within a document, the highest version whose lifecycle is `Finalized`. Lower versions are
  `Superseded` (with `SupersededBy` naming the version that replaced them); a newer version that is not finalized
  stays `Not finalized`. Viedoc's own version state is kept untouched in `VersionState`, and `VersionResolution`
  explains any case where the tool and the export differ.
- **Superseded** means *every version lower than the current one, whatever its own status* (v1–v3 of a document whose
  v4 is the current approved version), or a version the export itself marks `Superseded` or `Obsolete`. It is kept, and
  its status shows its own last lifecycle event — `Superseded; Finalized` was approved when it was replaced,
  `Superseded; Awaiting review` was not. A current version that is not approved is `Not finalized`, not Superseded.
- **Placement** for `-SeparateNotFinalized` uses each record's own last lifecycle event, not Superseded: an approved v2
  stays in `Documents/` even though v4 replaced it, and a v3 that was never approved goes to `Not finalized/`.

## Options

| Option | What it does |
|---|---|
| `-ArchivePath` | The eTMF `.zip` (or an already-unzipped export folder). |
| `-OutputPath` | Where to put the result (default: a new folder next to the `.zip`). |
| `-SeparateNotFinalized` | Put records that are not finalized in a sibling `Not finalized/` tree. |
| `-SeparateByStatus` | Add `01 Finalized` / `02 Not finalized` / `03 Superseded` folders inside each artifact. |
| `-FinalOnly` | A hand-over copy with only the latest Finalized version of each document. |
| `-SplitByMilestone` | One folder and viewer per milestone, instead of one combined TMF. |
| `-KeepOriginalNames` | Keep original file names (versions that share a name get ` v<version>`). |
| `-SkipIntegrityCheck` | Skip the MD5 check (faster, less thorough). |
| `-Force` | Replace a result folder from a previous run. Only a folder this tool made is replaced, and never one that is or contains the input. |
| `-Open` | Open the viewer when finished. |
| `-DryRun` | Show what would happen without creating anything. |

For full help: `Get-Help .\Convert-EtmfArchive.ps1 -Full`

---

## License

Provided under the **MIT License** (see the `LICENSE` file included with this package).
It comes with **no warranty and no support — use at your own risk**, and you are responsible
for validating its output for your own purposes.

