# myQBAddin - project rules

Rules for any Claude working in this folder (`C:\Projects\myQBAddin`). The owner
is Dwight Yackley (GitHub user ID `dwightyack`). Fuller notes are in
`docs/Claude project notes/`. The general rules come from the NS_Trade_Historical
project (`C:\Projects\NS_Trade\CLAUDE.md`, 2026-11-09); NS_Trade-specific rules
(NSCommon, shared\, CmdButtonsBuilder, SetHostWorkbook) do not apply here.

## Two folders, two jobs

| Location | Job | Claude |
|----------|-----|--------|
| `%APPDATA%\Microsoft\AddIns\myQBAddin.xlam` (`C:\Users\<user>\AppData\Roaming\Microsoft\AddIns`) | The live add-in. Excel loads it from here on every computer; Dwight edits and saves it here. | Read only. Of the files in that folder, touch only `myQBAddin.xlam`. Request access to it separately when a code request needs it. |
| `\\NAS\BBMDATA\TEMPLATE\AddIns_Shared\` | Published copy, written by `Save_Addin` (Ctrl+Shift+S, `my_Addins_Mod`) with a `.Saved` suffix. `CheckForWbUpdates` offers newer copies from here to other computers. | Never write. |
| `\\NAS\BBMDATA\TEMPLATE\AddIn_Backups\` | Backups `Save_Addin` writes before each save. | Never write. |
| `C:\Projects\myQBAddin` | This Git project: `vba\` export, outputs, exchange logs, notes. | Read and write. |

## What is here

| Path | Contents |
|------|----------|
| `vba/` | Text export of every component of the add-in (`.bas`, `.cls`, `.frm` + `.frx`). |
| `tools/export_vba.py` | Regenerates `vba/` from the saved add-in: `python tools/export_vba.py "<path>\myQBAddin.xlam" vba` (copied from NS_Trade). |
| `Claude Outputs/` | Zips and other files Claude has delivered. |
| `Claude Exchanges/` | One Word log per Claude exchange. |
| `docs/Claude project notes/` | Copies of the project notes kept in the Claude project "myQBAddin". |

## Used by other workbooks

- Many workbooks call this add-in through a VBA reference, so keep public procedure names, argument lists and return types unchanged unless Dwight asks. Before changing one, search the other projects in `C:\Projects` for callers and tell him what was found.
- Known callers (2026-11-09): NS_Trade's `NSCommon.xlam` calls `myQBAddin.whichOpen_WB_Dialog` (`COPY_WS_Pg_Setup`) and `myQBAddin.whichOpen_CSV_Dialog` (`CommonUtils_myQB`).
- Both this add-in and NSCommon have a form named `ufrmSheetNames2`. They are separate copies; do not mix their code.

## Git and GitHub

- Remote: `https://github.com/dwightyack/myQBAddin.git` (private, first pushed 2026-11-09). Branch `main`. Keep it private.
- Dwight pushes from Git Bash (`cd /c/Projects/myQBAddin`, `git push`). The linked-computer shell has no route to github.com.
- Always use the user ID `dwightyack` in any git or gh command shown to Dwight.
- Commit any NEW `.csv` file or Excel workbook (`.xlsx`, `.xlsm`, `.xlsb`, `.xls`, `.xlam`) found in the folder without asking, and say which files were committed. Excel `~$` lock files stay ignored.
- Everything Claude produces is saved in this folder and committed: zips and deliverables in `Claude Outputs/`, exchange logs in `Claude Exchanges/`, notes in `docs/Claude project notes/`.
- Push when Dwight asks for it.
- Do not commit while another Claude or Dwight is committing in the same folder.
- Git in the linked-computer shell needs delete permission on this folder (it removes its own lock files); request it before running git.

## Working with the add-in

- The add-in is normally open in Excel, so never write to it. Deliver VBA as `.bas` modules (and paste-in `.txt` files for form, sheet and `ThisWorkbook` code), ALWAYS zipped, in `Claude Outputs/` as `myQBAddin_VBA_<date>.zip` (`_<date>b` for a second one the same day). Dwight imports the code himself.
- Read the code from the saved `myQBAddin.xlam` at the start of every code request; never reuse an earlier extract, and do not treat `vba/` as current. Edits typed in the VBA editor but not saved (Ctrl+S) are not in the file.
- Tell Dwight the last-saved date and time of the copy that was read.
- Change as few lines as possible. Keep Dwight's own edits and comments.
- After a delivery is confirmed loaded, refresh `vba/` from the saved add-in and commit it.

## Code comment style

- All code gets comment headers in the NS_Trade `ChartsAndGraphs` form: a `'====` ruled block for each module and each procedure with `Procedure :`, `Module :` (or `Form :`), `Workbook  : myQBAddin.xlam`, a `'----` rule, then the sections that apply: PURPOSE, PARAMETERS, RETURNS, HOW IT WORKS (numbered steps), USAGE, NOTES / ASSUMPTIONS, REVISION HISTORY.
- Add headers to a procedure when Claude changes it; Dwight's existing comment text goes under PURPOSE verbatim. Do not restyle untouched code unless asked.
- Add inline comments wherever the logic is not obvious, so another programmer can follow it.
- Authors in REVISION HISTORY: the first entry (initial version) is `[DY]`; every later entry is `[DY & Claude]`.

```vb
' REVISION HISTORY
'   Date        Author         Description
'   ----------  -------------  ----------------------------------------------
'   2026-10-05  [DY]           Initial version
'   2026-10-05  [DY & Claude]  X-axis switched to date scale, monthly ticks
'==============================================================================
```

## VBA lessons (from NS_Trade)

- A statement may have at most 24 line continuations.
- Avoid variable names that match object members (`fileName`, `lines`, `action`, `report`): the VBA editor recases the member everywhere.
- In an error handler, copy `Err.Number` / `Err.Description` to variables before calling anything with its own `On Error`.
- An `.xlam` cannot be saved while `IsAddin` is False (error 1004); set `IsAddin = True` first.
- ActiveX buttons added to a sheet from code can end the running macro silently.
- Sheet and `ThisWorkbook` code must be pasted into the existing module, not imported.

## Replies to Dwight

- Do not repeat the standard notice (changes are not in the add-in yet, not compiled, replace these modules, run Debug > Compile). He knows the process.
- For a code delivery give: the zip file name, which modules are in it, the last-saved time of the add-in copy read, and what changed.
- ALWAYS end a code delivery with short, specific loading steps for that delivery: extract the zip in `Claude Outputs/`, which module to remove and import (or which code to paste where), Debug > Compile myQBAddin, then save with `Save_Addin` (Ctrl+Shift+S) so it is backed up and published. Mention anything to edit by hand.

## Exchange logs

- After every exchange, save a Word file in `Claude Exchanges/` named `YYYY-MM-DD HHMM - <short topic>.docx` (Dwight's local time, America/Chicago, when the request was sent) and commit it.
- Contents: title and date/time, then four sections: Request (a short summary of what he asked), Your input (verbatim) (his messages exactly as typed - every message of the exchange, including answers to Claude's questions, and the full text of any attached text file; separate messages with a --- line; a screenshot is noted as [Screenshot attached: ...]), Response (what was answered or changed), Files (what was delivered and where).
- Screenshots are described, not embedded.
- Do this without being asked and without commenting on it at length.

## Keeping these rules current

- When Dwight adds or changes a project rule, update this file and `docs/Claude project notes/` (and the notes in the Claude project) in the same exchange and commit them.
