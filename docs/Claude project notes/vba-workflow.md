# VBA change workflow for myQBAddin.xlam

Started 2026-11-09 from the general rules of the NS_Trade_Historical project. `CLAUDE.md` in the project folder holds the condensed rule set; this note keeps the reasons and history. Master copy: the Claude project "myQBAddin" (`claude/vba-workflow.md`).

## How a change request goes
1. Dwight saves the add-in in Excel (Ctrl+S in the VBA editor with myQBAddin selected, or `Save_Addin`).
2. Claude reads `myQBAddin.xlam` from `%APPDATA%\Microsoft\AddIns` (never an earlier extract, never `vba\`) and tells him its last-saved date and time.
3. Claude changes as few lines as possible, keeping his edits and comments, and delivers `.bas` files (paste-in `.txt` for form / sheet / ThisWorkbook code) ALWAYS zipped as `Claude Outputs\myQBAddin_VBA_<date>.zip`, also sent in the chat.
4. Dwight imports the modules into the live add-in, compiles, and saves with `Save_Addin` (backs up to AddIn_Backups, publishes to AddIns_Shared).
5. Claude refreshes `vba\` from the saved add-in once the change is in, writes the exchange log and commits locally. Dwight pushes (or asks Claude to).

## Code rules
- Never use AutoFit (columns or rows) unless Dwight asks for it (Dwight, 2026-11-09; same rule he set for NS_Trade).

## Used by other workbooks
- Many workbooks reference this add-in; public names and signatures are an interface. Search `C:\Projects` for callers before changing one.
- NS_Trade's NSCommon calls `whichOpen_WB_Dialog` and `whichOpen_CSV_Dialog` (2026-11-09).
- `ufrmSheetNames2` exists here and in NSCommon as separate copies.

## Code comment style (from NS_Trade, Dwight 2026-10-06 / 10-07)
- `'====` header block per module and per procedure: `Procedure :`, `Module :` (or `Form :`), `Workbook  : myQBAddin.xlam`, `'----` rule, then PURPOSE, PARAMETERS, RETURNS, HOW IT WORKS (numbered), USAGE, NOTES / ASSUMPTIONS, REVISION HISTORY as they apply.
- Inline comments where logic is not obvious.
- REVISION HISTORY authors: first entry `[DY]`, later entries `[DY & Claude]`.
- Headers are added to procedures Claude changes; untouched code is not restyled unless asked.

## Reply style (from NS_Trade, 2026-10-03 / 10-18)
- No boilerplate notice. Give zip name, modules, last-saved time read, what changed.
- Always end with loading steps for that delivery.

## Exchange logs (from NS_Trade, 2026-10-03 / 10-17)
- One Word file per exchange in `Claude Exchanges\`, `YYYY-MM-DD HHMM - <topic>.docx` (America/Chicago, time of the request): Request, Your input (verbatim), Response, Files. Committed locally.

## Git
- `C:\Projects\myQBAddin` initialised 2026-11-09, branch `main`. Remote `https://github.com/dwightyack/myQBAddin.git` (private), first pushed by Dwight from Git Bash 2026-11-09 1:04 pm. Dwight pushes (`cd /c/Projects/myQBAddin`, `git push`); the linked-computer shell has no route to github.com. Claude can push when asked: bundle `origin/main..main` on the computer, stage it, clone with push access in the cloud session, pull the bundle, push, then `git update-ref refs/remotes/origin/main main` on the computer and delete the bundle (first done 2026-11-09 1:16 pm). New workbooks / CSVs in the folder are committed without asking. Git in the linked-computer shell needs delete permission on the folder.

## VBA lessons carried over from NS_Trade
- Max 24 line continuations per statement.
- Variable names matching object members get recased by the editor (`fileName`, `lines`, `action`, `report`).
- Save `Err.Number` / `Err.Description` before calling other error-handled code.
- An `.xlam` cannot be saved while `IsAddin` is False.
- ActiveX buttons added from code can silently end the running macro.

## History
- 2026-11-09: project folder set up; `vba\` holds Dwight's own export of the add-in (59 files, 12,856 lines in .bas/.cls/.frm).
- 2026-11-09 1:04 pm: first push to GitHub (4 commits).
- 2026-11-09 1:16 pm: AutoFit rule added.
- 2026-11-09 1:16 pm: Claude pushed 3 commits through the GitHub connector (bundle route).
