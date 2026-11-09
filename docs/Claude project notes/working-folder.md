# Working folders for myQBAddin

## Project folder (since 2026-11-09)
- `C:\Projects\myQBAddin` - the Git project folder: `vba\` export, `tools\`, Claude's outputs and exchange logs, notes, `CLAUDE.md`. If a chat is linked to Dwight's computer and this folder is not connected, request access to it before starting work.

## Live add-in
- `C:\Users\<user>\AppData\Roaming\Microsoft\AddIns\myQBAddin.xlam` (`%APPDATA%\Microsoft\AddIns`). Excel loads it from here on every computer; Dwight's save and publish macros expect it here. Claude reads it only, and touches no other file in that folder. Request access to that folder separately, read only in practice, when a code request needs the current code.

## NAS folders used by Save_Addin (my_Addins_Mod)
- `\\NAS\BBMDATA\TEMPLATE\AddIns_Shared\` - published copy (`.Saved`); `CheckForWbUpdates` compares against the newest file here.
- `\\NAS\BBMDATA\TEMPLATE\AddIn_Backups\` - backup of the old copy before each save (`.<computer>.<user>` suffix).
- Network folders could not be connected from a cloud chat in NS_Trade (2026-10-06); only ask for them when something there is needed.
