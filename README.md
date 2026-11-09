# myQBAddin

Dwight Yackley's personal Excel add-in, `myQBAddin.xlam`: QuickBooks report
formatting, IIF import/export, reading QB spreadsheets, BRC list imports, plus
a library of general VBA procedures and UDFs that many other workbooks use.

## Where things live

| Location | What it is |
|----------|------------|
| `%APPDATA%\Microsoft\AddIns\myQBAddin.xlam` | The live add-in. Excel loads it from here on every computer. Edited and saved here. |
| `\\NAS\BBMDATA\TEMPLATE\AddIns_Shared\` | Published copy. `Save_Addin` (Ctrl+Shift+S) copies the saved add-in here; other computers pick up newer copies from here (`CheckForWbUpdates`). |
| `\\NAS\BBMDATA\TEMPLATE\AddIn_Backups\` | Backups written by `Save_Addin` before each save. |
| `C:\Projects\myQBAddin` (this folder) | Git project: text export of the VBA, Claude's outputs and exchange logs, notes. |

## Contents of this folder

| Path | What it is |
|------|------------|
| `vba/` | Text export of every VBA component in the add-in (`.bas`, `.cls`, `.frm`/`.frx`), for diffs and history. |
| `tools/export_vba.py` | Regenerates `vba/` from the saved add-in (Python 3, standard library, no Excel needed). |
| `Claude Outputs/` | Zipped `.bas` deliveries and other files Claude produces. |
| `Claude Exchanges/` | One Word log per Claude exchange. |
| `docs/Claude project notes/` | Copies of the project notes kept in the Claude project. |
| `CLAUDE.md` | Rules Claude follows in this folder. |

The `.xlam` is the source of truth. `vba/` is a readable copy of its code.

## VBA components (2026-11-09 export)

Standard modules:

| Module | Role (from its own comments) |
|--------|------|
| `QB_Reports` | Formats QuickBooks report output (e.g. Transaction Detail by Account). |
| `Import_Export_IIF_Format`, `Import_IIF_OLD_WAY` | IIF import/export and sheet formatting. |
| `qb_Import_Mod` | Copies QB worksheets between workbooks. |
| `BRC_Import_Lists`, `BRC_Interface` | BRC List Reporter imports and helpers. |
| `UDFs_QB_WB` | UDFs used by other workbooks through a reference to this add-in. |
| `CommonUtils_myQB` | `whichOpen_CSV_Dialog` / `whichOpen_CSV` (used by NSCommon). |
| `COPY_WS_Pg_Setup` | Copying sheets and page setup; `whichOpen_WB_Dialog` / `whichOpen_WB` (used by NSCommon). |
| `CompareWB` | Compares the active sheet with a sheet in another file. |
| `Hyper_Link_Mod` | Hyperlinks for K-1s, QB accounts and other files. |
| `modGoToSelected` | Go to the selected sheet/workbook; hides add-ins it un-hid. |
| `my_Addins_Mod` | Add-in management: view/hide, `Save_Addin` (backup, save, publish), update check. |
| `my_Exec_Mod` | Exec macros (some lines are PC specific). |
| `my_ribbon_Ctx`, `my_ribbon_Ctx_MISC` | Custom ribbon callbacks. |
| `uform_RELATED` | Support for `ufrmSheetNames2`. |
| `ResetLastCell`, `SCRUB_Doc_Mod`, `ZIP_FILE_FOLDER_EMAIL`, `Useful_AI_Generated_Stuff` | Utilities. |
| `zEXPORT_IMPORT_VBA`, `zImport_VBA_Mod` | Export/import of VBA components. |

Forms: `ufrmSheetNames2`, `ufrmWorkBooksOpen`, `ufrmWorkSheetsOpen`.
Also `ThisWorkbook.cls` and one `SheetNN.cls` per worksheet (only `Sheet17` and
`Sheet34` hold code).

## Workflow

1. Save the add-in in Excel (Ctrl+S in the VBA editor with myQBAddin selected,
   or `Save_Addin`).
2. Refresh the text export (run from this folder):

   ```
   python tools/export_vba.py "%APPDATA%\Microsoft\AddIns\myQBAddin.xlam" vba
   ```

3. Review and commit:

   ```
   git status
   git diff
   git add -A
   git commit -m "Describe the change"
   git push
   ```
