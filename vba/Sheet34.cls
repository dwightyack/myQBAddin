VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet34"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Calls are in qb_Import_Mod
Private Sub Go_To_Click()
    Go_To_Selected Selection
End Sub
Private Sub Import_Sheets_Click()
    Import_Sheets_Sub ThisWorkbook.name
End Sub
Sub Build_SheetNameTable_Click()
    Build_SheetNameTable_Sub ThisWorkbook.name
End Sub
Private Sub Calc_Due_Sheet_Click()
    ' DOESNT COMPILE. WORK ON THIS LATER
    ' Calc_WorkSheet ThisWorkbook.name
End Sub
Private Sub Worksheet_Change(ByVal target As Range)
    If target.Address <> Range("Sheet_Name").Address Then Exit Sub
    ' DOESNT COMPILE. WORK ON THIS LATER
    ' Calc_WorkSheet ThisWorkbook.name
End Sub

