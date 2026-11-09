Attribute VB_Name = "ResetLastCell"
Option Explicit

'----------------------------------------------------------------------
' Resets the "end" of a worksheet (what Ctrl+End jumps to) so it matches
' the last cell that actually contains a value or formula.
'
' It deletes every row below and every column to the right of the real
' data, which clears leftover formatting, then forces Excel to recalculate
' the used range.
'
' NOTE: Macros cannot be undone with Ctrl+Z. Save a copy first.
'----------------------------------------------------------------------

' Run this to reset the active sheet only
Public Sub ResetLastCell_ActiveSheet()
    ResetLastCell ActiveSheet
    MsgBox "Last cell reset to " & _
           ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Address(False, False), _
           vbInformation
End Sub

' Run this to reset every worksheet in the active workbook
Public Sub ResetLastCell_AllSheets()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ResetLastCell ws
    Next ws
    MsgBox "Last cell reset on all worksheets.", vbInformation
End Sub

' Core routine
Public Sub ResetLastCell(ByVal ws As Worksheet)
    Dim lastRow As Long, lastCol As Long
    Dim rFound As Range

    If ws.ProtectContents Then
        Debug.Print "Skipped protected sheet: " & ws.name
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Find the last row containing any value or formula
    Set rFound = ws.Cells.Find(what:="*", After:=ws.Cells(1, 1), _
                               LookIn:=xlFormulas, LookAt:=xlPart, _
                               SearchOrder:=xlByRows, SearchDirection:=xlPrevious)

    If rFound Is Nothing Then
        ' Sheet is empty - clear everything so the end resets to A1
        lastRow = 0
        lastCol = 0
    Else
        lastRow = rFound.row
        ' Find the last column containing any value or formula
        lastCol = ws.Cells.Find(what:="*", After:=ws.Cells(1, 1), _
                                LookIn:=xlFormulas, LookAt:=xlPart, _
                                SearchOrder:=xlByColumns, _
                                SearchDirection:=xlPrevious).Column
    End If

    ' Delete unused rows below the data
    If lastRow < ws.Rows.count Then
        ws.Range(ws.Rows(lastRow + 1), ws.Rows(ws.Rows.count)).Delete
    End If

    ' Delete unused columns to the right of the data
    If lastCol < ws.Columns.count Then
        ws.Range(ws.Columns(lastCol + 1), ws.Columns(ws.Columns.count)).Delete
    End If

    ' Force Excel to recalculate the used range
    Dim dummy As Long
    dummy = ws.UsedRange.Rows.count

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
