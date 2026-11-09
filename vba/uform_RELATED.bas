Attribute VB_Name = "uform_RELATED"
Option Explicit

Public Enum Form_Action_Enum
    ADD_WS = 1          ' Copy a work sheet to another workbook
    myQB_ADD_WS = 2     ' Add a balance sheet to MyQBAddin.xlsm workbook
    myQB_COMPARE_WS = 3 ' Compare a QB financial work sheet to active financial work sheet
    myADD_IIF = 4       ' Add/Replace COA or Class List work sheet from IIF file.
End Enum

' Must set PROCESS_ACTION variable before using Form ufrmSheetNames2
'
Dim COPYWB2 As Workbook ' Set to the Workbook that was Active then form is initialized
Public PROCESS_ACTION As Form_Action_Enum


' ----------------------------------------------------------------
' Purpose: Load ufrmWorkBooksOpen user form
'
' Puts up a list of open Workbooks for user selection.
'
'-------------------------------------------------------------------
Function SelectWorkbook_Click() As Workbook
    'Debug.Print "SelectWorkbook_Click(): ..."
    
    ' Show the user form
    ufrmWorkBooksOpen.Show ' vbModeless

    ' THIS IS SO COOL...
    ' Get the selected workbook object from the userform's property
    Set SelectWorkbook_Click = ufrmWorkBooksOpen.GetSelectedWorkbook
    
    ' Unload the form to free up memory
    Unload ufrmWorkBooksOpen

    ' Check if a workbook was selected
    If Not SelectWorkbook_Click Is Nothing Then
        ' User clicked Select. Use the SelectWorkbook_Click variable for your code.
        ' Example: Activate the selected workbook
        ' SelectWorkbook_Click.Activate
        'MsgBox "The workbook '" & SelectWorkbook_Click.name & "' has been selected.", vbInformation
    Else
        ' User clicked Cancel.
        MsgBox "User cancelled. No workbook was selected.", vbExclamation
    End If

End Function

' ----------------------------------------------------------------
' Purpose: Load ufrmWorkSheetsOpen user form
'
' Selects a workbook and then a worksheet if necessary.
'
'-------------------------------------------------------------------
Function SelectWorksheet_Click() As Worksheet
    'Debug.Print "SelectWorksheet(): ..."
    
    '  ******** FIRST, SELECT a workbook *********
    Dim wb As Workbook
    Set wb = SelectWorkbook_Click()
    
    If wb Is Nothing Then
        Set SelectWorksheet_Click = Nothing
        Exit Function
    End If
    
    ' ********* THEN SELECT A Worksheet *************
    ' Check number of worksheets
    If wb.Worksheets.count = 1 Then
        ' Only one worksheet, return it
        Set SelectWorksheet_Click = wb.Worksheets(1)
    Else
        ' Multiple worksheets, prompt user to select
        ufrmWorkSheetsOpen.SetWorkbook wb
        ufrmWorkSheetsOpen.Show
        
        Set SelectWorksheet_Click = ufrmWorkSheetsOpen.GetSelectedWorksheet
        
        Unload ufrmWorkSheetsOpen
    End If
    
    ' Check if a worksheet was selected
    If Not SelectWorksheet_Click Is Nothing Then
        'MsgBox "The worksheet '" & SelectWorksheet_Click.name & "' in workbook '" & wb.name & "' has been selected.", vbInformation
    Else
        MsgBox "No worksheet was selected.", vbExclamation
    End If

End Function



' ******************************************************
'
'   my Original select a sheet form with Process Action Code
'
' ----------------------------------------------------------------
' Purpose: Load ufrmSheetNames2 user form
'
' Puts up a list of sheet names in the Active Workbook for user selection.
'
'-------------------------------------------------------------------
Sub SelectSheetForm_Click()
    Set COPYWB2 = ActiveWorkbook 'Save Active workbook... is newly opened workbook.

    ' *** If only ONE work sheet in workbook. Go straight to process Action ***
    If ActiveWorkbook.Worksheets.count = 1 Then
        Debug.Print "SelectSheetForm_Click(): Only One worksheet in workbook... Calling Process_Selection()."
        Call Process_Selection(ActiveWorkbook.Worksheets(1).name)
        Exit Sub
    End If

10    ufrmSheetNames2.Show vbModeless ' Uses PROCESS_ACTION to determine what to do after selection

End Sub

' *** PROCESS The User Selected Sheet HERE ****
'
' Use the PROCESS_ACTION variable to Determine the Action
' Values of PROCESS_ACTION
'   1. ADD_WS - Adds selected worksheet to OPEN_WB After the Active Sheet Name
'               Activate the added sheet.
'   2. myQB_ADD_WS - Sames as ADD_WS but also Rebuilds the SheetName Table
'
'   3. myQB_COMPARE_WS - Compares two Quick Books Financial Statements for differences.
'
'   4. myADD_IIF - Add/Replace COA or Class List work sheet from IIF file.
'
Sub Process_Selection(selectedSheet As String)
Dim t1 As Date, t_old As Date, t_new As Date, sKey As Variant
Dim bsName As String, currName As String
Dim m As String, ans As Long, new_sheetName As String
Dim acct_cnt As Long, class_cnt As Long
Dim lastRow As Long, lastCol As Long

    t1 = Timer
' NOTES ON CopyWorksheetCloseWB:
' CopyWorksheetCloseWB(wB As Workbook, sheetName As String, beforeafterSheet As String,
'              Optional afterflag As Boolean = False, Optional replaceflag As Boolean = True)
' OPEN_WB as Copy To wb;  wB  = Open workbook to copy From and Close
' sheetName     = Name of the sheet to copy
' beforeafterSheet = Place sheetName before or after this sheet in wB
' afterflag:    default False = Place before the beforeafterSheet in wB
' replaceflag:  default True = Rename existing sheet with "Old ", then replace with copied sheet
' Return the FINAL name of the new sheet in wB

  Select Case PROCESS_ACTION
    Case ADD_WS:
  ' Copy the selected sheet name from the Workbook to the WB set as OPEN_WB global
       OPEN_WB.Activate  ' Activate original WB that was active before myOpenWorkbook macro was called
       currName = ActiveSheet.name
       bsName = CopyWorksheetCloseWB(COPYWB2, selectedSheet, currName, True, False)
       ActiveWorkbook.Worksheets(currName).Activate  ' Activate Sheet again
    Case myQB_ADD_WS:
  ' Copy the selected sheet name from the Workbook to the WB set as OPEN_WB global
       OPEN_WB.Activate  ' Activate original WB
       currName = ActiveSheet.name
       bsName = CopyWorksheetCloseWB(COPYWB2, selectedSheet, currName, False, False)
       ActiveWorkbook.Worksheets(currName).Activate  ' Activate the Sheet again
       
       ' ReBuild the sheet name table
       Call Build_SheetNameTable_Sub(ActiveWorkbook.name)
       Range("Sheet_Name").Value = bsName ' Set to the sheet name imported
    Case myQB_COMPARE_WS:
  ' Compare selected Sheet to OPEN_WB active sheet using QBLookup
       COPYWB2.Activate  ' Activate original WB
' ??? FOR SOME REASON ACTIVE SHEET "TAB NAME" IS NOT DISPLAYED UNTIL AFTER QBlOOKUPCOMPARE !!!!!
       Call QBLookupCompare(OPEN_WB.ActiveSheet.UsedRange, COPYWB2.ActiveSheet.UsedRange)
  ' Insert an IIF worksheet to the OPEN_WB
    Case myADD_IIF:
               
        ' Check we have a valid IIF File
        Call ValidateInIIF(COPYWB2.Worksheets(selectedSheet), acct_cnt, class_cnt)
        If acct_cnt = 0 And class_cnt = 0 Then
            m = "Accounts or Classes are missing from selected IIF file." & vbCrLf
            m = m & "'" & COPYWB2.name & "'" & vbCrLf
            m = m & "Accounts = " & acct_cnt & "  Classes = " & class_cnt & vbCrLf
            m = m & vbCrLf & "Aborting..."
            MsgBox m
            Exit Sub
        End If
        
        ' Call Macro to copy then format the sheet to COA
            ' NOTES ON CopyWorksheets:
            ' CopyWorksheets(from_wB As Workbook, to_wB as Workbook, sheetNames As String, beforeafterSheet As String,
            '              Optional afterflag As Boolean = False, Optional replaceflag As Boolean = True)
            ' Copies SheetNames in from_WB to before or after beforeafterSheet of to_wB workbook
            ' Returns the FINAL names of the new sheets in to_WB as a COMMA SEPARATED STRING
            '
            '    from_wB        = WB to copy FROM
            '    to_WB          = WB to copy TO
            '    sheetNames     = Names of the sheets to copy as a COMMA SEPARATED STRING
            '    beforeafterSheet = Place new sheet before or after this sheet name in to_WB
            '    afterflag:      default False = Place before the beforeafterSheet in to_WB
            '    replaceflag:    default True = Replace if sheetname already exists in to_WB
        bsName = CopyWorksheets(COPYWB2, OPEN_WB, selectedSheet, OPEN_WB.Sheets(1).name) ' replace sheet

        ' *** Format the new Sheet ***
        If bsName <> "" Then new_sheetName = FormatSheet_COA(selectedSheet)
        
        ' Call macro to find the latest date an Account in the OLD COA.
        t_old = GetLatestDateInColumnV(OPEN_WB.Sheets("OLD " & new_sheetName))
        
        ' Call macro to find the latest date an Account in the NEW COA.
        t_new = GetLatestDateInColumnV(OPEN_WB.Sheets(new_sheetName))
         
        ' Reset Highlighting to No Fill on all rows Except row 1
        With OPEN_WB.Sheets(new_sheetName)
            lastRow = Last(1, .Cells)
            lastCol = Last(2, .Cells)
            .Range(.Cells(2, 1), .Cells(lastRow, lastCol)).Interior.ColorIndex = xlColorIndexNone
        End With

        If t_new > t_old Then
            m = "New Accounts have been created since last import. " & vbCrLf & vbCrLf
        Else
            m = ""
        End If
        m = m & "Highlight newly added Accounts in '" & new_sheetName & "'?"
        ans = MsgBox(m, vbQuestion + vbYesNo, "NEWEST DATE TIME >> " & Format(t_new, "mm-dd-yy hh:mm:ss AM/PM"))
        If ans = vbYes Then
            m = "Highlight accounts. Last date in old COA = " & Format(t_old, "mm-dd-yy hh:mm:ss AM/PM") & "?"
            sKey = Application.InputBox("Enter a Date", m, _
                    Format(t_old, "MM-DD-YY"), , , , , 2)
            If sKey = "False" Then Exit Sub
            ' Call macro to highlight rows in column 22 (V) that are newer than sKey
            ans = HighlightRowsAfterDate(OPEN_WB.Sheets(new_sheetName), 22, CDate(sKey))
            m = "Rows Highlighted = " & ans
            MsgBox m
        End If
       ans = MsgBox("IMPORT the CLASS LIST too?", vbQuestion + vbYesNo, "Process_Selection()")
       If ans <> vbYes Then COPYWB2.Close: Exit Sub
        
        ' Call Macro to copy AND CLOSE then format the sheet to CL
       bsName = CopyWorksheetCloseWB(COPYWB2, selectedSheet, "OLD BBM") ' replace sheet
       If bsName <> "" Then Call FormatSheet_Class_List(selectedSheet)
 End Select
    
 Debug.Print "ufrmSheetNames2.Process_Selection(): " & OPEN_WB.ActiveSheet.UsedRange.count & _
    " cells processed, Time elasped = " & Format((Timer - t1), "0.0000") & " Seconds"
    
End Sub

'AI generated
Function GetLatestDateInColumnV(ByVal ws As Worksheet) As Date
Dim lastRow As Long ' variable to store the last row with data
Dim latestDate As Date ' variable to store the latest date found

    ' Find the last row with data in column V
    ' xlUp starts from the bottom of the sheet and goes up until it finds the first non-empty cell
    lastRow = ws.Cells(ws.Rows.count, "V").End(xlUp).row

    ' Check if there is any data in column V
    If lastRow > 0 Then
        ' Use WorksheetFunction.Max to find the latest date in the range V1 to the lastRow
        ' This function ignores non-numeric values and errors, returning the maximum date value
        latestDate = Application.WorksheetFunction.Max(ws.Range("V1:V" & lastRow))
        ' Assign the found latest date to the function's return value
        GetLatestDateInColumnV = latestDate
    Else
        ' If no data is found, return a default empty date or handle the error as needed
        ' For this example, it returns 0, which corresponds to the date 1/0/1900 in Excel
        GetLatestDateInColumnV = 0
    End If
End Function

Private Sub test_ValidateInIIF()
Dim acct_cnt As Long, class_cnt As Long
    Call ValidateInIIF(ActiveSheet, acct_cnt, class_cnt)
    Debug.Print "Accounts = " & acct_cnt & "  Classes = " & class_cnt
End Sub
' AI Generated
    ' This function counts the occurrences of a ACCT and CLASS in Column 1 (A)
    ' of a specified worksheet and returns the counts to the variables passed in
Sub ValidateInIIF(ws As Worksheet, ByRef accts As Long, ByRef classes As Long)

    Dim lastRow As Long
    Dim searchRange As Range

    ' Determine the last row with data in column A of the targetSheet
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row

    ' Define the range to search in (Column A from row 1 to lastRow)
    Set searchRange = ws.Range("A1:A" & lastRow)

    ' Use WorksheetFunction.CountIf to count the occurrences
    accts = Application.WorksheetFunction.CountIf(searchRange, "ACCNT")
    ' Use WorksheetFunction.CountIf to count the occurrences
    classes = Application.WorksheetFunction.CountIf(searchRange, "CLASS")

End Sub

' AI Generated
Function HighlightRowsAfterDate(ws As Worksheet, dateColumn As Long, comparisonDate As Date) As Long
Dim lastRow As Long, lastCol As Long
Dim highlightedCount As Long
Dim i As Long

    highlightedCount = 0
    ' Find the last row with data in the specified column and last column
    lastRow = ws.Cells(ws.Rows.count, dateColumn).End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    
    ' Reset Highlighting to No Fill on all rows Except row 1
    ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Interior.ColorIndex = xlColorIndexNone
    
    ' Loop through each row from the second row (assuming headers in the first)
    For i = 2 To lastRow
        ' Check if the cell contains a valid date and is newer than the comparison date
        If IsDate(ws.Cells(i, dateColumn).Value) And ws.Cells(i, dateColumn).Value > comparisonDate Then
            ' Highlight the entire row in light yellow (ColorIndex 6)
            With ws
                .Range(.Cells(i, 1), .Cells(i, lastCol)).Interior.Color = RGB(255, 255, 0)  ' Light Yellow '6  yellow
            End With
            highlightedCount = highlightedCount + 1
        End If
    Next i
    HighlightRowsAfterDate = highlightedCount
End Function

