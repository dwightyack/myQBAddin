Attribute VB_Name = "CompareWB"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Compares Active workseet to a worksheet from a second file you select.
'
'   Currently, the compare uses the QBLookup() function. (may modify that later)
'   So, this sub is good for comparing two balance or financial statements.
'   Financial cells with a different value will be highlighted in Red.
'   Lookin in MacroList2 to call this sub.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub QB_Compare_click()       '
    Call OpenCompareWB(ActiveSheet.name)
End Sub

Private Sub OpenCompareWB(Original_sh As String)    '
    Dim wb As Workbook, answer As Long
    Application.Calculation = xlManual ' Turn OFF Calculation
    Application.EnableEvents = False   ' Turn OFF
    ORIGINALSH_GLOBAL = Original_sh ' set it globally for later use in FORMs
    Set OPEN_WB = ActiveWorkbook  ' Set is globally for later use in FORMS
    If Not SheetExists(ORIGINALSH_GLOBAL, ActiveWorkbook) Then ' This Workbook is the Add-In
        answer = MsgBox("Aborting!" & vbCrLf & _
            "Missing Sheet Name '" & ORIGINALSH_GLOBAL & "' from This Workbook." _
        & vbCrLf, vbInformation)
        GoTo abortHere
    End If
    ActiveWorkbook.Sheets(Original_sh).Activate ' Make Add-in Active Workbook
    
    Call ClearFillandComments(ActiveSheet) ' Clear All Comments and and Fill
    
    ' Open and Select second workbook to compare
    Set wb = myOpenWorkbook(True) ' Open File Dialog to COMPARE TO ( TRUE = Read Only)
    If wb Is Nothing Then Exit Sub
    
    Call Compare_worksheets(wb)    'select sheet and COMPARE

abortHere:
    Application.Calculation = xlAutomatic ' Turn ON Calculation
    Application.EnableEvents = True ' Turn Events back on
    ORIGINALSH_GLOBAL = ""
End Sub

Private Sub Compare_worksheets(wb As Workbook)        '
Dim t1 As Date, tsecs As Double, m As String, answer As Long
    t1 = Timer
    ' IF ONLY ONE WORKSHEET, NO NEED TO CALL FORM to SELECT WORKSHEET
    If wb.Sheets.count > 1 Then
        PROCESS_ACTION = myQB_COMPARE_WS
        Call SelectSheetForm_Click ' Show User Form
    Else
        ' Do SAME things here as in USERFORM2 for Process Action 2
        ' Compare selected Sheet to OPEN_WB active sheet using QBLookup
       wb.Activate
       Call QBLookupCompare(OPEN_WB.ActiveSheet.UsedRange, wb.ActiveSheet.UsedRange)
    End If
    
    Call CommentDisplay_On
    'tsecs = Format((Timer - t1), "0.0000")
    'MsgBox "Compare_click() Done...   Time elasped = " & tsecs & " Seconds" & _
    '        vbCrLf & "   Which is " & TimeString(CLng(tsecs), True)
End Sub
Sub QBLookupCompare(r1 As Range, r2 As Range)
    Dim hdr As Range, acct As Range
    Dim last_col As Long, last_row As Long
    Dim val1 As Range, val2 As Range
    On Error Resume Next
    ' Too many cells ??
    'MsgBox ("r1 cells = " & r1.Cells.CountLarge)
    last_row = Last(1, r1)
    last_col = Last(2, r1)
    For Each hdr In Range(r1.Cells(1, 1), r1.Cells(2, last_col))
        If hdr <> "" And Left(hdr, 1) <> "(" Then

            Debug.Print "hdr = " & hdr
            For Each acct In Range(r1.Cells(3, 1), r1.Cells(last_row, 7))
                If acct Is Nothing Then GoTo NEXT_ACCT
                If IsEmpty(acct) Then GoTo NEXT_ACCT
                If acct = "" Then GoTo NEXT_ACCT

                Set val1 = qbLOOKUP(r1, acct, hdr)
                If val1 Is Nothing Then GoTo NEXT_ACCT
                If val1 = "" Then GoTo NEXT_ACCT
                If IsEmpty(val1) Then GoTo NEXT_ACCT
                
                Set val2 = qbLOOKUP(r2, acct, hdr)
                If val2 Is Nothing Then GoTo NEXT_ACCT
                If val2 = "" Then GoTo NEXT_ACCT
                If IsEmpty(val2) Then GoTo NEXT_ACCT
                'Debug.Print "acct = " & acct & "  hdr = " & hdr & "  val2 = " & val2
                
                If val1 <> val2 Then
                    Debug.Print "QBLookupCompare: DIFF acct = '" & acct & "'  hdr = '" & hdr & "'  val1 = '" & val1 & "' <> '" & val2 & "'"
                    'MsgBox "QBLookupCompare: DIFF acct = '" & acct & "'  hdr = '" & hdr & "'  val1 = '" & val1 & "' <> '" & val2 & "'"
                    
                    Call Fill_Red(val2)
                End If
NEXT_ACCT:
            'val1 = Empty
            'Set val1 = Nothing
            'val2 = Empty
            Next
NEXT_HDR:
        End If
    Next
End Sub
Sub QBLookupCompare2(r1 As Range, r2 As Range)
    Dim cell As Range, cell2 As Range, hdr As Variant, acct As Variant
    Dim last_col As Long, last_row As Long
    Dim val1 As Variant, val2 As Variant
    
    MsgBox ("r1 cells = " & r1.Cells.CountLarge)
    last_row = Last(1, r1)
    last_col = Last(2, r1)
    
    ' Strip out Meta characters in any account Names
    Call Meta_Strip(Range(r1.Cells(3, 1), r1.Cells(last_row, 7)))
    Call Meta_Strip(Range(r2.Cells(3, 1), r2.Cells(last_row, 7)))
    
    For Each cell In Range(r1.Cells(1, 1), r1.Cells(2, last_col))
        If cell <> "" And Left(cell, 1) <> "(" Then
            hdr = cell
            Debug.Print "hdr = " & cell
            For Each cell2 In Range(r1.Cells(3, 1), r1.Cells(last_row, 7))
                If IsEmpty(cell2) Then GoTo NEXT_ACCT
                If cell2 = "" Then GoTo NEXT_ACCT
                acct = cell2
                
                val1 = qbLOOKUP(r1, acct, hdr)
                If IsEmpty(val1) Then GoTo NEXT_ACCT
                If val1 = "" Then GoTo NEXT_ACCT
                
                val2 = qbLOOKUP(r2, acct, hdr)
                If IsEmpty(val2) Then GoTo NEXT_ACCT
                If val2 = "" Then GoTo NEXT_ACCT
                Debug.Print "acct = " & cell2 & "  hdr = " & hdr & "  val2 = " & val2
                
                If val1 <> val2 Then
                    Debug.Print "val1 <> val2: acct = " & cell2 & "  hdr = " & hdr & "  val1 = " & val1 & "  val2 = " & val2
                    MsgBox "val1 <> val2: acct = " & cell2 & "  hdr = " & hdr & "  val1 = " & val1 & "  val2 = " & val2
                    
                    'Call Fill_Red(val1)
                End If
NEXT_ACCT:
            Next
NEXT_HDR:
        End If
    Next
End Sub
Private Sub Test_Meta_Strip()
    Dim r As Range
    Dim last_row As Long
    last_row = Last(1, Cells)
    Set r = Range(Cells(3, 1), Cells(last_row, 7))
    Call Meta_Strip(r)
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Removes any Meta Characters in the string values
'
' Use: To make qbLookup work if Meta characters happen to be in a QBheader or QBAccount row name
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Meta_Strip(ByRef r As Range)
    Dim cell As Range, pos As Long, new_str
    For Each cell In r
        new_str = Application.WorksheetFunction.Substitute(cell, "*", "")
        new_str = Application.WorksheetFunction.Substitute(new_str, "?", "")
        new_str = Application.WorksheetFunction.Substitute(new_str, "#", "")
        new_str = Application.WorksheetFunction.Substitute(new_str, "[", "")
        If new_str <> cell Then
            Debug.Print "Meta_Strip: " & cell.Worksheet.name & " Stripped " & cell & " to " & new_str
            cell = new_str
        End If
    Next
End Sub

Private Sub CellbyCellCompare(r1 As Range, r2 As Range)
    Dim cell As Range, cell2 As Range
    MsgBox ("r1 cells = " & r1.Cells.CountLarge)
    For Each cell In r1
        Set cell2 = r2(cell.row, cell.Column)
        If cell.Value <> cell2.Value Then
            Call Fill_Red(cell)
        End If
    Next cell
End Sub
'Private Sub test_VBqbLookup()
'   '''' VBqbLookup DOES NOT WORK!!!!
'    Dim cell As Range, answer As Variant
'    For Each cell In Selection
'         answer = VBqbLOOKUP("Sheet1 (2)", cell, "BBM-Operational")
'        If IsEmpty(answer) Then GoTo NEXT_ONE
'        'If answer Is Nothing Then GoTo NEXT_ONE
'        cell.Offset(0, 3) = answer
'        MsgBox " cell = " & cell & "   Value = " & answer
'NEXT_ONE:
'    Next
'End Sub
Private Sub OLDTest_Meta_Strip()
    Dim cell As Range, pos As Long, new_str, row As Long
    row = 5
    For Each cell In Selection
        new_str = Application.WorksheetFunction.Substitute(cell, "*", "")
        new_str = Application.WorksheetFunction.Substitute(new_str, "?", "")
        new_str = Application.WorksheetFunction.Substitute(new_str, "#", "")
        new_str = Application.WorksheetFunction.Substitute(new_str, "[", "")
        If new_str <> cell Then
            cell = new_str
            
            'ActiveSheet.Range("H" & row) = new_str
            'row = row + 1
        End If
    Next
End Sub
Private Sub Test_Meta_Substitute()
    Dim cell As Range, pos As Long, new_str
    For Each cell In ActiveSheet.UsedRange
        new_str = Application.WorksheetFunction.Substitute(cell, "*", "~*")
        new_str = Application.WorksheetFunction.Substitute(new_str, "?", "~?")
        new_str = Application.WorksheetFunction.Substitute(new_str, "#", "~#")
        new_str = Application.WorksheetFunction.Substitute(new_str, "[", "~[")
        cell.Offset(0, 1) = new_str
    Next
End Sub


Private Sub Test_File_Dialog() ' Test File Dialog  Object
    Dim myObj As FileDialog
    Dim myDirString As String
    Set myObj = Application.FileDialog(msoFileDialogFilePicker)
    With myObj
        .InitialFileName = "C:\Users\" & Environ$("Username") & "\Desktop"
        .InitialFileName = "N:\Dwight\QUICKBOOKS WORK AREA\" ' make sure to add the last "\"
        .Filters.Clear
        .Filters.Add "Custom Excel Files", "*.xlsx, *.csv, *.xls"
        .FilterIndex = 1
        .Show
    End With
    Dim yourObj3 As FileDialog
    Dim yourDirString3 As String
    Set yourObj3 = Application.FileDialog(msoFileDialogFilePicker)
    With yourObj3
        .InitialFileName = "N:\Dwight\QUICKBOOKS WORK AREA\"  ' make sure to add the last "\"
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        .FilterIndex = 1
        .Show
    End With
End Sub


