Attribute VB_Name = "Import_IIF_OLD_WAY"
Option Explicit

Sub Import_COA_Click()  ' Run macro from a shortcut
    Call Import_IIF
End Sub
Private Sub Import_IIF()   '
    Dim origWB As Workbook, wb As Workbook, answer As Long
    
    Application.Calculation = xlManual ' Turn OFF Calculation
    Application.EnableEvents = False   ' Turn OFF
    
    Set origWB = ActiveWorkbook
    Set wb = myOpenWorkbook(True, "*.IIF", origWB)  ' Open File Dialog to copy from ( TRUE for Read Only)
    If wb Is Nothing Then Exit Sub
    
    ' *** Set global Variables for the User Form ***
    PROCESS_ACTION = myADD_IIF
    
    ' *** Call User Form ***
    Call SelectSheetForm_Click   'Show User Form to add and format IIF worksheet
abortHere:
    Application.Calculation = xlAutomatic ' Turn ON Calculation
    Application.EnableEvents = True ' Turn Events back on
End Sub

' Function: FormatSheet_COA
    ' Determine new sheet name from Entity in the Class with REFNUM = 1
    ' Rename existing sheet, if any with "OLD " in front
    ' Copy and Paste Row 1 Columns O-V (Header Row)
    ' Copy and Paste Row 2 Columns O-V (Data Rows)
    ' Format Columns
    ' Format Rows 2 and on
    ' Copy ClassTypes and my Comment Columns for esisting Accounts
    ' Change Tab Color
    ' Copy Print Format Seetings
    ' Freeze View of Row 1 (header Row)


Function FormatSheet_COA(ByVal copy_SH As String) As String ' Called from User Form
Dim last_row As Long, i As Long, cell As Range
Dim answer As Long, m As String, newShName As String, template_SH As String
Dim delete_template As Boolean
     If Not SheetExists(copy_SH) Then
        m = "FormatSheet_COA() Error: Sheet ' & copy_SH & " ' does not exist... Aborting."
        Debug.Print m
        MsgBox m
     End If
    
    ' Turn off Events
    Application.Calculation = xlManual ' Turn OFF Calculation
    Application.EnableEvents = False   ' Turn OFF
    Application.ScreenUpdating = False ' Turn OFf
    'Sheets(copy_SH).Select ' Leave this here
    'On Error Resume Next
   
    last_row = Last(1, Sheets(copy_SH).Cells)
    
    ' Determine new sheet name from Entity in the Class with REFNUM = 1
    For i = last_row To 1 Step -1  ' Step backwards from bottom
        Set cell = Cells(i, 1)
        If cell = "CLASS" And cell.Offset(0, 2) = 1 Then
            newShName = cell.Offset(0, 1) & " COA"
        End If
    Next i
    
    If newShName = "" Or InStr(1, newShName, "?") > 0 Then
        m = "FormatSheet_COA() Error: newShName = " & newShName & vbCrLf
        m = m & " Could not find a valid Class to identify the Entity. Aborting."
        Debug.Print m
        MsgBox m
        GoTo abortHere
    End If
    
    ' Remove Unnecessary Rows
    For i = last_row To 1 Step -1  ' Step backwards from bottom
        Set cell = Cells(i, 1)
        If cell <> "ACCNT" And cell <> "!ACCNT" Then
            Rows(cell.row & ":" & cell.row).Delete ' Delete Row
        End If
    Next i
    ' Check ACCT rows exist
    If Last(1, Sheets(copy_SH).Cells) = 0 Then
        m = "Aborting!" & vbCrLf
        m = m & " There were no ACCT rows in IIF worksheet " & copy_SH
        answer = MsgBox(m, vbInformation)
        GoTo abortHere
    End If
    
   
    ' Rename existing sheet, if any with "OLD " in front
    If SheetExists(newShName) Then
        m = "Renaming '" & newShName & "' to 'OLD " & newShName & "' worksheet in " & ActiveWorkbook.name
        answer = MsgBox(m, vbInformation)
        If SheetExists("OLD " & newShName) Then
            ' Remove old sheet or template first
            Application.DisplayAlerts = False
            Sheets("OLD " & newShName).Delete
            Application.DisplayAlerts = True
        End If
        Sheets(newShName).name = "OLD " & newShName
        template_SH = "OLD " & newShName
    Else
        template_SH = "COA"
        m = "FormatSheet_COA(): sheet name = '" & newShName & "' does not exist." & vbCrLf
        m = m & "  COPYING '" & template_SH & "' as a template from workbook " & ThisWorkbook.name & "..."
        m = m & vbCrLf & vbCrLf & "Press OK or Cancel to Abort..."
        Debug.Print m
        answer = MsgBox(m, vbInformation + vbOKCancel, " FormatSheet_COA() ")
        
        template_SH = CopyWorksheets(ThisWorkbook, ActiveWorkbook, template_SH, Sheets(1).name, _
                        False, False) 'False, False so that myQbaddin to MyQBAddin can add the  ws.
        If template_SH = "" Then
            m = "CopyWorksheets(ThisWorkbook, ActiveWorkbook, template_SH) failed."
            Debug.Print m
            MsgBox m
            GoTo abortHere
        End If
        
        Sheets(template_SH).name = "OLD " & newShName
        template_SH = "OLD " & newShName
        
        'delete_template = True ' set to True to delete ws at bottom of this sub.
    End If
   
    
' Copy and Paste Row 1 Columns O-V (Header Row)
    Sheets(template_SH).Range("O1:V1").Copy Sheets(copy_SH).Range("O1:V1")
' Copy and Paste Row 2 Columns O-V (Data Rows)
    last_row = Last(1, Sheets(copy_SH).Cells)
    Sheets(template_SH).Range("O2:V2").Copy Sheets(copy_SH).Range("O2:V" & last_row)
' Format Columns
    Sheets(template_SH).Select
    Columns("A:V").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets(copy_SH).Select
    Sheets(copy_SH).Range("A:V").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
' Format Rows 2 and on
    last_row = Last(1, Sheets(copy_SH).Cells)
    Sheets(template_SH).Range("A2:V2").Copy
    Sheets(copy_SH).Range("A2:V" & last_row).Select ' Format all Rows 2 and on
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
' Copy ClassTypes and my Comment Columns for esisting Accounts
    Call Copy_ClassTypesToActiveSheet
' Change Tab Color
    Call changeTabColor(copy_SH, xlThemeColorLight2) ' Change tab to Navy Blue
' Copy Print Format Seetings
    Call copyPrintSettings(template_SH, copy_SH)
    Call copyPrintTitleRows(template_SH, copy_SH)
' Freeze View of Row 1 (header Row)
    Call FreezeTopRow(copy_SH)
    
    'MsgBox "About to rename Sheet... " & copy_SH
    
    Sheets(copy_SH).name = newShName
    If delete_template Then
            Application.DisplayAlerts = False
            Sheets(template_SH).Delete
            Application.DisplayAlerts = True
    End If
    
abortHere:
    Application.ScreenUpdating = True ' Turn Back on
    Application.Calculation = xlAutomatic ' Turn ON Calculation
    Application.EnableEvents = True ' Turn Events back on
    FormatSheet_COA = newShName
End Function

' Need to clean up globals
' Copy comments and stuff for my additional columns
Private Sub Copy_ClassTypesToActiveSheet()
    Dim t1 As Date
    Dim CtHdrCell As Range, AccHdrCell, ACCRange As Range, searchRange As Range
    Dim last_col As Long, RangeStr As String, last_row As Long, cell As Range
    Dim SavedCalcMode As Long, SavedEnableEvents As Boolean
    
    
    If ORIGINALSH_GLOBAL = "" Then Exit Sub ' Need this Global variable set
    If ORIGINALSH_GLOBAL = ActiveSheet.name Then Exit Sub ' Nothing to do; can't copy to itself.
    
    On Error Resume Next
    t1 = Timer
    SavedCalcMode = Application.Calculation ' Save Mode to restore later
    SavedEnableEvents = Application.EnableEvents ' Save Mode to restore later
    Application.Calculation = xlManual ' Turn OFF Calculation
    Application.EnableEvents = False   ' Turn OFF
    
    Set CtHdrCell = ActiveSheet.Range("1:1").Find("Comment") ' find Comment column (the last column in range)
    Set AccHdrCell = ActiveSheet.Range("1:1").Find("ACCNUM") ' find ACCNUM column
    If IsEmpty(AccHdrCell) Then
        Debug.Print ("  Copy_ClassTypes(): Couldn't find column 'ACCNUM' in Chart of Accounts header row")
        Exit Sub
    End If
    If IsEmpty(CtHdrCell) Then
        Debug.Print ("  Copy_ClassTypes(): Couldn't find 'Comment' Header in Chart of Accounts header row")
        Exit Sub
    End If
    last_row = Last(1, Range(Chr(64 + AccHdrCell.Column) & ":" & Chr(64 + AccHdrCell.Column))) ' First Column
    
    ' Change Column Numbers to Column Letters
    Set searchRange = ActiveWorkbook.Sheets(ORIGINALSH_GLOBAL).Range(Chr(64 + AccHdrCell.Column) & 1 & ":" & Chr(64 + CtHdrCell.Column) & last_row)
    Set ACCRange = ActiveSheet.Range(Chr(64 + AccHdrCell.Column) & 2 & ":" & Chr(64 + AccHdrCell.Column) & last_row)
    
    For Each cell In ACCRange
' Tested both LookUp_ByIndex and LookUp_ClassType. The Time differences were Imperceptible
'        cell.Offset(0, 10) = LookUp_ByIndex(cell.Value, SearchRange, 11) ' 10 = Col R
'        cell.Offset(0, 11) = LookUp_ByIndex(cell.Value, SearchRange, 12) ' 11 = Col S
'        cell.Offset(0, 12) = LookUp_ByIndex(cell.Value, SearchRange, 13) ' 12 = Col T
'        cell.Offset(0, 13) = LookUp_ByIndex(cell.Value, SearchRange, 14) ' 13 = Col U
        cell.Offset(0, 10) = LookUp_ClassType(cell.Value, searchRange, "BBM-OP")  ' 10 = Col R
        cell.Offset(0, 11) = LookUp_ClassType(cell.Value, searchRange, "8825-R")  ' 11 = Col S
        cell.Offset(0, 12) = LookUp_ClassType(cell.Value, searchRange, "DY")      ' 12 = Col T
        cell.Offset(0, 13) = LookUp_ClassType(cell.Value, searchRange, "Comment") ' 13 = Col U
    Next
    
    Application.Calculation = SavedCalcMode ' Restore to original State
    Application.EnableEvents = SavedEnableEvents       ' Restore to original State
    'MsgBox "Copy_ClassTypes() Done...  Time elasped = " & Format((Timer - t1), "0.0000") & " Seconds"
End Sub


Sub FormatSheet_Class_List(copy_SH As String) ' Called from User Form
Dim last_row As Long, i As Long, cell As Range, m As String
Dim answer As Long, newShName As String, template_SH  As String, delete_template As Boolean
     If Not SheetExists(copy_SH) Then
        m = "FormatSheet_Class_List() Error: Sheet ' & copy_SH & " ' does not exist... Aborting."
        Debug.Print m
        MsgBox m
     End If
    
    Application.Calculation = xlManual ' Turn OFF Calculation
    Application.EnableEvents = False   ' Turn OFF
    Application.ScreenUpdating = False ' Turn OFf
    
    last_row = Last(1, Sheets(copy_SH).Cells)
    
    ' Determine new sheet name from Entity in the Class with REFNUM = 1
    For i = last_row To 1 Step -1  ' Step backwards from bottom
        Set cell = Cells(i, 1)
        If cell = "CLASS" And cell.Offset(0, 2) = 1 Then
            newShName = cell.Offset(0, 1) ' BBM, MGH, YHC, etc
        End If
    Next i
    
    
' Remove Unnecessary Rows
    For i = last_row To 1 Step -1  ' Step backwards from bottom
        Set cell = Cells(i, 1)
        If cell <> "CLASS" And cell <> "!CLASS" Then
            Rows(cell.row & ":" & cell.row).Delete ' Delete Row
        End If
    Next
    If Last(1, Sheets(copy_SH).Cells) = 0 Then
        m = "Aborting!" & vbCrLf
        m = m & " There were no CLASS rows in IIF worksheet " & copy_SH
        answer = MsgBox(m, vbInformation)
        GoTo abortHere
    End If
   
    ' Rename existing sheet, if any with "OLD " in front
    If SheetExists(newShName) Then
        m = "Renaming '" & newShName & "' to 'OLD " & newShName & "' worksheet in " & ActiveWorkbook.name
        answer = MsgBox(m, vbInformation)
        If SheetExists("OLD " & newShName) Then
            ' Remove old sheet or template first
            Application.DisplayAlerts = False
            Sheets("OLD " & newShName).Delete
            Application.DisplayAlerts = True
        End If
        Sheets(newShName).name = "OLD " & newShName
        template_SH = "OLD " & newShName
    Else
        template_SH = "CL"
        m = "FormatSheet_Class_List(): sheet name = '" & newShName & "' does not exist." & vbCrLf
        m = m & "  COPYING '" & template_SH & "' as a template from workbook " & ThisWorkbook.name & "..."
        Debug.Print m
        answer = MsgBox(m, vbInformation)
        
        template_SH = CopyWorksheets(ThisWorkbook, ActiveWorkbook, template_SH, Sheets(1).name, _
                        False, False) 'False, False so that myQbaddin to MyQBAddin can add the  ws.
        If template_SH = "" Then
            m = "CopyWorksheets(ThisWorkbook, ActiveWorkbook, template_SH) failed."
            Debug.Print m
            MsgBox m
            GoTo abortHere
        End If
        
        Sheets(template_SH).name = "OLD " & newShName
        template_SH = "OLD " & newShName
        
        'delete_template = True ' set to True to delete ws at bottom of this sub.
    End If
   
' Copy and Paste Row 1 Columns G-I (Header Row)
    Sheets(template_SH).Range("G1:I1").Copy Sheets(copy_SH).Range("G1:I1")
' Copy and Paste Row 2 Columns G-I (Data Rows)
    last_row = Last(1, Sheets(copy_SH).Cells)
    Sheets(template_SH).Range("G2:I2").Copy Sheets(copy_SH).Range("G2:I" & last_row)
' Format Columns
    Sheets(template_SH).Columns("A:I").Copy
    Sheets(copy_SH).Activate
    Sheets(copy_SH).Columns("A:I").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
' Format Rows 2 and on
    last_row = Last(1, Sheets(copy_SH).Cells)
    Sheets(template_SH).Range("A2:I2").Copy
    Sheets(copy_SH).Range("A2:I" & last_row).Select ' Format all Rows 2 and on
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    ' Copy ClassTypes and Comments for Existing Accounts
    '    Call Copy_MyClassTypesToActiveClassSheet ' Not used now
    
' Change Tab Color to Black
    Call changeTabColor(copy_SH, xlThemeColorLight2) ' Change tab to Navy Blue
' Copy Print Format Seetings
    Call copyPrintSettings(template_SH, copy_SH)
    Call copyPrintTitleRows(template_SH, copy_SH)
' Freeze View of Row 1 (header Row)
    Call FreezeTopRow(copy_SH)
    'MsgBox "About to rename Sheet... " & copy_SH
    If SheetExists("Old Class List", OPEN_WB) Then
        Application.DisplayAlerts = False
        OPEN_WB.Sheets("Old Class List").Delete
        Application.DisplayAlerts = True
    End If
    
    Sheets(copy_SH).name = newShName
    If delete_template Then
            Application.DisplayAlerts = False
            Sheets(template_SH).Delete
            Application.DisplayAlerts = True
    End If

abortHere:
    Application.ScreenUpdating = True ' Turn Back on
    Application.Calculation = xlAutomatic ' Turn ON Calculation
    Application.EnableEvents = True ' Turn Events back on
End Sub


