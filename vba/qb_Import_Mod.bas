Attribute VB_Name = "qb_Import_Mod"
Option Explicit

Sub Import_Sheets_Sub(wName As String)
Dim wb As Workbook ' For opening a new workbook
Dim bsName As String, shName As String
Dim m As String, ans As Long
    If wName <> ActiveWorkbook.name Then
        m = "NOT the Active Worbook!   Name = " & wName & vbCrLf & vbCrLf
        m = m & "Do you want to Continue?"
        ans = MsgBox(m, vbYesNoCancel, "WARNING!.. WORKBOOK NOT ACTIVE! ")
        If ans <> vbYes Then Exit Sub
    End If

    Application.Calculation = xlManual ' Turn OFF Calculation
    Application.EnableEvents = False
    shName = ActiveSheet.name ' Save Active sheet name for  Activating later
    Set wb = myQBAddin.myOpenWorkbook(True, "*.xls*") ' Open File Dialog to copy from ( TRUE for Read Only)
    If wb Is Nothing Then GoTo ErrorHandler
    
    ' IF ONLY ONE WORKSHEET, NO NEED TO CALL FORM to SELECT WORKSHEET
    If wb.Sheets.count > 1 Then
        PROCESS_ACTION = myQB_ADD_WS  'MyQBAddin
        Call SelectSheetForm_Click   'Show User Form to add worksheets
    Else
' Do EXACT same actions as Process_Selection() in USER_FORM SHEETNAMES2

' NOTES ON CopyWorksheetCloseWB:
' CopyWorksheetCloseWB(wB As Workbook, sheetName As String, beforeafterSheet As String,
'              Optional afterflag As Boolean = False, Optional replaceflag As Boolean = True)
' OPEN_WB as Copy To wb;  wB  = Open workbook to copy From and Close
' sheetName     = Name of the sheet to copy
' beforeafterSheet = Place sheetName before or after this sheet in wB
' afterflag:    False = Place before the beforeafterSheet in wB
' replaceflag:  True = Replace if sheetname already exists in wB
' Return the FINAL name of the new sheet in wB
        bsName = CopyWorksheetCloseWB(wb, wb.Sheets(1).name, shName, False, False)
        ActiveWorkbook.Worksheets(shName).Activate  ' Activate the Sheet Again
       
       ' ReBuild the sheet name table
       Call Build_SheetNameTable_Sub(ActiveWorkbook.name)
       Range("Sheet_Name").Value = bsName ' Set to the sheet name imported
    End If
ErrorHandler:
    Application.Calculation = xlAutomatic ' Turn ON Calculation
    Application.EnableEvents = True
End Sub
' This was IMPROVED from original in myQBAddin
Sub Build_SheetNameTable_Sub(wName As String)
Dim sT As Range, ws As Worksheet
Dim r As Long, qbper_end As String
Dim str As String, yr As Long, mo As Long, dy As Long
Dim m As String, ans As Long
    'If wName <> ActiveWorkbook.name Then
    '    M = "NOT the Active Worbook!   Name = " & wName & vbCrLf & vbCrLf
    '    M = M & "Do you want to ACTIVATE and Continue?"
    '    ans = MsgBox(M, vbYesNoCancel, "WARNING!.. WORKBOOK NOT ACTIVE! ")
    '    If ans <> vbYes Then Exit Sub
    '    Workbooks(wName).Activate
    'End If
    Application.Calculation = xlManual ' Turn OFF Calculation
    Application.EnableEvents = False
    On Error GoTo ErrorHandler
    Set sT = Range("SheetNameTable")
    If sT Is Nothing Then
        MsgBox "Range Name 'SheetNameTable' not found...  Exiting"
        GoTo ErrorHandler
    End If
    sT.ClearContents
    sT.Cells(1, 1) = "Sheet Name"
    sT.Cells(1, 2) = "Period Ending"
    sT.Cells(1, 3) = "Created"
    r = 1
    For Each ws In Worksheets
        qbper_end = ""
        ' If Tab color is anything but default, skip
        If ws.Tab.Color <> False Then GoTo adj_Ws
        
        'MsgBox "ws " & ws.Name & "  Tab Color = " & ws.Tab.Color & "  Tab Theme = " & ws.Tab.ThemeColor
    
        On Error Resume Next
        str = ""
        ' QB Period Ending is defined as Defined Name "QBENDDATE" in each QB worksheet
        If Not IsEmpty(ws) Then
            str = ws.Names("QBENDDATE").Value
            If str = "" Then
                m = "Skipping worksheet '" & ws.name & "'!  Set a tab color on sheet to skip this message..." & vbCrLf
                m = m & vbCrLf & "  QBENDDATE is NOT defined."
                ans = MsgBox(m, vbInformation)
                GoTo adj_Ws
            End If
            yr = Mid(str, 2, 4) ' first char is '='
            mo = Mid(str, 6, 2)
            dy = Mid(str, 8, 2)
            'MsgBox "yr = " & yr & vbCrLf & _
            '        "mo = " & mo & vbCrLf & _
            '        "dy = " & dy & vbCrLf
            qbper_end = DateValue(mo & "/" & dy & "/" & yr) ' very important line... format date properly
            If qbper_end <> "" Then
                sT.Cells(r + 1, 1) = ws.name
                sT.Cells(r + 1, 2) = qbper_end
                sT.Cells(r + 1, 3) = Get_QBHdr_DateTime(From_Print_Header(ws.PageSetup.LeftHeader))
                r = r + 1  ' Increment the row
            End If
        End If
        On Error GoTo ErrorHandler
adj_Ws:
    Next ws
ErrorHandler:
    ' Add this in where the table linked to the pivot is updated
    ActiveWorkbook.RefreshAll

    Application.Calculation = xlAutomatic ' Turn ON Calculation
    Application.EnableEvents = True
End Sub

Sub Go_To_Sub(wName As String)
Dim sheetnmrng As Range, row As Range, col As Range
Dim pos As Long, answer As Variant, i As Long, result() As String  ' Array
Dim m As String, ans As Long
    If wName <> ActiveWorkbook.name Then
        m = "NOT the Active Worbook!   Name = " & wName & vbCrLf & vbCrLf
        m = m & "Do you want to Continue?"
        ans = MsgBox(m, vbYesNoCancel, "WARNING!.. WORKBOOK NOT ACTIVE! ")
        If ans <> vbYes Then Exit Sub
    End If
'MsgBox Selection.Formula(1)
    pos = InStr(UCase(Selection(1).Formula), "QBLOOKUP(")
    If pos = 0 Then
        MsgBox ("Can't find 'qbLOOKUP' in the Selected Cells")
        Exit Sub
    End If
    m = Mid(Selection(1).Formula, pos + 9) ' start past the "qbLOOKUP("
    result = Split(m, ",") ' Split string with ',' as seperator
    Set sheetnmrng = Cells.Range(result(0))
    Set row = Cells.Range(result(1))
    Set col = Cells.Range(result(2))
    Set answer = UDFs_QB_WB.qbLOOKUP(sheetnmrng, row, col)

    If answer Is Nothing Then
        MsgBox "answer is Nothing"
        Exit Sub
    End If
    If IsEmpty(answer) Then
        MsgBox "answer is Empty"
        Exit Sub
    End If
    'MsgBox "answer = " & answer & " Address = " & answer.Address
    Application.GoTo Reference:=Worksheets(sheetnmrng.Value).Range(answer.Address) ' , Scroll:=True  ' Don'
    
    ' Clear any Filled cells on worksheet first
    With ActiveSheet.Cells.Interior
        .pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ' the next line doesn't seem to do anything.
    answer.Activate
    ' fill answer cell with yellow
    Call Fill_Yellow(answer)
End Sub
Private Sub test_Get_QBHdr_DateTime()
    MsgBox Get_QBHdr_DateTime(Selection)
End Sub
' This function gets the date and time out of a Name in a QB Bal Sheet
Function Get_QBHdr_DateTime(str_In As String) As Variant
Dim s As String
    s = Trim(Left(str_In, InStr(1, str_In, " Cash"))) ' strip "Cash Basis" off end
    If s = "" Then Get_QBHdr_DateTime = "No Print Header": Exit Function
    If Mid(s, 2, 1) = ":" Then s = "0" & s ' make sure Hour is 2 digit format
    Get_QBHdr_DateTime = Str_TO_DATE(s, "hh:mm pm MM/dd/yy")
End Function
Private Sub test_Str_TO_DATE()
   MsgBox Str_TO_DATE("01:13 pm 11/11/24", "hh:mm pm MM/dd/yy")
End Sub
' This function gets the date and time out of a formatted string. Very cool!
' Used by Get_QBHdr_DateTime to read QB Bal Sheet Header
'https://stackoverflow.com/questions/53021679/convert-string-to-datetime-in-excel-with-individual-format-parameter-like-to-da
Function Str_TO_DATE(ByRef src As String, ByRef frmt As String) As Variant
Dim y As Long, m As Long, d As Long, h As Long, min As Long, s As Long
Dim am As Boolean, pm As Boolean
Dim pos As Long
    If src = "" Then Exit Function
    If Len(src) <> Len(frmt) Then
    ' Len of src must match Len of format
        Str_TO_DATE = CVErr(xlErrNA)  ' #N/A error = 2042
        Exit Function
    End If

    pos = InStr(1, frmt, "yyyy", vbTextCompare)
    If pos > 0 Then
        y = Val(Mid(src, pos, 4))
    Else: pos = InStr(1, frmt, "yy", vbTextCompare)
        If pos > 0 Then
            y = Val(Mid(src, pos, 2))
            If y < 80 Then y = y + 2000 Else y = y + 1900
        End If
    End If

    pos = InStr(1, frmt, "mmm", vbTextCompare)
    If pos > 0 Then
        m = Month(DateValue("01 " & (Mid(src, pos, 3)) & " 2000"))
    Else: pos = InStr(1, frmt, "MM", vbBinaryCompare)
        If pos > 0 Then m = Val(Mid(src, pos, 2))
    End If

    pos = InStr(1, frmt, "dd", vbTextCompare)
    If pos > 0 Then d = Val(Mid(src, pos, 2))

    pos = InStr(1, frmt, "hh", vbTextCompare)
    If pos > 0 Then h = Val(Mid(src, pos, 2))
    If InStr(1, src, "am", vbTextCompare) > 0 Then am = True
    If InStr(1, src, "AM", vbTextCompare) > 0 Then am = True
    If InStr(1, src, "a.m.", vbTextCompare) > 0 Then am = True
    If InStr(1, src, "a. m.", vbTextCompare) > 0 Then am = True
    If InStr(1, src, "pm", vbTextCompare) > 0 Then pm = True
    If InStr(1, src, "p.m.", vbTextCompare) > 0 Then pm = True
    If InStr(1, src, "p. m.", vbTextCompare) > 0 Then pm = True
    If am And h = 12 Then h = 0
    If pm And h <> 12 Then h = h + 12

    pos = InStr(1, frmt, "mm", vbBinaryCompare)
    If pos > 0 Then min = Val(Mid(src, pos, 2))

    pos = InStr(1, frmt, "ss", vbTextCompare)
    If pos > 0 Then s = Val(Mid(src, pos, 2))

    Str_TO_DATE = DateSerial(y, m, d) + TimeSerial(h, min, s)
End Function

Private Sub test_Created_Date()
    MsgBox From_Print_Header(Selection)
End Sub
Function From_Print_Header(arg1) As String
Dim i As Long, char As String, formating_flag As Boolean
Dim ans As String
On Error Resume Next
' Remove print formating
    For i = 1 To Len(arg1)
        char = Mid(arg1, i, 1)
        If char = "&" Then formating_flag = Not formating_flag
        If formating_flag Then GoTo next_char
        If char = vbCrLf Then GoTo next_char
        If char = vbCr Then GoTo next_char
        If char = vbLf Then GoTo next_char
        If char = vbNewLine Then GoTo next_char
        If char = "&" Then
            i = i + 1 ' skip next char after &
            GoTo next_char
        End If
        ans = ans & char
next_char:
    Next
    From_Print_Header = ans
End Function
' Hide Buttons for Printing
Sub HideButtons()
    On Error Resume Next
    ActiveSheet.Build_SheetNameTable.Visible = False
    On Error GoTo 0
End Sub
' restore view Buttons from  Printing
Sub ViewButtons()
    On Error Resume Next
    ActiveSheet.Build_SheetNameTable.Visible = True
    On Error GoTo 0
End Sub


