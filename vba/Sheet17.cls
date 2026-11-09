VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Calls are in qb_Import_Mod
Private Sub Go_To_Click()
    Go_To_Sub ThisWorkbook.name
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


Private Sub Go_T()
Dim sheetnmrng As Range, row As Range, col As Range
Dim pos As Long, ans As Long, i As Long, result() As String  ' Array
Dim s As String, answer As Variant
'MsgBox Selection.Formula(1)
    pos = InStr(UCase(Selection(1).Formula), "QBLOOKUP(")
    If pos = 0 Then
        MsgBox ("Can't find 'qbLOOKUP' in the Selected Cells")
        Exit Sub
    End If
    s = Mid(Selection(1).Formula, pos + 9) ' start past the "qbLOOKUP("
    result = Split(s, ",") ' Split string with ',' as seperator
    Set sheetnmrng = Cells.Range(result(0))
    Set row = Cells.Range(result(1))
    Set col = Cells.Range(result(2))
    Set answer = qbLOOKUP(sheetnmrng, row, col)

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

Private Sub Import_S()
Dim wb As Workbook, m As String, ans As Long
Dim bsName As String

    Application.Calculation = xlManual ' Turn OFF Calculation
    Application.EnableEvents = False
    Set wb = myOpenWorkbook(True, "*.xls*") ' Open File Dialog to copy from ( TRUE for Read Only)
    If wb Is Nothing Then GoTo ErrorHandler
    
    ' IF ONLY ONE WORKSHEET, NO NEED TO CALL FORM to SELECT WORKSHEET
    If wb.Sheets.count > 1 Then
        PROCESS_ACTION = myQB_ADD_WS

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
        bsName = CopyWorksheetCloseWB(wb, wb.Sheets(1).name, "QB DUE To From", True, False)
        ActiveWorkbook.Worksheets("QB DUE To From").Activate  ' Activate the Sheet
    End If
    Call Build_SheetNameTable_Click
ErrorHandler:
    Application.Calculation = xlAutomatic ' Turn ON Calculation
    Application.EnableEvents = True
End Sub
Sub Build_SheetNameTab()
    Dim sT As Range, ws As Worksheet
    Dim r As Long, qbper_end As String, ans As String
    Dim str As String, yr As Long, mo As Long, dy As Long
    Application.Calculation = xlManual ' Turn OFF Calculation
    Application.EnableEvents = False
    On Error GoTo ErrorHandler
    Set sT = Range("SheetNameTable")
    
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
        ' QB Period Ending is defined as Defined Name "QBENDDATE" in each QB worksheet
        If Not IsEmpty(ws) Then
            str = ws.Names("QBENDDATE").Value
            If str = "" Then
                ans = MsgBox("Warning!  QBENDDATE is NOT defined in worksheet " & ws.name & ".", vbInformation)
                GoTo adj_Ws
            End If
            yr = Mid(str, 2, 4) ' first char is '='
            mo = Mid(str, 6, 2)
            dy = Mid(str, 8, 2)
            'MsgBox "yr = " & yr & vbCrLf & _
            '        "mo = " & mo & vbCrLf & _
            '        "dy = " & dy & vbCrLf
            qbper_end = DateValue(mo & "/" & dy & "/" & yr) ' very important line... format date properly
        End If
        If qbper_end <> "" Then
            sT.Cells(r + 1, 1) = ws.name
            sT.Cells(r + 1, 2) = qbper_end
            sT.Cells(r + 1, 3) = Get_QBHdr_DateTime(From_Print_Header(ws.PageSetup.LeftHeader))
            r = r + 1  ' Increment the row
        End If
        On Error GoTo ErrorHandler
adj_Ws:
    Next ws
ErrorHandler:
    Application.Calculation = xlAutomatic ' Turn ON Calculation
    Application.EnableEvents = True
End Sub
Private Sub test_Get_QBHdr_DateTime()
    MsgBox Get_QBHdr_DateTime(Selection)
End Sub
Private Function Get_QBHdr_DateTime(str_In As String) As Variant
Dim s As String
    s = Trim(Left(str_In, InStr(1, str_In, " Cash"))) ' strip "Cash Basis" off end
    If Mid(s, 2, 1) = ":" Then s = "0" & s ' make sure Hour is 2 digit format
    Get_QBHdr_DateTime = TO_DATE(s, "hh:mm pm MM/dd/yy")
End Function
'https://stackoverflow.com/questions/53021679/convert-string-to-datetime-in-excel-with-individual-format-parameter-like-to-da
Function TO_DATE(ByRef src As String, ByRef frmt As String) As Variant
Dim y As Long, m As Long, d As Long, h As Long, min As Long, s As Long
Dim am As Boolean, pm As Boolean
Dim pos As Long

    If Len(src) <> Len(frmt) Then
        TO_DATE = CVErr(xlErrNA)  ' #N/A error
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

    TO_DATE = DateSerial(y, m, d) + TimeSerial(h, min, s)
End Function

Private Sub test_Created_Date()
    MsgBox From_Print_Header(Selection)
End Sub
Private Function From_Print_Header(arg1) As String
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
Sub Calc_Due_to_From()
    Dim sht1 As String, sht2 As String, var As Variant, m As String
    Dim t1 As Date, per_end As Date
    'On Error Resume Next
    t1 = Timer
        Range("Comment") = "Calculating..."
        Application.Calculation = xlManual ' Turn OFF Calculation
        Application.EnableEvents = False   ' Turn OFF events
        Err.Clear
        sht1 = Range("Sheet_Name").Value
        If sht1 = "" Then
           Debug.Print "Cell Changed... Sheet Name is nothing ''" & vbCrLf
           GoTo Finish_Up
        End If
        var = Application.VLookup(sht1, Range("SheetNameTable"), 2, 0)
        If IsError(var) Then
           Debug.Print "Cell Changed... Could not find Period Ending for Sheet Name '" & sht1 & "'" & vbCrLf
           GoTo Finish_Up
        End If
        On Error Resume Next
        per_end = var
        Debug.Print "Cell Changed...  sht1 ='" & sht1 & "' " & _
        " per_end ='" & per_end & "'"
        
        Range("Period_Ending") = per_end ' new Sheet_Name
Finish_Up:
        ' Turn ON Sheet Calculation
        Sheets("QB DUE To From").EnableCalculation = True
        Application.Calculate ' Calculate Once!
        ' Turn OFF Sheet Calculation
        Sheets("QB DUE To From").EnableCalculation = False
        
        Application.Calculation = xlAutomatic ' Turn ON Calculation
        Range("Comment") = Format((Timer - t1), "0.0000") & " Seconds"
        Application.EnableEvents = True ' Turn Events back on
        On Error GoTo 0
End Sub
Private Sub Switch_ThisSheetName(SheetName1 As String, SheetName2 As String, Optional Not_Found As Long = 0)
    Dim str1 As String, str2 As String
    ' Parameter checks
    If SheetName1 = "" Then Exit Sub
    If SheetName2 = "" Then Exit Sub
    If Not_Found = 1 Then ' replace no matter what SheetName1 is
        str1 = "qbLOOKUP(*!"
    Else
        str1 = "qbLOOKUP(" & SheetName1 & "!"
    End If
    str2 = "qbLOOKUP('" & SheetName2 & "'!" ' always add the apostrophe "'" for saftey sake.

    If Not SheetExists(SheetName2) Then
       Debug.Print "Switch_ThisSheetName(): Sheet2 Does NOT exist. Sheetname = '" & SheetName2 & "'"
       If Not_Found = -1 Then MsgBox "Switch_ThisSheetName(): Sheet2 Does NOT exist. Sheetname = '" & SheetName2 & "'"
       Exit Sub
    End If
    
    On Error Resume Next
    Cells.Replace what:=str1, Replacement:=str2, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:= _
    False, SearchFormat:=False, ReplaceFormat:=False
    
    On Error GoTo 0
End Sub
Private Function SheetExists(strSheetName As String, Optional wbWorkbook As Workbook) As Boolean
    If wbWorkbook Is Nothing Then Set wbWorkbook = ActiveWorkbook           'or ThisWorkbook - whichever appropriate
    Dim obj As Object
    If hasQuotes(strSheetName) = True Then  'Sheet Names sometimes have single quotes
        strSheetName = stripQuotes(strSheetName) ' Strip single quotes from name
    End If

    On Error GoTo HandleError
    Set obj = wbWorkbook.Sheets(strSheetName)
    SheetExists = True
    On Error GoTo 0
    Exit Function
HandleError:
    On Error GoTo 0
    SheetExists = False
End Function
Private Function hasQuotes(sname As String) As Boolean ' see if sname enclosed in quotes
    hasQuotes = False
    If Left(sname, 1) = "'" Then hasQuotes = True 'some Sheet Names have single quotes
End Function
Private Function stripQuotes(sname As String) As String '
    stripQuotes = Mid(sname, 2, Len(sname) - 2)
    ' Some Sheet Names are single quotes, Remove first and last chars
End Function
Private Sub HideButtons()
    On Error Resume Next
    ActiveSheet.Build_SheetNameTable.Visible = False
    On Error GoTo 0
End Sub
Private Sub ViewButtons()
    On Error Resume Next
    ActiveSheet.Build_SheetNameTable.Visible = True
    On Error GoTo 0
End Sub
Private Sub Worksheet_FollowHyperlink(ByVal target As Hyperlink)

End Sub
