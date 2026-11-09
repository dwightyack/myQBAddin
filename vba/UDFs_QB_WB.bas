Attribute VB_Name = "UDFs_QB_WB"
Option Explicit

' NOTE ON USING REFERENCES TO THIS ADDIN FROM OTHER PROJECTS
' The calling project must have a reference set to this project.
' You set this reference in the VBE (Visual Basic Editor)
' by going to Tools > References and selecting the relevant library/project.
'
' Don't use simply FunctionName, to prevent naming naming conflicts.
' Using simply the function will work but very dangerous (because of duplicate names)
'
' Once a reference is established,  call the function or sub using
' the syntax ProjectName.ModuleName.FunctionName
'

Public UDFs_QB_WB_DATE As Date
'


Sub ClearFillandComments_Click()
    ' If worksheet Tab color is anything but default, skip
    If ActiveWorkbook.ActiveSheet.Tab.Color <> False Or _
        ActiveWorkbook.ActiveSheet.Tab.ThemeColor <> 0 Then Exit Sub
    Call ClearFillandComments(ActiveWorkbook.ActiveSheet)
End Sub
Sub qbCheck_NegBalances_click()   ' Ctrl + N
    ' If worksheet Tab color is anything but default, skip
    If ActiveWorkbook.ActiveSheet.Tab.Color <> False Or _
        ActiveWorkbook.ActiveSheet.Tab.ThemeColor <> 0 Then Exit Sub
    Call ClearFill(ActiveWorkbook.ActiveSheet)
    Call qbCheck_NegBalances
End Sub
Sub Check_QBWorksheet_click()       ' Ctrl + T   Run ALL Tests on Active Worksheet
Dim t1 As Date, tsecs As Double, m As String, answer As Long
    ' If worksheet Tab color is anything but default, skip
    If ActiveWorkbook.ActiveSheet.Tab.Color <> False Or _
        ActiveWorkbook.ActiveSheet.Tab.ThemeColor <> 0 Then
        m = "Active Sheet Tab Color is NOT default Color!" & vbCrLf & vbCrLf
        m = m & "Exiting Sub..."
        answer = MsgBox(m, vbOKOnly, ActiveSheet.name)
        Exit Sub
    End If
    t1 = Timer
    Call ClearFillandComments(ActiveSheet) ' Clear All Comments and and Fill
    Call qbCheck_Accounts_Class(ActiveSheet) ' Check Accounts and Class Type Match
    Call qbCheck_NegBalances      ' Highlight Negative Balances Cells in Red
    Call CommentDisplay_On
    'tsecs = Format((Timer - t1), "0.0000")
    'MsgBox "Check_QBWorksheet_click() Done...   Time elasped = " & tsecs & " Seconds" & _
    '        vbCrLf & "   Which is " & TimeString(CLng(tsecs), True)
End Sub
Sub qbCheck_Accounts_Class_click() ' NEEDS MORE TESTING
    ' If worksheet Tab color is anything but default, skip
    If ActiveWorkbook.ActiveSheet.Tab.Color <> False Or _
        ActiveWorkbook.ActiveSheet.Tab.ThemeColor <> 0 Then Exit Sub
    Call ClearFill(ActiveWorkbook.ActiveSheet)
    Call qbCheck_Accounts_Class(ActiveWorkbook.ActiveSheet)
End Sub
'Public Function VBqbLOOKUP(arg1_rng As Variant, ByVal RowHeader As Variant, ByVal colHeader As Variant, Optional ByVal Not_Found As Long = 0) As Variant
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    '  This UDF (user defined function) calls qbLoookup but with Verbatum (no Meta Characters)
'    '  If Meta Charachers happen to be in RowHeader or colHeader, then place a tilde (~) before the Meta character
'    '
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'
'    ' DNU THIS FUNCTION - NEEDS MORE DEBUGGING
'
'    Dim nRowheader As Variant, ncolHeader As Variant
'
'    nRowheader = Application.WorksheetFunction.Substitute(RowHeader, "*", "~*")
'    nRowheader = Application.WorksheetFunction.Substitute(nRowheader, "?", "~?")
'    nRowheader = Application.WorksheetFunction.Substitute(nRowheader, "#", "~#")
'    nRowheader = Application.WorksheetFunction.Substitute(nRowheader, "[", "~[")
'
'    ncolHeader = Application.WorksheetFunction.Substitute(colHeader, "*", "~*")
'    ncolHeader = Application.WorksheetFunction.Substitute(ncolHeader, "?", "~?")
'    ncolHeader = Application.WorksheetFunction.Substitute(ncolHeader, "#", "~#")
'    ncolHeader = Application.WorksheetFunction.Substitute(ncolHeader, "[", "~[")'
'
'    'MsgBox "RowHeader = " & RowHeader & " colHeader = " & colHeader
'    VBqbLOOKUP = qbLOOKUP(arg1_rng, nRowheader, ncolHeader, Not_Found)
'End Function

Public Function qbLOOKUP(arg1_rng As Variant, RowHeader As Variant, colHeader As Variant, _
                Optional Not_Found As Long = 0, Optional myDebugFlag As Boolean = False, _
                Optional Negative_Flip As Boolean = False) As Variant
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '  This UDF (user defined function) is for use with QuickBook reports exported to spreadsheets
    '  It returns the value of the cell cross referenced by RowHeader and ColHeader in the Range arg1_rng.
    '  If more than one RowHeader match is found, use the LAST MATCHING ROW FOUND in arg1_rng.
    '  If more than one ColHeader match is found, use the FIRST MATCHING COLUMN FOUND in arg1_rng.
    '
    ' Arguments:
    '  1. arg1_rng is the search Range
    '  2. Rowheader is the Row cross reference
    '  3. Colheader is the Column cross reference
    '  4. (Optional) Not_Found argument:
    '       -1 debug message boxes popup
    '        0 return 0 if not found (default)
    '       +1 return "Not Found" if not found
    '  5.  (Optional) Negative_Flip argument:
    '       -1 return the value found times -1 (ie. flip the negative sign of the value)
    '       other than -1,  returns the value found
    '
    '  Metacharacters (wildcards) are allowed in the RowHeader and ColHeader arguments.
    '  ?   Is any single char
    '  #   Is any single digit(0-9)
    '  *   Is zero or more characters. Note "*" alone is not allowed because it matches EVERYTHING
    '  [charlist]  Any char in the list
    '  [!charlist] Any char NOT in the char list
    '
    ' EXAMPLE:
    '  A question(?) mark matches any single character;
    '  An asterisk (*) matches any sequence of characters.
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim c_index As Variant, ret1 As Variant, ret2 As Variant, fill As Range
    ' Always use Variant for Application.function return values
    Dim cell As Range, r_index As Long, i As Long, y As Long, c2 As String
    Dim t1 As Date, rng As Range, ws As Worksheet, sheetName As String
    Dim var1 As Variant, m As Variant, ans As Long
    Dim answer As Variant
    
    t1 = Timer
    On Error Resume Next
    
    '
    '*STEP 1. ******** Validate Input *******************************************
    '
    Select Case VarType(arg1_rng)
        Case vbObject:
            Set rng = arg1_rng
            sheetName = rng.Worksheet.name
        Case vbString:
            sheetName = arg1_rng
            Set ws = Worksheets(sheetName)
            Set rng = ws.Cells
        Case vbArray + vbVariant:  ' Range
            Set rng = arg1_rng
            sheetName = rng.Worksheet.name
        Case Else
            'MsgBox "qbLOOKUP(): arg1_rng has unknown VarType " & VarType(arg1_rng)
            m = "qbLOOKUP(): arg1_rng has unknown VarType " & VarType(arg1_rng)
            myDebug m, myDebugFlag
    End Select
    
    If Not_Found = 1 Then qbLOOKUP = "Not Found" ' set default return value
   
    If rng Is Nothing Or IsEmpty(rng) Then
        'MsgBox "qbLOOKUP() EXITING:" & "    arg1_rng is EMPTY "
        m = "qbLOOKUP() EXITING:" & "    arg1_rng is EMPTY "
        myDebug m, myDebugFlag
        Exit Function
    End If
    ' MsgBox ("Selection Address = " & Selection.Address)
    If IsEmpty(RowHeader) Or RowHeader = "*" Then
        m = "qbLOOKUP() EXITING:" & "    RowHeader cell " & RowHeader.Address & " is EMPTY "
        myDebug m, myDebugFlag
        Exit Function
    End If
    If IsEmpty(colHeader) Or colHeader = "*" Then
        m = "qbLOOKUP() EXITING:" & "   RowHeader cell " & RowHeader.Address & "  = " & RowHeader & vbCrLf & _
        "    ColHeader cell " & colHeader.Address & " is EMPTY "
        myDebug m, myDebugFlag
        Exit Function
    End If
    ret1 = Application.CountIf(rng, RowHeader)
    ret2 = Application.CountIf(rng, colHeader)
      
    If ret1 = 0 Then ' Does RowHeader exist in rng?
        m = "qbLOOKUP(): Can't find '" & RowHeader & "'"
        'rng.Select
        'MsgBox "qbLOOKUP(): Can't find '" & RowHeader & "'"
        myDebug m, myDebugFlag
        'If Not_Found = -1 Then MsgBox "qbLOOKUP(): Can't find '" & RowHeader & "'"
        If Not_Found = -1 Then Debug.Print "qbLOOKUP(): Can't find '" & RowHeader & "'"
    End If
    
    If ret2 = 0 Then ' Does ColHeader exist in rng?
        m = "qbLOOKUP(): Can't find '" & colHeader & "'"
        myDebug m, myDebugFlag
        'If Not_Found = -1 Then MsgBox "qbLOOKUP(): Can't find '" & colHeader & "'"
        If Not_Found = -1 Then Debug.Print "qbLOOKUP(): Can't find '" & colHeader & "'"
    End If
    
    If (ret1 < 1 Or ret2 < 1) Then Exit Function ' Row Header or ColHeader NOT exist

    '
    ' *STEP 2. * Find ColHeader and get the relative c_index ***************************
    '
    For i = 1 To 5  ' QB has 5 Column Header Rows max  Change this later!
        ret1 = Application.CountIf(rng.Rows(i), colHeader)
        If (ret1) > 0 Then ' Does RowHeader exist in this row?
            c_index = Application.Match(colHeader, rng.Rows(i), 0) ' Looks in Row(i) for ColHeader
        End If
    Next i
    If IsError(c_index) Then
        m = "qbLOOKUP(): ERROR in Application.Match() " & _
        "   c_index = " & c_index & _
        "   ColHeader = " & colHeader  ' Debug
        myDebug m, myDebugFlag
        If Not_Found = -1 Then MsgBox "qbLOOKUP(): ERROR in Application.Match() " & _
        "   c_index = " & c_index & _
        "   ColHeader = " & colHeader  ' Debug
        Set qbLOOKUP = Nothing
        Exit Function '
    End If

    '
    ' *STEP 3. * Find RowHeader and get the relative r_index **************
    '
    i = qbHowManyRowLevels(rng.Worksheet) ' determines Columns to search like Rng.Columns("A:G")
    
    '   Expression                      is Column  Comment
    '   Range("B2:B3").Columns("B")     C          Column character interpreted as number, e.g. "B" is always 2nd column, even if not in the original range
    '   Range("B2:B3").Columns("B:C")   C to D     Column character interpreted as number, e.g. "C" is always 3rd column which ic Column D
    '   .Columns is RELATIVE!... find Last Pattern in range from first column the number of QB ROW HEADERS
    
    If i = 0 Then ' QB ROW HEADERS not found
        'ans = MsgBox("Warning!" & vbCrLf & "Name: 'QBROWHEADERS' is NOT defined in Name Manager" &
        '    " in the Scope of this worksheet!", vbInformation)
        r_index = LastPatternRow(rng.Columns("A:A"), RowHeader, Not_Found) ' default to First column of Rng
        'Application.Calculation = xlManual ' Turn OFF Calculation here because of repeat calls
    Else
        c2 = Chr(65 + i - 1) ' A + Number of columns from QB Row Headers - 1
        c2 = "A:" & c2 ' The total number of columns with QB Row Headers
        'MsgBox "c2 = '" & c2 & " '"
        r_index = LastPatternRow(rng.Columns(c2), RowHeader, Not_Found) ' Search to First column to last QB Row HEADERS column
    End If
    If r_index <= 0 Then
        m = "qbLOOKUP(): found Nothing" & _
        "   r_index = " & r_index & _
        "   Row Header = " & RowHeader  ' Debug msg
        myDebug m, myDebugFlag
        Set qbLOOKUP = Nothing ' Nothing threw an error when assigning to an object
        'Set qbLOOKUP = Empty
        Exit Function '
    End If
    
    'Fill_Yellow (rng.Cells(r_index, c_index))
    '
    ' *STEP 4. ***  Set result **************
    '
    answer = rng.Cells(r_index, c_index) ' return range variant
    If Not Negative_Flip Then
         qbLOOKUP = answer
    Else ' Return NEGATIVE_FLIP when flag set
        If IsNumeric(answer) Then ' Has to be numeric to flip
            qbLOOKUP = answer * -1 ' flip sign
        End If
    End If
    m = "qbLOOKUP() Exited: Value = " & qbLOOKUP & "  Row = " & r_index & "   Column = " & c_index & "(" & Chr(64 + c_index) & ")" & _
    "   Time elasped = " & Format((Timer - t1), "0.0000") & " Seconds"
    myDebug m, myDebugFlag
End Function ' qbLookup
Private Sub myDebug(m As Variant, flag As Boolean)
    If Not flag Then Exit Sub
    Debug.Print m
End Sub
'UDF:
'  Adds two Dates or Date plus Time and Returns a date
'  Args can be text, date time, int, float, etc
'  Text can be "11/19/24", "8:20 PM", "11/19/24 8:20 PM", etc.
Function myDatePlusTime(arg1 As Range, arg2 As Range)
    myDatePlusTime = CDate(arg1) + CDate(arg2)
End Function
'UDF:
'  Returns "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" of a Date
'  Arg1 must be a Date format
Function DayOfWeek(arg1 As Range)
    DayOfWeek = Choose(Weekday(arg1), "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
End Function

'UDF:
'  Remove All Non-Numeric Characters from cell
'https://stackoverflow.com/questions/43076380/how-can-i-remove-non-numeric-characters-in-cells-and-concatenate-result-to-url
Function GetNums(target As Range)
    Dim MyStr As String, i As Integer
    MyStr = ""
    If Len(target.Value) = 0 Then GoTo GoExit
    If target.Value = "None" Then GoTo GoNone
    For i = 1 To Len(target.Value)
        If IsNumeric(Mid(target, i, 1)) Then MyStr = MyStr & Mid(target, i, 1)
        If Mid(target, i, 1) = "." Then
            If IsNumeric(Mid(target, i + 1, 1)) Then MyStr = MyStr & Mid(target, i, 1)
        End If
    Next i
    GoTo GoExit
GoNone:
    GetNums = "None"
    Exit Function
GoExit:
    GetNums = Val(MyStr)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Sub replaces references in ALL cells in active worksheet that have a forumula 'qbLookup(SheetName1' to 'qbLookup(SheetName2' on the Active Worksheet
' Also replace any "'" in the first char followed by "=qblookup(" or "=-qblookup("
'
'   Optional Not_Found argument Values:
'       1 =  If SheetName1 does not exist in Workbook, ignore SheetName1 and use "qbLOOKUP(*!" to replace
'       0 =  If SheetName1 or SheetName2 does not exist in Workbook, Exit Sub
'      -1 =  Put up a msgbox with error message of the problem, then exit.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Switch_SheetName(SheetName1, SheetName2, Optional Not_Found As Long = 0)
    Dim str1 As String, str2 As String
    ' Parameter checks
    On Error Resume Next
    If SheetName1 = "" Then Exit Sub
    If SheetName2 = "" Then Exit Sub
    
    If Not_Found = 1 Then ' replace no matter what SheetName1 is
        str1 = "qbLOOKUP(*!"
    Else
        str1 = "qbLOOKUP(" & SheetName1 & "!"
    End If
    str2 = "qbLOOKUP('" & SheetName2 & "'!" ' always add the apostrophe "'" for saftey sake.

    'If Not SheetExists(SheetName2) Then
    '   Debug.Print "Switch_SheetName(): Sheet2 Does NOT exist. Sheetname = '" & SheetName2 & "'"
    '   If Not_Found = -1 Then MsgBox "Switch_SheetName(): Sheet2 Does NOT exist. Sheetname = '" & SheetName2 & "'"
    '   Exit Sub
    'End If
    
    Cells.Replace what:=str1, Replacement:=str2, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:= _
    False, SearchFormat:=False, ReplaceFormat:=False
    
    On Error GoTo 0
End Sub
Function WorksheetExists(sname) As Boolean ' Works on ActiveWorkbook
    WorksheetExists = Evaluate("ISREF('" & sname & "'!A1)") 'a trick to see if sheet exits!!
End Function
Function SheetExists(ByVal strSheetName As String, Optional wbWorkbook As Workbook) As Boolean
    If wbWorkbook Is Nothing Then Set wbWorkbook = ActiveWorkbook           'or ThisWorkbook - whichever appropriate
    Dim obj As Object, sname As String
    If hasQuotes(strSheetName) = True Then  'Sheet Names sometimes have single quotes
        sname = stripQuotes(strSheetName) ' Strip single quotes from name
    Else
        sname = strSheetName
    End If
    Err.Clear
    On Error GoTo HandleError
    Set obj = wbWorkbook.Sheets(sname)
    SheetExists = True
    Exit Function
HandleError:
    Err.Clear
    SheetExists = False
End Function
Function hasQuotes(ByVal sname As String) As Boolean ' see if sname enclosed in quotes
    hasQuotes = False
    If Left(sname, 1) = "'" Then hasQuotes = True 'some Sheet Names have single quotes
End Function
Function stripQuotes(ByVal sname As String) As String '
' Some Sheet Names are single quotes, Remove first and last chars
    stripQuotes = Mid(sname, 2, Len(sname) - 2)
End Function
Private Sub qbCheck_Accounts_Class(sheet As Worksheet)
' NEEDS MORE WORK
' Check Accounts with Class shown in ColHeader are proper
' Highlight nonzero cells with account and class that do NOT match
'
    Dim levels As Long, r As Long, c As Long, totalcolumns As Long
    Dim colH As String, rowH As String, cell As Range
    levels = qbHowManyRowLevels(sheet)
    totalcolumns = Last(2, sheet.Cells)
    For c = levels + 1 To totalcolumns
        For r = 1 To 2 ' first two rows Change this later!
            Set cell = sheet.Cells(r, c)
            'cell.Select  ' debug
            If cell Like "(*)" Or cell Like "Total *" Or cell = "TOTAL" _
            Or IsEmpty(cell) Or cell Like "MIP-*" = True Then
                Debug.Print cell & "  ...Skip This COLUMN HEADER!"
            Else
                'Debug.Print "Checking this Column Header!  " & cell
                Call Check_qbRows(sheet, cell) ' the rows with column header
            End If
        Next
    Next
End Sub
Private Sub qbCheck_NegBalances()
' Highlight any Negative balances in Red
Dim levels As Long, r As Long, c As Long, totalcolumns As Long
Dim cell As Range
    levels = qbHowManyRowLevels(ActiveSheet)
    totalcolumns = Last(2, Cells)
    For c = levels + 1 To totalcolumns
        For r = 1 To Last(1, Cells) '
            Set cell = Cells(r, c)
            'cell.Select
            If cell.Value < 0 Then Call Fill_Red(cell)
        Next
    Next
End Sub
Private Sub test_qbFill_Acct_row()
    Call qbFill_Acct_Row(ActiveSheet, Selection)
End Sub
Sub qbFill_Acct_Row(sheet As Worksheet, acct As Range)
' acct arg can use Meta Characters
' Highlight any Liability acct rows in Red, all other acct matches in green
Dim levels As Long, r As Long, c As Long, account_col As Long, totalcolumns As Long
Dim cell As Range, red_green As Boolean
    
    If (Mid(acct.Value, GetPositionOfFirstNumericCharacter(acct.Value), 1)) = "2" Then red_green = True
    levels = qbHowManyRowLevels(sheet)
    totalcolumns = Last(2, sheet.Cells)
    
    If levels = 1 Then
        account_col = find_header_col("Account", sheet.Cells)
    End If
    If account_col > 0 Then
        For r = 1 To Last(1, sheet.Cells) '
            Set cell = sheet.Cells(r, account_col)
            'cell.Select
            If cell.Value Like acct.Value Then
                Call fill_row(sheet, r, totalcolumns, red_green)
            End If
        Next
    Else
        For c = 1 To levels
            For r = 1 To Last(1, sheet.Cells) '
                Set cell = sheet.Cells(r, c)
                'cell.Select
                If cell.Value Like acct.Value Then
                    Call fill_row(sheet, r, totalcolumns, red_green)
                End If
            Next
        Next
    End If
End Sub
Private Function find_header_col(hdr As String, rng As Range) As Long
Dim r As Long, c As Long, ret1 As Long
    On Error Resume Next
    For r = 1 To 2  ' QB has 2 Column Header Rows max  Change this later!
        For c = 1 To Last(2, rng)
            If rng(r, c).Value Like hdr Then find_header_col = c
        Next
    Next r
End Function

Private Sub fill_row(sheet As Worksheet, row As Long, end_col As Long, red As Boolean)
Dim rng As Range
    Set rng = sheet.Range(Cells(row, 1), Cells(row, end_col))
    If red Then
        Call Fill_Red(rng)
    Else
        Call Fill_Green(rng)
    End If
End Sub
'https://stackoverflow.com/questions/3547744/vba-how-to-find-position-of-first-digit-in-string
Function GetPositionOfFirstNumericCharacter(ByVal s As String) As Integer
Dim currentCharacter As String, i As Long
    For i = 1 To Len(s)
        currentCharacter = Mid(s, i, 1)
        If IsNumeric(currentCharacter) = True Then
            GetPositionOfFirstNumericCharacter = i
            Exit Function
        End If
    Next i
End Function

Private Sub Check_qbRows(sheet As Worksheet, colHeader As Range)
' Check Accounts with Class shown in ColHeader are proper
' Highlight any nonzero amount with improper account and class match in Blue
' Highlight any ColHeader that doesn't match a class in Yellow

    Dim r As Long, c As Long, lastRow As Long, levels As Long, m As String
    Dim RowHeader As Range, lookedup As Range, classtype As Range
    levels = qbHowManyRowLevels(sheet)
    lastRow = Last(1, sheet.Cells)
    For r = 1 To lastRow
        For c = 1 To levels ' all ther QB Row Header columns
            Set RowHeader = sheet.Cells(r, c)
            ' rowheader.Select
            If RowHeader Like "[^0-9]*" Then
                Set lookedup = sheet.Cells(RowHeader.row, colHeader.Column)
                If IsEmpty(lookedup) = True Or lookedup.Value = 0 Then
                    'Debug.Print lookedup & "' lookedup IsEmpty or = 0  ...Ok"
                Else
                    'lookedup.Select ' for debug
                    On Error Resume Next
                    Set classtype = Get_MyClassType(ThisWorkbook.Sheets("Class List").Cells, "*" & colHeader.Value)
                    If Err <> 0 Then
                        m = "You are checking the Class without the 'Class List' worksheet!"
                        'MsgBox M, vbExclamation
                        Debug.Print m
                        Exit Sub
                    End If
                    If classtype Is Nothing Then
                        Debug.Print "Can't find a Class Type for '" & colHeader.Value & "'"
                        Call Fill_Yellow(colHeader)
                        Call Fill_Yellow(lookedup) ' can't find class type for column header
                        Exit Sub
                    End If
                    If Check_qbAccount_Type(RowHeader, classtype) = False Then
                        Call Fill_Blue(lookedup) ' class type does not have a 'YES' for the account
                        Call AddToComment(lookedup, "Class Type '" & classtype & "' not valid for this Account")
                    End If
                End If
            End If
            Next c
    Next r
End Sub
Private Function Check_qbAccount_Type(RowHeader As Range, classtype As Range) As Boolean
' Returns True if RowHeader (ie. Account) and classtype match
    Dim accNum As Long, answer As String
    Dim CtHdrCell As Range, AccHdrCell, searchRange As String
    
    On Error Resume Next
    Check_qbAccount_Type = False ' default return value
    If RowHeader Is Nothing Then Exit Function
    If IsEmpty(RowHeader) Then Exit Function
    If classtype Is Nothing Then Exit Function
    If IsEmpty(classtype) Then Exit Function
        
    accNum = Val(RowHeader) ' converts contents to to number
                                ' stops at first non numeric chars.
    If accNum <= 0 Then
        Debug.Print "Can't find Account Number in row header."
        Exit Function
    End If
    ' Range Must Include 'ACCNUM' and  "Class Type" columns of the Chart of Accounts sheet
    ' Removed the following line because took up to 5 minutes to test one balance sheet.
    'answer = qbLOOKUP(ThisWorkbook.Sheets("Chart of Accounts").Range("H:V"), accnum, classtype)
    Set CtHdrCell = ThisWorkbook.Sheets("Chart of Accounts").Range("1:1").Find(classtype) ' find account type column
    Set AccHdrCell = ThisWorkbook.Sheets("Chart of Accounts").Range("1:1").Find("ACCNUM") ' find ACCNUM column
    If IsEmpty(CtHdrCell) Then
        Debug.Print ("  Check_qbAccount_Type(): Couldn't find classtype " & classtype & "in Chart of Accounts header row")
        Exit Function
    End If
    If IsEmpty(AccHdrCell) Then
        Debug.Print ("  Check_qbAccount_Type(): Couldn't find column 'ACCNUM' in Chart of Accounts header row")
        Exit Function
    End If
    searchRange = Chr(64 + AccHdrCell.Column) & ":" & Chr(64 + CtHdrCell.Column)
    'MsgBox ("SearchRange = " & SearchRange & "   index = " & CtHdrCell.Column - AccHdrCell.Column + 1)
    answer = Application.VLookup(accNum, ThisWorkbook.Sheets("Chart of Accounts").Range(searchRange), CtHdrCell.Column - AccHdrCell.Column + 1, 0)
    If IsEmpty(answer) Then Exit Function
    
    If answer = "YES" Then Check_qbAccount_Type = True ' Match
End Function
Public Function Header_Index(rng As Range, ByVal colHeader As String) As Long
rng.Select ' debug
        Header_Index = Application.Match(colHeader, rng, 0) 'Looks in rng for colHeader string; return index
End Function
Function LookUp_ClassType(account As Long, searchRange As Range, Class_type As String) As String
    Dim index As Long, result As Variant
    index = Header_Index(searchRange.Rows(1), Class_type)
    If index < 0 Then Exit Function ' not found
    result = Application.VLookup(account, searchRange, index, 0)
    If IsError(result) Then
        Debug.Print ("LookUp_ClassType(): ERROR on Applicaton.VLookup " & _
        "   Account = " & account & "   Class_type = " & Class_type)  ' Error 2042 is the VBA equivalent of #N/A
        Exit Function '
    End If
    LookUp_ClassType = result
End Function
Function LookUp_ByIndex(LookupKey As Variant, searchRange As Range, index As Long) As Variant
    LookUp_ByIndex = Application.VLookup(LookupKey, searchRange, index, 0)
    'If IsError(LookUp_ByIndex) Then
    '    Debug.Print ("LookUp_ByIndex(): ERROR on Applicaton.VLookup " & _
    '    "   LookupKey = " & LookupKey & "   Index = " & Index)  ' Error 2042 is the VBA equivalent of #N/A
    '    LookUp_ByIndex = Empty
    'End If
End Function
Private Function Get_MyClassType(classlist As Range, classname As String, Optional Not_Found As Long = 0) As Range
    Dim r_index As Long
    ' Find RowHeader and get the relative r_index
    r_index = LastPatternRow(classlist, classname, Not_Found)

    If r_index <= 0 Then
        Debug.Print "Get_MyClassType(): Can't find" & _
        "   classname = '" & classname & "'" & _
        "   r_index = " & r_index

        Set Get_MyClassType = Nothing
        Exit Function '
    End If
    'MsgBox "Result: " & vbCrLf & _
    '"r_index =  " & r_index & vbCrLf & _
    '"Value = '" & Application.index(rng, r_index, c_index) & "'" ' Debug
        
    Set Get_MyClassType = classlist(r_index, 8) ' 8th column over in classlist is MyClassType
End Function
Function LastPatternRow(r As Variant, pattern As Variant, Optional Not_Found As Long = 0) As Long
    Dim cell As Range, searchRange As Range, LastPattern As Range
    Dim y As Long
    Set LastPattern = Nothing
    LastPatternRow = 0
    If Not_Found = -1 Then Debug.Print "    LastPatternRow() Enterd: '" & pattern & "' Rows = " & r.Rows.count & "   Columns = " & r.Columns.count
    y = 1
    Set searchRange = RangeTrimmed(r, Not_Found)
    If searchRange Is Nothing Then
        If Not_Found = -1 Then Debug.Print "    LastPatternRow(): Range r is Outside of last used cell in Worksheet."
        Exit Function
    End If
    For Each cell In searchRange
        ' Check the pattern
        'Application.EnableEvents = False
        'SearchRange.Worksheet.Activate
        'SearchRange.Select
        'MsgBox "LastPatternRow() Search Range and pattern" & pattern
        Application.EnableEvents = True
        ' Used 'Like' so  metacharacters work
        If cell Like pattern Then
            If Not_Found = -1 Then Debug.Print "   Match = '"; cell & "'" & " Row " & cell.row
            Set LastPattern = cell
        End If
        y = y + 1
    Next
    If Not LastPattern Is Nothing Then LastPatternRow = LastPattern.row  ' return Row number
    'Debug.Print "    LastPatternRow() Exited:  LastPattern Row = " & LastPattern.Row & "   y = " & y
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   RangeTrimmed()
'   Return  range trimmed to LAST USED row and LAST USED column of the Worksheet.
'   Nothing is returned if Rng is COMPLETELY OUTSIDE of the last Column or last Row of the Worksheet.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function RangeTrimmed(rng As Variant, Optional Not_Found As Long = 0) As Variant
    Dim lastRow As Long, lastCol As Long
    Dim rng2 As Range
    'Debug.Print "    RangeTrimmed() Before: Rows = " & rng.Rows.Count & "   Columns = " & rng.Columns.Count
    '   FYI ... Last() returns last used cell of the WORKSHEET (not the Rng)
    Set RangeTrimmed = Nothing
    lastRow = Last(1, rng)
    lastCol = Last(2, rng)
    If lastRow = 0 Or lastCol = 0 Then Exit Function ' Rng OUTSIDE of last used cell in Worksheet
    
    '    Intersect returns NOTHING if ranges don't intersect
    Set RangeTrimmed = Intersect(rng, Range(rng.Cells(1, 1), rng.Cells(lastRow, lastCol)))
    If Not RangeTrimmed Is Nothing Then
        If Not_Found = -1 Then Debug.Print "    RangeTrimmed()  After: Rows = " & RangeTrimmed.Rows.count & "   Columns = " & RangeTrimmed.Columns.count
        'On Error Resume Next
        'RangeTrimmed.Select  ' Debug this line only works on Active worksheet
        'err.clear
        'MsgBox "RangeTrimmed() Here is the RangeTrimmed" ' Debug
    Else
        If Not_Found = -1 Then Debug.Print "    RangeTrimmed() After: ERROR...Intersect returned Nothing."
    End If
End Function
Public Function qbHowManyRowLevels(sheet As Worksheet) As Long
    Dim str As String
    qbHowManyRowLevels = 0
    On Error Resume Next
    ' QB Row Header LEVELS is a Defined Name "QBROWHEADERS" in Name Manager
    If Not IsEmpty(sheet) Then
        str = sheet.Names("QBROWHEADERS").Value
        qbHowManyRowLevels = Val(Right(str, Len(str) - 1)) ' strips the "=" off the string
    End If
    If qbHowManyRowLevels = 0 Then
        Debug.Print "Missing defined Name 'QBROWHEADERS' in Formulas:Name Manager for this worksheet! Setting to default: 7"
        qbHowManyRowLevels = 7
    End If
End Function
Public Function qbPeriodEnding(ws As Worksheet) As String
Dim str As String, yr As String, mo As String, dy As String
        qbPeriodEnding = ""
        On Error Resume Next
        ' QB's Period Ending is defined as Defined Name "QBENDDATE" in each QB worksheet
        If Not IsEmpty(ws) Then
            str = ws.Names("QBENDDATE").Value
            yr = Mid(str, 2, 4) ' first char is '='
            mo = Mid(str, 6, 2)
            dy = Mid(str, 8, 2)
            qbPeriodEnding = DateValue(mo & "/" & dy & "/" & yr) ' very important line... format date properly
        End If
End Function
Public Function CellFormula(rng As Range) As String
    ' Return cell formula as a string
    CellFormula = rng.Formula
End Function
Public Function IsRangeNothing(name As String, Optional prompt As Boolean = True) As Boolean
Dim answer As Long
    On Error Resume Next
    If Range(name) Is Nothing Then
        IsRangeNothing = True
        If prompt Then
            answer = MsgBox("Range Name '" & name & "' is not defined in this worksheet.", _
                vbInformation & vbOKOnly, ActiveSheet.name)
        End If
    End If
End Function

Private Sub test_ConvertTextToNumber()
    ConvertTextToNumber Selection
End Sub
' Fixes numbers stored as text
Private Sub ConvertTextToNumber(rng As Range)
Dim cell As Range
  On Error Resume Next
  ' Iterate through each cell in the range
  For Each cell In rng
    ' Check if the cell contains a value and if it's text
    If Not IsEmpty(cell.Value) Then
      ' Multiply the cell value by 1 to convert it to a number
       cell.Value = cell.Value * 1
    End If
  Next cell
End Sub

Private Sub test_AddToComment()
    Call AddToComment(ActiveSheet.Range("I5"), "Hello World!")
End Sub
Sub AddToComment(r As Range, NewComment As Variant)
Dim OldComment As Variant
    'If error then old comment does not exist
    On Error Resume Next
    OldComment = r.Comment.Text
    If Not OldComment Is Nothing Then NewComment = OldComment & vbCrLf & NewComment
    r.Comment.Delete
    r.AddComment (NewComment)
    On Error GoTo 0
End Sub
Sub CommentDisplay_On()
    Call ChangeCommentDisplay(xlCommentAndIndicator)
End Sub
Sub CommentDisplay_Off()
    Call ChangeCommentDisplay(xlCommentIndicatorOnly)
End Sub
Private Sub ChangeCommentDisplay(indicatorType As Long)
    Application.DisplayCommentIndicator = indicatorType
End Sub
Private Sub test_Fill_Match()
    Call Fill_Match(Selection, ActiveWorkbook.Sheets("Test qbLOOKUP").Range("B1"))
End Sub
Sub Fill_Match(r1 As Range, r2 As Range)       ' fill interior of all cells in range r2 to match range r1
' This Sub Does NOT work when called from a UDF!
On Error Resume Next
        With r2.Interior
            .pattern = r2.Interior.pattern
            .PatternColorIndex = r2.Interior.PatternColorIndex
            .Color = r2.Interior.Color
            .TintAndShade = r2.Interior.TintAndShade
            .PatternTintAndShade = r2.Interior.PatternTintAndShade
        End With
End Sub
Private Sub test_Fill()
    Fill_Clear Selection
    'Fill_Red Selection
    MsgBox "is_Fill_NON = " & is_Fill_None(Selection)
    Fill_Clear Selection
End Sub
Sub ClearFill(ws As Worksheet, Optional HEADER As Long = 0) ' Clear ALL cells of fill colors
    With ws.Cells.Interior
        .pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
Sub ClearFillandComments(ws As Worksheet)
    ' Clears ALL Comments then Fill
    ws.Cells.ClearComments
    Call CommentDisplay_Off
    Call ClearFill(ws)
End Sub
Function is_Fill_None(r As Range) As Boolean
    With r.Interior
        If .pattern <> xlNone Then Exit Function
        If .TintAndShade <> 0 Then Exit Function
        If .PatternTintAndShade <> 0 Then Exit Function
    End With
    is_Fill_None = True
End Function
Sub Fill_Clear(r As Range)
    With r.Interior
        .pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
Private Sub test_Fill_Blue()
    Call Fill_Blue(Selection)
End Sub
Sub Fill_Blue(r As Range) ' fill all cells in range with Light Blue
    With r.Interior
        .pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2 ' Light Blue
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
End Sub
Sub Fill_Yellow(ByVal r As Range) ' fill all cells in range with Yellow
    With r.Interior
        .pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535 ' Yellow
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
Sub Fill_Red(r As Range) ' fill all cells in range with Light Red
    With r.Interior
        .pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 8224255 ' Light Red
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
Sub Fill_Green(r As Range) ' fill all cells in range with Light Green
    With r.Interior
        .pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419  ' Light Green
        .PatternTintAndShade = 0
    End With
End Sub
' You Only need to run this macro one time.
Private Sub Register_qbLOOKUP() ' Register user defined function qbLOOKUP to a separate category
    Dim vqbArgDescr(1 To 5) As Variant
    vqbArgDescr(1) = "Name of the QuickBooks worksheet or a range to search."
    vqbArgDescr(2) = "Pattern search for the Row."
    vqbArgDescr(3) = "Pattern search for the Column Header."
    vqbArgDescr(4) = "Optional (default = 0): 0 Returns 0 if not found; 1 Returns 'Not Found' if not found."
    vqbArgDescr(5) = "Optional (default = False): = True debug.prints extra stuff."

    Application.MacroOptions _
                             Macro:="qbLOOKUP", Description:="Lookup a number in QuickBooks Statement", _
                             Category:="QB Addin User Functions", _
                             ArgumentDescriptions:=vqbArgDescr
End Sub
Private Sub UnRegister_qbLOOKUP() ' Register user defined function qbLOOKUP back to default category 14
    Dim vqbArgDescr(1 To 4) As Variant
    vqbArgDescr(1) = "Name of the QuickBooks worksheet or a range to search."
    vqbArgDescr(2) = "Pattern search for the Row."
    vqbArgDescr(3) = "Pattern search for the Column Header."
    vqbArgDescr(4) = "Optional (default = 0): 0 Returns 0 if not found; 1 Returns 'Not Found' if not found."
    vqbArgDescr(5) = "Optional (default = False): = True debug.prints extra stuff."

    Application.MacroOptions _
                             Macro:="qbLOOKUP", Description:="Lookup a number in QuickBooks Statement", _
                             Category:=14, _
                             ArgumentDescriptions:=vqbArgDescr
End Sub
Public Function TimeString(Seconds As Long, Optional verbose As Boolean = False) As String
' http://www.freevbcode.com/ShowCode.asp?ID=85
'   Takes number of seconds and returns
'   a time string containing hours, minutes, and seconds.
'   If the optional parameter Verbose is false, the return value looks like 1:31.08.
'   If it is true, the return value would be 1 hour, 31 minutes, and 8 seconds.
' EXAMPLE
'   if verbose = false, returns  something like
'       02:22.08
'   if true, returns
'       2 hours, 22 minutes, and 8 seconds
''''''''''''''''''''''''''''''''''''''''''''''''''
Dim lHrs As Long
Dim lMinutes As Long
Dim lSeconds As Long
    lSeconds = Seconds
    lHrs = Int(lSeconds / 3600)
    lMinutes = (Int(lSeconds / 60)) - (lHrs * 60)
    lSeconds = Int(lSeconds Mod 60)
Dim sAns As String
    If lSeconds = 60 Then
        lMinutes = lMinutes + 1
        lSeconds = 0
    End If
    
    If lMinutes = 60 Then
        lMinutes = 0
        lHrs = lHrs + 1
    End If
    
    sAns = Format(CStr(lHrs), "#####0") & ":" & _
      Format(CStr(lMinutes), "00") & "." & _
      Format(CStr(lSeconds), "00")
    
    If verbose Then sAns = TimeStringtoEnglish(sAns)
    TimeString = sAns
    End Function
Private Function TimeStringtoEnglish(sTimeString As String) As String
' Goes with function above here
Dim sAns As String
Dim sHour, sMin As String, sSec As String
Dim iTemp As Integer, sTemp As String
Dim iPos As Integer
    iPos = InStr(sTimeString, ":") - 1
    sHour = Left$(sTimeString, iPos)
    If CLng(sHour) <> 0 Then
        sAns = CLng(sHour) & " hour"
        If CLng(sHour) > 1 Then sAns = sAns & "s"
        sAns = sAns & ", "
    End If
    
    sMin = Mid$(sTimeString, iPos + 2, 2)
    iTemp = sMin
    
    If sMin = "00" Then
        sAns = IIf(Len(sAns), sAns & "0 minutes, and ", "")
    Else
       sTemp = IIf(iTemp = 1, " minute", " minutes")
       sTemp = IIf(Len(sAns), sTemp & ", and ", sTemp & " and ")
       sAns = sAns & Format$(iTemp, "##") & sTemp
    End If
    iTemp = Val(Right$(sTimeString, 2))
    sSec = Format$(iTemp, "#0")
    sAns = sAns & sSec & " second"
    If iTemp <> 1 Then sAns = sAns & "s"
    TimeStringtoEnglish = sAns
End Function
Sub BuildSheetNameTableCall() ' This had to be in a module to work
    Call OPEN_WB.Sheets("QB DUE To From").Build_SheetNameTable_Click ' Still not working?
End Sub
Public Function MyCDATE(arg1 As Variant)
    MyCDATE = CDate(arg1)
End Function

'**************************************************
'
' Split up long text cells in Active Sheet by
'   inserting an additional row below
'   then add split text to the additional row
'
'**************************************************

Sub Split_LongTextLines()
' Define the maximum allowed string length here
Const MAX_LEN As Long = 180
Const MAX_SPLITS As Long = 10

Dim ws As Worksheet
Dim rngCol As Range
Dim col As Long, lastRow As Long, firstRow As Long, currentRow As Long
Dim textToSplit As String, currentChunk As String, remainder As String
Dim spacePos As Long, splitCount As Long
Dim m As String, ans As Long
Dim lastUsedCol As Long
    
    Set ws = ActiveSheet
    Set rngCol = ActiveSheet.Cells ' Example range to process
    
    lastUsedCol = Last(2, rngCol)
    If lastUsedCol > 1 Then
        m = "Warning! More than one Column! Used Columns = " & lastUsedCol & "." & vbCrLf & vbCrLf
        m = m & "Do you want to continue?"
        ans = MsgBox(m, vbYesNo, "SplitLongStrings():")
        If ans <> vbYes Then Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Loop through each column in the specified range
    For col = 1 To lastUsedCol
        firstRow = rngCol.Cells(1, col).row
        lastRow = ws.Cells(ws.Rows.count, rngCol.Columns(col).Column).End(xlUp).row
        
        ' Loop backwards from bottom to top
        For currentRow = lastRow To firstRow Step -1
            ' Check if the cell contains an error before trying to read its string value
            If Not IsError(ws.Cells(currentRow, rngCol.Columns(col).Column).Value) Then
            
                textToSplit = Trim(ws.Cells(currentRow, rngCol.Columns(col).Column).Value)
                splitCount = 0
                
                ' Loop if string exceeds the constant length AND we haven't hit the split limit
                Do While Len(textToSplit) > MAX_LEN And splitCount < MAX_SPLITS
                    ' Find nearest white space looking forward from the max length position
                    spacePos = InStr(MAX_LEN, textToSplit, " ")
                    
                    ' If no space is found forward, stop splitting to avoid infinite loops
                    If spacePos = 0 Then Exit Do
                    
                    ' Separate the string based on the forward space position
                    currentChunk = Trim(Left(textToSplit, spacePos - 1))
                    remainder = Trim(Mid(textToSplit, spacePos + 1))
                    
                    ' Insert a completely blank row below the current cell
                    ws.Rows(currentRow + 1).Insert Shift:=xlDown
                    
                    ' Assign the split values
                    ws.Cells(currentRow, rngCol.Columns(col).Column).Value = currentChunk
                    remainder = "- " & remainder
                    ws.Cells(currentRow + 1, rngCol.Columns(col).Column).Value = remainder
                    
                    ' Increment split tracker and move pointer down to evaluate the remainder
                    splitCount = splitCount + 1
                    currentRow = currentRow + 1
                    textToSplit = remainder
                Loop
            
                ' Reset the row pointer back upward by the number of rows we just added
                currentRow = currentRow - splitCount
            End If
        Next currentRow
    Next col
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Last()
'
' http://www.rondebruin.nl/win/s9/win005.htm
' Use the "Last" function to Find the last row, column or cell in range or worksheet
'
' choice as argument
' 1 = last row
' 2 = last column
' 3 = last cell
'  row and columns numbers are absolute (not relative).
' Argument 2 is the range where you want to search in
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Last(choice As Long, rng As Variant)
    Dim lrw As Long
    Dim lCol As Long
    Select Case choice
        Case 1:
            On Error Resume Next
            Last = rng.Find(what:="*", _
                After:=rng.Cells(1), _
                LookAt:=xlPart, _
                LookIn:=xlFormulas, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious, _
                MatchCase:=False).row
            On Error GoTo 0
        Case 2:
            On Error Resume Next
            Last = rng.Find(what:="*", _
                After:=rng.Cells(1), _
                LookAt:=xlPart, _
                LookIn:=xlFormulas, _
                SearchOrder:=xlByColumns, _
                SearchDirection:=xlPrevious, _
                MatchCase:=False).Column
            On Error GoTo 0
        Case 3:
            On Error Resume Next
            lrw = rng.Find(what:="*", _
                After:=rng.Cells(1), _
                LookAt:=xlPart, _
                LookIn:=xlFormulas, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious, _
                MatchCase:=False).row
            On Error GoTo 0
            On Error Resume Next
            lCol = rng.Find(what:="*", _
                After:=rng.Cells(1), _
                LookAt:=xlPart, _
                LookIn:=xlFormulas, _
                SearchOrder:=xlByColumns, _
                SearchDirection:=xlPrevious, _
                MatchCase:=False).Column
            On Error GoTo 0
            On Error Resume Next
            Last = rng.Parent.Cells(lrw, lCol).Address(False, False)
            If Err.Number > 0 Then
                Last = rng.Cells(1).Address(False, False)
                Err.Clear
            End If
            On Error GoTo 0
    End Select
End Function
