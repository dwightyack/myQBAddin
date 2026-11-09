Attribute VB_Name = "modGoToSelected"
Option Explicit

Private Const MSG_TITLE As String = "Go To Selected"
Private Const USER_ERR As Long = vbObjectError + 1000

' Add-ins this module switched to IsAddin = False (for Rehide_Addins)
Private mShownAddins As Collection

'==============================================================================
' Convenience callers
'==============================================================================

' Assign this to a button / shortcut key.
Public Sub Go_To_Selected_Current()
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a cell that contains a QBLOOKUP() formula first.", vbExclamation, MSG_TITLE
        Exit Sub
    End If
    Go_To_Selected Selection
End Sub

' Puts add-in files (.xlam / .xla) that are currently showing as normal
' workbooks back into add-in mode. Does not rely on module-level state, so it
' still works after the VBA project has been reset.
Public Sub Rehide_Addins()
    Dim wb As Workbook
    Dim toHide As Collection
    Dim item As Variant
    Dim wasSaved As Boolean
    Dim ext As String
    Dim okList As String, failList As String
    Dim msg As String

    ' Collect first: changing IsAddin removes the book from Application.Workbooks,
    ' which would disturb a For Each over that same collection.
    Set toHide = New Collection
    For Each wb In Application.Workbooks
        ext = LCase$(Mid$(wb.name, InStrRev(wb.name, ".") + 1))
        If (ext = "xlam" Or ext = "xla") And Not wb.IsAddin Then toHide.Add wb
    Next wb

    If toHide.count = 0 Then
        MsgBox "No add-in workbooks are currently showing.", vbInformation, MSG_TITLE
        Exit Sub
    End If

    For Each item In toHide
        Set wb = item
        wasSaved = wb.Saved
        On Error Resume Next
        Err.Clear
        wb.IsAddin = True
        If Err.Number = 0 Then
            wb.Saved = wasSaved                  ' don't trigger a save prompt
            okList = okList & vbNewLine & "  " & wb.name
        Else
            failList = failList & vbNewLine & "  " & wb.name & " - " & Err.Description
        End If
        On Error GoTo 0
    Next item

    Set mShownAddins = Nothing

    If Len(okList) > 0 Then msg = "Returned to add-in mode:" & okList
    If Len(failList) > 0 Then
        If Len(msg) > 0 Then msg = msg & vbNewLine & vbNewLine
        msg = msg & "Could not return to add-in mode:" & failList
        MsgBox msg, vbExclamation, MSG_TITLE
    Else
        MsgBox msg, vbInformation, MSG_TITLE
    End If
End Sub

'==============================================================================
' Go_To_Selected
'
' Usage:   Go_To_Selected Selection
'
' 1. Takes a Range (normally the user's Selection). Its first cell must
'    contain a formula with QBLOOKUP( ... ) in it.
' 2. QBLOOKUP(sheetName, rowHeader, colHeader, ...)
'       arg 1 = sheet name          (text, cell reference or formula)
'       arg 2 = row header value    (text, cell reference or formula)
'       arg 3 = column header value (text, cell reference or formula)
'    Each argument is evaluated in the context of the formula's own sheet.
' 3. The sheet is looked up in: the formula's workbook, the active workbook,
'    all open workbooks, then all open add-ins (.xlam / .xla).
'    "[Book.xlsx]Sheet1" form is also accepted to name the workbook.
' 4. The row header and column header are found on that sheet, and the cell
'    where they intersect is selected. Header values may use Excel wildcards
'    (* ? and ~ to escape), e.g. "1101 *". Only the normal layout counts:
'    column header ABOVE the cell and row header to the LEFT of it.
' 5. The workbook is brought to the front with that cell selected. If the
'    sheet is in an add-in (or a hidden workbook), the user is asked first;
'    run Rehide_Addins later to put add-ins back into add-in mode.
'
' All problems are reported to the user with a message box.
'==============================================================================

Public Sub Go_To_Selected(ByVal sel As Range)
    Dim srcCell As Range
    Dim srcWs As Worksheet
    Dim ws As Worksheet
    Dim target As Range
    Dim args As Collection
    Dim sheetSpec As Variant
    Dim rowKey As Variant
    Dim colKey As Variant
    Dim srcAddr As String

    On Error GoTo ErrHandler

    '--- 1. Validate the input range -----------------------------------------
    If sel Is Nothing Then RaiseUser "No cell was passed in. Select a cell that contains a QBLOOKUP() formula."

    Set srcCell = sel.Cells(1, 1)
    Set srcWs = srcCell.Worksheet
    srcAddr = "[" & srcWs.Parent.name & "]'" & srcWs.name & "'!" & srcCell.Address(False, False)

    If Not srcCell.HasFormula Then
        RaiseUser "Cell " & srcAddr & " does not contain a formula." & vbNewLine & _
                  "Select a cell that contains a QBLOOKUP() formula."
    End If

    '--- 2. Extract and evaluate the arguments -------------------------------
    Set args = GetQBLookupArgs(CStr(srcCell.Formula), srcAddr)
    If args.count < 3 Then
        RaiseUser "The QBLOOKUP() formula in " & srcAddr & " has only " & args.count & _
                  " argument(s); three are required (sheet name, row header, column header)." & _
                  vbNewLine & vbNewLine & "Formula: " & srcCell.Formula
    End If

    sheetSpec = EvaluateArg(srcWs, CStr(args(1)), 1, "sheet name", srcAddr)
    rowKey = EvaluateArg(srcWs, CStr(args(2)), 2, "row header", srcAddr)
    colKey = EvaluateArg(srcWs, CStr(args(3)), 3, "column header", srcAddr)

    '--- 3. Locate the target sheet (workbooks and add-ins) ------------------
    Set ws = ResolveSheet(CStr(sheetSpec), srcWs.Parent, CStr(args(1)))

    If ws.Visible <> xlSheetVisible Then
        RaiseUser "Sheet '" & ws.name & "' in '" & ws.Parent.name & "' is hidden." & _
                  vbNewLine & "Unhide it and try again."
    End If

    '--- 4. Find the intersection --------------------------------------------
    Set target = FindIntersection(ws, rowKey, colKey, CStr(args(2)), CStr(args(3)))

    '--- 5. Show it -----------------------------------------------------------
    ShowCell target

    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    If Err.Number = USER_ERR Then
        MsgBox Err.Description, vbExclamation, MSG_TITLE
    Else
        MsgBox "Unexpected error " & Err.Number & ":" & vbNewLine & Err.Description, _
               vbCritical, MSG_TITLE
    End If
End Sub


'==============================================================================
' Workbook / sheet lookup
'==============================================================================

' spec may be "Sheet1", "'Sheet 1'", or "[Book.xlsx]Sheet1".
Private Function ResolveSheet(ByVal spec As String, ByVal homeWb As Workbook, _
                              ByVal argText As String) As Worksheet
    Dim wbName As String, shName As String
    Dim p As Long
    Dim wb As Variant
    Dim books As Collection
    Dim searched As String

    spec = Trim$(spec)
    If Len(spec) >= 2 And Left$(spec, 1) = "'" And Right$(spec, 1) = "'" Then
        spec = Replace(Mid$(spec, 2, Len(spec) - 2), "''", "'")
    End If

    shName = spec
    If Left$(spec, 1) = "[" Then
        p = InStr(spec, "]")
        If p > 2 Then
            wbName = Mid$(spec, 2, p - 2)
            shName = Mid$(spec, p + 1)
        End If
    End If
    shName = Trim$(shName)

    If Len(shName) = 0 Then
        RaiseUser "The sheet name argument (" & argText & ") evaluates to """ & spec & _
                  """, which does not contain a sheet name."
    End If

    Set books = CandidateWorkbooks(homeWb)

    If Len(wbName) > 0 Then
        ' A specific workbook was named
        For Each wb In books
            If WorkbookNameMatches(wb, wbName) Then
                Set ResolveSheet = SheetInBook(wb, shName)
                If ResolveSheet Is Nothing Then
                    RaiseUser "Workbook '" & wb.name & "' is open, but it has no worksheet named '" & _
                              shName & "'." & vbNewLine & "(from QBLOOKUP argument 1: " & argText & ")"
                End If
                Exit Function
            End If
        Next wb
        RaiseUser "Workbook '" & wbName & "' is not open (checked workbooks and add-ins)." & _
                  vbNewLine & "(from QBLOOKUP argument 1: " & argText & ")"
    End If

    ' Search every candidate workbook / add-in in priority order
    For Each wb In books
        Set ResolveSheet = SheetInBook(wb, shName)
        If Not ResolveSheet Is Nothing Then Exit Function
        searched = searched & vbNewLine & "  " & wb.name & IIf(wb.IsAddin, "  (add-in)", "")
    Next wb

    RaiseUser "Could not find a worksheet named '" & shName & "'" & vbNewLine & _
              "(from QBLOOKUP argument 1: " & argText & ")" & vbNewLine & vbNewLine & _
              "Searched:" & searched
End Function

' Priority: formula's workbook, active workbook, ThisWorkbook, open workbooks,
' then open add-ins. No duplicates.
Private Function CandidateWorkbooks(ByVal homeWb As Workbook) As Collection
    Dim c As Collection
    Dim wb As Workbook
    Dim ai As Object

    Set c = New Collection
    AddBook c, homeWb
    AddBook c, ActiveWorkbook
    AddBook c, ThisWorkbook

    For Each wb In Application.Workbooks
        AddBook c, wb
    Next wb

    ' Open add-ins are not part of Application.Workbooks
    On Error Resume Next
    For Each ai In Application.AddIns2          ' Excel 2010+
        If ai.IsOpen Then
            Set wb = Nothing
            Set wb = Workbooks(ai.name)
            If Not wb Is Nothing Then AddBook c, wb
        End If
    Next ai
    If Err.Number <> 0 Then                      ' older Excel: fall back to AddIns
        Err.Clear
        For Each ai In Application.AddIns
            If ai.Installed Then
                Set wb = Nothing
                Set wb = Workbooks(ai.name)
                If Not wb Is Nothing Then AddBook c, wb
            End If
        Next ai
    End If
    On Error GoTo 0

    Set CandidateWorkbooks = c
End Function

Private Sub AddBook(ByVal c As Collection, ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    c.Add wb, LCase$(wb.name)                    ' key prevents duplicates
    On Error GoTo 0
End Sub

' Matches "Book.xlsx" or "Book" (without extension).
Private Function WorkbookNameMatches(ByVal wb As Workbook, ByVal nm As String) As Boolean
    Dim base As String
    Dim p As Long

    If StrComp(wb.name, nm, vbTextCompare) = 0 Then
        WorkbookNameMatches = True
        Exit Function
    End If
    p = InStrRev(wb.name, ".")
    If p > 0 Then base = Left$(wb.name, p - 1) Else base = wb.name
    WorkbookNameMatches = (StrComp(base, nm, vbTextCompare) = 0)
End Function

Private Function SheetInBook(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.name, sheetName, vbTextCompare) = 0 Then
            Set SheetInBook = ws
            Exit Function
        End If
    Next ws
End Function

Private Function FirstVisibleWindow(ByVal wb As Workbook) As Window
    Dim win As Window
    For Each win In wb.Windows
        If win.Visible Then
            Set FirstVisibleWindow = win
            Exit Function
        End If
    Next win
End Function

'==============================================================================
' Formula parsing
'==============================================================================

' Returns the top-level arguments of the first QBLOOKUP( call in formula f.
' Respects "string literals", 'quoted sheet names', nested (), {} and [].
Private Function GetQBLookupArgs(ByVal f As String, ByVal srcAddr As String) As Collection
    Const FN As String = "QBLOOKUP("
    Dim args As Collection
    Dim i As Long, n As Long, startPos As Long, depth As Long
    Dim ch As String, prevCh As String, cur As String
    Dim inDq As Boolean, inSq As Boolean, closed As Boolean

    n = Len(f)

    ' Locate the function name outside of string literals / sheet names
    For i = 1 To n
        ch = Mid$(f, i, 1)
        If ch = """" And Not inSq Then
            inDq = Not inDq
        ElseIf ch = "'" And Not inDq Then
            inSq = Not inSq
        ElseIf Not inDq And Not inSq Then
            If UCase$(Mid$(f, i, Len(FN))) = FN Then
                If i = 1 Then prevCh = "" Else prevCh = Mid$(f, i - 1, 1)
                If Not IsNameChar(prevCh) Then     ' avoid e.g. MYQBLOOKUP(
                    startPos = i + Len(FN)
                    Exit For
                End If
            End If
        End If
    Next i

    If startPos = 0 Then
        RaiseUser "The formula in " & srcAddr & " does not contain QBLOOKUP()." & _
                  vbNewLine & vbNewLine & "Formula: " & f
    End If

    ' Split the arguments at top-level commas
    Set args = New Collection
    inDq = False: inSq = False: depth = 0: cur = ""

    For i = startPos To n
        ch = Mid$(f, i, 1)
        If inDq Then
            cur = cur & ch
            If ch = """" Then inDq = False          ' "" escapes re-open on next char
        ElseIf inSq Then
            cur = cur & ch
            If ch = "'" Then inSq = False
        Else
            Select Case ch
                Case """"
                    inDq = True: cur = cur & ch
                Case "'"
                    inSq = True: cur = cur & ch
                Case "(", "{", "["
                    depth = depth + 1: cur = cur & ch
                Case ")", "}", "]"
                    If depth = 0 And ch = ")" Then
                        args.Add Trim$(cur)
                        closed = True
                        Exit For
                    End If
                    depth = depth - 1: cur = cur & ch
                Case ","
                    If depth = 0 Then
                        args.Add Trim$(cur)
                        cur = ""
                    Else
                        cur = cur & ch
                    End If
                Case Else
                    cur = cur & ch
            End Select
        End If
    Next i

    If Not closed Then
        RaiseUser "Could not parse the QBLOOKUP() arguments in " & srcAddr & _
                  " (unbalanced parentheses or quotes)." & vbNewLine & vbNewLine & _
                  "Formula: " & f
    End If

    Set GetQBLookupArgs = args
End Function

Private Function IsNameChar(ByVal c As String) As Boolean
    If Len(c) = 0 Then Exit Function
    IsNameChar = (c Like "[A-Za-z0-9_.]")
End Function

' Evaluates one argument in the context of the formula's sheet and returns a
' single scalar value.
Private Function EvaluateArg(ByVal ctxWs As Worksheet, ByVal argText As String, _
                             ByVal argNum As Long, ByVal argRole As String, _
                             ByVal srcAddr As String) As Variant
    Dim v As Variant
    Dim label As String
    Dim errDesc As String

    argText = Trim$(argText)
    Do While Left$(argText, 1) = "@"                ' strip implicit-intersection operator
        argText = Trim$(Mid$(argText, 2))
    Loop

    label = "QBLOOKUP argument " & argNum & " (" & argRole & ") in " & srcAddr

    If Len(argText) = 0 Then RaiseUser label & " is empty."
    If Len(argText) > 255 Then
        RaiseUser label & " is too long to evaluate (over 255 characters):" & _
                  vbNewLine & argText
    End If

    On Error Resume Next
    v = ctxWs.Evaluate(argText)       ' a Range result returns its value
    If Err.Number <> 0 Then
        errDesc = Err.Description
        On Error GoTo 0
        RaiseUser label & " could not be evaluated:" & vbNewLine & argText & _
                  vbNewLine & vbNewLine & errDesc
    End If
    On Error GoTo 0

    If IsArray(v) Then v = FirstElement(v)

    If IsError(v) Then
        RaiseUser label & vbNewLine & "  " & argText & vbNewLine & _
                  "evaluates to the error " & ErrText(v) & "."
    End If
    If IsEmpty(v) Then
        RaiseUser label & vbNewLine & "  " & argText & vbNewLine & "evaluates to a blank value."
    End If
    If Len(Trim$(CStr(v))) = 0 Then
        RaiseUser label & vbNewLine & "  " & argText & vbNewLine & "evaluates to a blank value."
    End If

    EvaluateArg = v
End Function

Private Function FirstElement(ByVal arr As Variant) As Variant
    On Error Resume Next
    FirstElement = arr(LBound(arr, 1), LBound(arr, 2))
    If Err.Number <> 0 Then
        Err.Clear
        FirstElement = arr(LBound(arr, 1))
    End If
End Function

Private Function ErrText(ByVal v As Variant) As String
    Select Case True
        Case v = CVErr(xlErrNA):    ErrText = "#N/A"
        Case v = CVErr(xlErrRef):   ErrText = "#REF!"
        Case v = CVErr(xlErrName):  ErrText = "#NAME?"
        Case v = CVErr(xlErrValue): ErrText = "#VALUE!"
        Case v = CVErr(xlErrDiv0):  ErrText = "#DIV/0!"
        Case v = CVErr(xlErrNum):   ErrText = "#NUM!"
        Case v = CVErr(xlErrNull):  ErrText = "#NULL!"
        Case Else:                  ErrText = "(Excel error)"
    End Select
End Function

'==============================================================================
' Searching
'==============================================================================

' Finds every cell matching rowKey and colKey and returns the first
' intersection where the column header is ABOVE the cell and the row header
' is to the LEFT of it. Any other arrangement counts as not found.
Private Function FindIntersection(ByVal ws As Worksheet, ByVal rowKey As Variant, _
                                  ByVal colKey As Variant, ByVal rowArgText As String, _
                                  ByVal colArgText As String) As Range
    Dim ur As Range
    Dim data As Variant
    Dim r As Long, c As Long, r0 As Long, c0 As Long
    Dim rowHits As Collection, colHits As Collection
    Dim rh As Variant, ch As Variant
    Dim where As String

    where = "sheet '" & ws.name & "' in '" & ws.Parent.name & "'"

    Set ur = ws.UsedRange
    r0 = ur.row
    c0 = ur.Column

    If ur.CountLarge = 1 Then
        ReDim data(1 To 1, 1 To 1)
        data(1, 1) = ur.Value
    Else
        data = ur.Value
    End If

    Dim rowM As Variant, colM As Variant
    rowM = BuildMatcher(rowKey)
    colM = BuildMatcher(colKey)

    Set rowHits = New Collection
    Set colHits = New Collection

    For r = 1 To UBound(data, 1)
        For c = 1 To UBound(data, 2)
            If ValuesMatch(data(r, c), rowM) Then rowHits.Add Array(r + r0 - 1, c + c0 - 1)
            If ValuesMatch(data(r, c), colM) Then colHits.Add Array(r + r0 - 1, c + c0 - 1)
        Next c
    Next r

    If rowHits.count = 0 And colHits.count = 0 Then
        RaiseUser "Neither header was found on " & where & ":" & vbNewLine & _
                  "  Row header " & FormatKey(rowKey) & "  (from " & rowArgText & ")" & vbNewLine & _
                  "  Column header " & FormatKey(colKey) & "  (from " & colArgText & ")"
    ElseIf rowHits.count = 0 Then
        RaiseUser "Row header " & FormatKey(rowKey) & " (from " & rowArgText & _
                  ") was not found on " & where & "."
    ElseIf colHits.count = 0 Then
        RaiseUser "Column header " & FormatKey(colKey) & " (from " & colArgText & _
                  ") was not found on " & where & "."
    End If

    For Each rh In rowHits
        For Each ch In colHits
            ' column header above (smaller row), row header to the left (smaller column)
            If ch(0) < rh(0) And rh(1) < ch(1) Then
                Set FindIntersection = ws.Cells(rh(0), ch(1))
                Exit Function
            End If
        Next ch
    Next rh

    RaiseUser "Both headers were found on " & where & ", but not in a layout that " & _
              "forms an intersection (column header above, row header to the left):" & vbNewLine & _
              "  Row header " & FormatKey(rowKey) & " - " & rowHits.count & " match(es), first at " & _
              HitAddress(rowHits(1)) & vbNewLine & _
              "  Column header " & FormatKey(colKey) & " - " & colHits.count & " match(es), first at " & _
              HitAddress(colHits(1))
End Function

Private Function HitAddress(ByVal hit As Variant) As String
    HitAddress = Application.ConvertFormula("R" & hit(0) & "C" & hit(1), xlR1C1, xlA1, xlRelative)
End Function

' Pre-computes how a header key is matched, so the pattern is built once.
' Returns Array(isNumber, numberValue, usesWildcards, likePattern, plainText)
' Excel wildcard rules (same as MATCH / VLOOKUP / COUNTIF):
'   *  = any run of characters     ?  = any single character
'   ~* ~? ~~ = a literal *, ? or ~
Private Function BuildMatcher(ByVal key As Variant) As Variant
    Dim isNum As Boolean, numVal As Double
    Dim txt As String, pat As String, ch As String, nxt As String
    Dim i As Long
    Dim wild As Boolean

    isNum = IsNumberType(key)
    If isNum Then numVal = CDbl(key)

    txt = CStr(key)
    wild = (InStr(txt, "*") > 0 Or InStr(txt, "?") > 0 Or InStr(txt, "~") > 0)

    If wild And Not isNum Then
        ' Convert the Excel wildcard pattern into a VBA Like pattern
        i = 1
        Do While i <= Len(txt)
            ch = Mid$(txt, i, 1)
            If ch = "~" And i < Len(txt) Then
                nxt = Mid$(txt, i + 1, 1)
                Select Case nxt
                    Case "*": pat = pat & "[*]": i = i + 1
                    Case "?": pat = pat & "[?]": i = i + 1
                    Case "~": pat = pat & "~": i = i + 1
                    Case Else: pat = pat & "~"
                End Select
            Else
                Select Case ch
                    Case "*": pat = pat & "*"
                    Case "?": pat = pat & "?"
                    Case "#": pat = pat & "[#]"     ' literal in Excel, special in Like
                    Case "[": pat = pat & "[[]"     ' literal in Excel, special in Like
                    Case Else: pat = pat & ch
                End Select
            End If
            i = i + 1
        Loop
        pat = LCase$(pat)
    Else
        wild = False
    End If

    BuildMatcher = Array(isNum, numVal, wild, pat, Trim$(txt))
End Function

' Numbers/dates compare numerically. Text keys with * ? ~ use Excel wildcard
' matching; other text compares as trimmed, case-insensitive text.
Private Function ValuesMatch(ByVal cellVal As Variant, ByVal m As Variant) As Boolean
    If IsError(cellVal) Then Exit Function
    If IsEmpty(cellVal) Then Exit Function

    If m(0) And IsNumberType(cellVal) Then                       ' number vs number
        ValuesMatch = (Abs(CDbl(cellVal) - m(1)) <= 0.000000001 * _
                       Application.WorksheetFunction.Max(1, Abs(m(1))))
    ElseIf m(2) Then                                            ' wildcard pattern
        ValuesMatch = (LCase$(Trim$(CStr(cellVal))) Like m(3))
    Else                                                        ' plain text
        ValuesMatch = (StrComp(Trim$(CStr(cellVal)), m(4), vbTextCompare) = 0)
    End If
End Function

Private Function IsNumberType(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate, vbDecimal, vbByte, 20 '20 = LongLong
            IsNumberType = True
    End Select
End Function

Private Function FormatKey(ByVal v As Variant) As String
    FormatKey = """" & CStr(v) & """"
End Function

'==============================================================================
' Display
'==============================================================================

Private Sub ShowCell(ByVal target As Range)
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim win As Window
    Dim addr As String
    Dim errDesc As String
    Dim wasSaved As Boolean

    Set ws = target.Worksheet
    Set wb = ws.Parent
    addr = "[" & wb.name & "]'" & ws.name & "'!" & target.Address(False, False)

    '--- Add-in workbooks have no visible window until IsAddin = False -------
    If wb.IsAddin Then
        If MsgBox("The target cell " & addr & " is in the add-in '" & wb.name & "'." & _
                  vbNewLine & vbNewLine & _
                  "Show the add-in as a normal workbook so the cell can be selected?" & _
                  vbNewLine & "(Run Rehide_Addins afterwards to put it back.)", _
                  vbQuestion + vbYesNo, MSG_TITLE) = vbNo Then Exit Sub

        wasSaved = wb.Saved
        wb.IsAddin = False
        wb.Saved = wasSaved                          ' don't trigger a save prompt
        If mShownAddins Is Nothing Then Set mShownAddins = New Collection
        On Error Resume Next
        mShownAddins.Add wb.name, LCase$(wb.name)
        On Error GoTo 0
    End If

    '--- Hidden workbook window (e.g. PERSONAL.XLSB) --------------------------
    Set win = FirstVisibleWindow(wb)
    If win Is Nothing Then
        If wb.Windows.count = 0 Then
            RaiseUser "Found " & addr & ", but workbook '" & wb.name & "' has no window to show it in."
        End If
        If MsgBox("Workbook '" & wb.name & "' is hidden." & vbNewLine & _
                  "Unhide it to show " & addr & "?", vbQuestion + vbYesNo, MSG_TITLE) = vbNo Then Exit Sub
        wasSaved = wb.Saved
        wb.Windows(1).Visible = True
        wb.Saved = wasSaved
        Set win = wb.Windows(1)
    End If

    Application.ScreenUpdating = True
    If Application.WindowState = xlMinimized Then Application.WindowState = xlNormal
    win.Activate
    If win.WindowState = xlMinimized Then win.WindowState = xlNormal
    ws.Activate

    On Error Resume Next
    Application.GoTo target, False
    If Err.Number <> 0 Then
        errDesc = Err.Description
        On Error GoTo 0
        If ws.ProtectContents Then
            RaiseUser "Found " & addr & ", but it could not be selected because the " & _
                      "sheet is protected." & vbNewLine & "Unprotect the sheet and try again."
        Else
            RaiseUser "Found " & addr & ", but it could not be selected:" & vbNewLine & errDesc
        End If
    End If

    ' Make sure the cell is on screen
    If Intersect(ActiveWindow.VisibleRange, target) Is Nothing Then
        ActiveWindow.ScrollRow = Application.WorksheetFunction.Max(1, target.row - 3)
        ActiveWindow.ScrollColumn = Application.WorksheetFunction.Max(1, target.Column - 3)
    End If

    ' Bring Excel to the foreground (best effort)
    Err.Clear
    AppActivate ActiveWindow.Caption
    If Err.Number <> 0 Then
        Err.Clear
        AppActivate Application.Caption
    End If
    On Error GoTo 0

    If target.EntireRow.Hidden Or target.EntireColumn.Hidden Then
        MsgBox "Selected " & addr & ", but its row or column is hidden or filtered out." & _
               vbNewLine & "Unhide or clear filters to see it.", vbInformation, MSG_TITLE
    End If
End Sub

Private Sub RaiseUser(ByVal msg As String)
    Err.Raise USER_ERR, MSG_TITLE, msg
End Sub
