Attribute VB_Name = "QB_Reports"
Option Explicit

Private Sub test_formatWsDebitsCreditsReport()
    formatWsDebitsCreditsReport ActiveSheet
End Sub

' Format a "Transaction Detail by Account report" output ws from QB.
' Shrink the width, make it fit the display.
Sub formatWsDebitsCreditsReport(ws As Worksheet)
Dim last_col As Long, last_row As Long
Dim j As Long, debitsCol As Long
    'If ws Is Nothing Then MsgBox "ws is nothing.": Exit Sub
    last_row = Last(1, ws.Cells)
    last_col = Last(2, ws.Cells)
    
    ' Minimize any columns that are less then 2.5 to .5. They're blank anyway.
    For j = 1 To last_col
        If ws.Columns(j).ColumnWidth < 2.5 Then
            ws.Columns(j).ColumnWidth = 0.5
        End If
    Next
    
    ' Place " " in blank columm left of "Debit" to prevent bleed over.
    debitsCol = getColumnNumberA("Debit", ws.Cells.Range("1:1"))
    Call addSpaceToRange(Range(ws.Cells(1, debitsCol - 1), ws.Cells(last_row, debitsCol - 1)))
    
    ' Narrow certain Columns to better fit screen
    ws.Columns(4).ColumnWidth = 9 ' Col D = The last (right most) column for account names.
    j = getColumnNumberA("Name", ws.Cells.Range("1:1"))
    ws.Columns(j).ColumnWidth = 15 '
    j = getColumnNumberA("Memo", ws.Cells.Range("1:1"))
    ws.Columns(j).ColumnWidth = 15 '
    j = getColumnNumberA("Class", ws.Cells.Range("1:1"))
    ws.Columns(j).ColumnWidth = 10 '
    j = getColumnNumberA("Split", ws.Cells.Range("1:1"))
    ws.Columns(j).ColumnWidth = 10 '
    
    ' my Format page for printing qb report
    formatPageWidth ws
End Sub

Private Sub formatPageWidth(ws As Worksheet)
    Application.PrintCommunication = False
    With ws.PageSetup
    ' Format margins to Narrow
    ' and fit to one page wide
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True

End Sub


' Places a space in each cell entire range
' Range MUST be all null
' AI Generated
Private Sub addSpaceToRange(r As Range)
Dim cell As Range
    If Not areAllCellsNull(r) Then
        MsgBox "All cells are NOT null. Skipping this Range."
        Exit Sub
    End If
    ' Loop through each cell in the specified range
    For Each cell In r
        ' Assign a single space to the cell's value
        cell.Value = " "
    Next cell
End Sub

' AI generated
Private Function areAllCellsNull(rng As Range) As Boolean
    ' Returns True if all cells in the targetRange are blank, False otherwise.
    areAllCellsNull = (rng.Cells.count = WorksheetFunction.CountBlank(rng))
End Function

'
'AI generated
Function isWB_Open(ByVal workbookName As String) As Boolean
    Dim wb As Workbook
    On Error Resume Next ' Ignore errors if a workbook is not found
    Set wb = Workbooks(workbookName)
    On Error GoTo 0 ' Reset error handling
    If Not wb Is Nothing Then
        isWB_Open = True
    Else
        isWB_Open = False
    End If
End Function

Private Sub tst_displayDebitsandCredits()
Dim ws As Worksheet
    'test_displayDebitsandCredits Selection
    'dDebitsandCredits Selection, ws, "1/1/2024", Now()
    Set ws = Sheets("Sheet1")
    dDebitsandCredits 15118, ws, "1/1/2024", Now()
End Sub
Sub displayDebitsandCredits(rng As Range)
Dim ws As Worksheet
    If getLinkAddress(rng) = "" Then
        MsgBox "Selection should be a hyperlinked cell here.  Exiting..."
    End If
    dDebitsandCredits rng, ws, "1/1/2024", "12/31/2024"
End Sub


' Displays Total Debits and Credits of an account between start and end on dates
'                       in a qb detail transactions worksheet
'   acct = the account number to search
'   ws = worksheet to search QB detail transactions by 'Debit' and 'Credit' columns
'   sDate = start date in 'Date' column for transactions
'   eDate = end date in 'Date' column for transactions
' if acct is a HyperLinked cell then worksheet to search is derived from HyperLink.
'
Sub dDebitsandCredits(ByVal acct, ByRef ws As Worksheet, ByVal sDate As Date, ByVal eDate As Date)
Dim date_col As Long, debit_col As Long, credit_col As Long
Dim dateRange As Range, debitRange As Range, creditRange As Range
Dim srchRange As Range, srchVal As Variant, srchWs As Worksheet
Dim last_row As Long, srow As Long, erow As Long
Dim ttl_debits As Double, ttl_credits As Double
Dim m As String, ans As Long
    
    If TypeOf acct Is Range Then
        ' Check if range is nothing
        If acct Is Nothing Then Exit Sub
    
        ' Check if only one cell is selected to avoid issues with range selections
        If acct.CountLarge > 1 Then _
        MsgBox "More than one cell selected.  Exiting...", , "dDebitsandCredits()": Exit Sub
        
        srchVal = acct.Value
    Else
        srchVal = acct
    End If
    
    If getLinkAddress(acct) = "" And ws Is Nothing Then
        m = "Acct: " & acct & " is not a Hyperlinked cell and no worksheet passed in.  Exiting..."
        MsgBox m, , "dDebitsandCredits()"
        Exit Sub
    End If
    
    If TypeOf acct Is Range And ws Is Nothing Then
        Set srchWs = getWsFromLinkSubAddress(acct)
    Else
        Set srchWs = ws
    End If
    If srchWs Is Nothing Then Exit Sub
    
      
    ' Get Column numbers
    date_col = getColumnNumberA("Date", srchWs.Cells.Range("1:1"))
    debit_col = getColumnNumberA("Debit", srchWs.Cells.Range("1:1"))
    credit_col = getColumnNumberA("Credit", srchWs.Cells.Range("1:1"))
    last_row = Last(1, srchWs.Cells)
    
    If date_col = 0 Or debit_col = 0 Or credit_col = 0 Then Exit Sub
    If last_row < 2 Then Exit Sub
    
    Set srchRange = srchWs.Cells.Range("A:D") ' First four columns
    
    On Error Resume Next
    ' Borrowed this .find code from my Build_Acct_Links macro in Hyper_Link_Mod
    srow = srchRange.Find(what:=srchVal, _
            After:=srchRange.Cells(1), _
            LookAt:=xlPart, _
            LookIn:=xlFormulas, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, _
            MatchCase:=False).row
    erow = srchRange.Find(what:=srchVal, _
            After:=srchRange.Cells(1), _
            LookAt:=xlPart, _
            LookIn:=xlFormulas, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlPrevious, _
            MatchCase:=False).row - 1 ' minus one to skip the total row.
    
    On Error GoTo 0
    If srow = 0 Or erow = 0 Then
        m = "Could not locate account: '" & acct & "'" & vbCrLf
        m = m & "srow = " & srow & "   " & "erow = " & erow
        MsgBox m
        Exit Sub
    End If
    'set sumif ranges
    Set dateRange = Range(srchWs.Cells(srow, date_col), srchWs.Cells(erow, date_col))
    Set debitRange = Range(srchWs.Cells(srow, debit_col), srchWs.Cells(erow, debit_col))
    Set creditRange = Range(srchWs.Cells(srow, credit_col), srchWs.Cells(erow, credit_col))
        
    ttl_debits = Application.WorksheetFunction.SumIfs( _
                    debitRange, _
                    dateRange, ">=" & sDate, _
                    dateRange, "<=" & eDate)

    ttl_credits = Application.WorksheetFunction.SumIfs( _
                    creditRange, _
                    dateRange, ">=" & sDate, _
                    dateRange, "<=" & eDate)

    ' Format the answer to display
    m = "Account = " & acct & vbCrLf & vbCrLf
    m = m & " " & Format(sDate, "MM/DD/YY") & " to " & Format(eDate, "MM/DD/YY") & vbCrLf & vbCrLf
    m = m & "Debits = " & Format(ttl_debits, "###,##0.00") & vbCrLf
    m = m & "Crebits = " & Format(ttl_credits, "###,##0.00")
    ans = MsgBox(m, vbOKOnly, "Totals for ACCT " & acct)
    Exit Sub
error_Out:
    ans = MsgBox(m, , "Error in DisplayCreditsandDebits()")
End Sub

' used by dDebitsandCredits
Private Function getWsFromLinkSubAddress(ByVal rng As Range) As Worksheet
Dim wb As Workbook, ws As Worksheet
Dim FN As String, m As String, ans As Long
    
    m = getLinkAddress(rng)
    If m = "" Then
        MsgBox "Hyperlink is nothing.  Exiting...": Exit Function
    End If
    FN = getFileNameFromMixedPath(m)
    If Not isWB_Open(FN) Then
        m = m & vbCrLf & "Hyperlink file is not open." & vbCrLf & vbCrLf
        m = m & "Open it?"
        ans = MsgBox(m, vbYesNo, "LINK NOT OPEN...")
        If ans <> vbYes Then Exit Function
        ' open the hyperlink workbook with read/write
        Set wb = myOpenWorkbook(True, , , getLinkAddress(rng))
    Else
        Set wb = Workbooks(FN)
    End If
    If wb Is Nothing Then MsgBox "wb is nothing.  Exiting...": Exit Function
    
    m = getLinkSubAddress(rng)
    If m = "" Then
        MsgBox "Hyperlink Subaddress is nothing.  Exiting...": Exit Function
    End If
    ' Get sheet name from sub address
    m = extractTextBetweenSingleQuotes(m)
    If m = "" Then
        MsgBox "Sheet Name in Subaddress is nothing.  Exiting...": Exit Function
    End If
    Set ws = wb.Sheets(m)
    If ws Is Nothing Then
        MsgBox "Sheet Name '" & m & "' is not in wb '" & wb.name & "'.  Exiting...": Exit Function
    End If
        
    Set getWsFromLinkSubAddress = ws
End Function

'AI generated
Private Function extractTextBetweenSingleQuotes(ByVal inputString As String) As String
    Dim firstQuotePos As Long
    Dim secondQuotePos As Long

    ' Find the position of the first single quote
    firstQuotePos = InStr(inputString, "'")

    ' If a first single quote is found
    If firstQuotePos > 0 Then
        ' Find the position of the second single quote, starting after the first one
        secondQuotePos = InStr(firstQuotePos + 1, inputString, "'")

        ' If a second single quote is found
        If secondQuotePos > 0 Then
            ' Extract the substring between the two quotes
            extractTextBetweenSingleQuotes = Mid(inputString, firstQuotePos + 1, secondQuotePos - firstQuotePos - 1)
        Else
            ' No second single quote found
            extractTextBetweenSingleQuotes = ""
        End If
    Else
        ' No first single quote found
        extractTextBetweenSingleQuotes = ""
    End If
End Function

Private Function which_Year() As String
Dim m As String, ans As Long
Dim last_year
    last_year = year(Now()) - 1
    m = "For which YEAR?"
    which_Year = Application.InputBox(m & _
            vbCrLf & "Enter a Year to Summarize.", _
    "Year", last_year, , , , 2)
    
    If which_Year = "False" Then which_Year = "User Canceled.": Exit Function  ' User Canceled
    
    If Not isYearValid(which_Year) Then
            which_Year = " User entered invalid Year '" & which_Year & "'"
    End If
End Function
Private Function isYearValid(yr As Variant) As Boolean
Dim last_year
    last_year = year(Now()) - 1
    isYearValid = False
    
    If Not isStringAllDigits(yr) Then Exit Function
    
    ' check for up to 5 years ago then OK
    If CLng(yr) <= last_year + 1 And CLng(yr) > last_year - 4 Then
        isYearValid = True
    End If
End Function

' AI Generated
Private Function isStringAllDigits(ByVal inputString As String) As Boolean
    Dim i As Long
    
    If Len(inputString) = 0 Then
        isStringAllDigits = False ' An empty string is not all digits
        Exit Function
    End If
    
    For i = 1 To Len(inputString)
        If Not IsNumeric(Mid(inputString, i, 1)) Then
            isStringAllDigits = False
            Exit Function
        End If
    Next i
    
    isStringAllDigits = True
End Function
Private Sub test_GetColumnNumberA()
Dim col_no As Long, col_no2 As Long
    col_no = getColumnNumberA("Amount")
    col_no2 = getColumnNumberA("NAME")
    
    MsgBox "col_no = " & col_no & "  col_no2 = " & col_no2
End Sub


' Returns first column number of text in range
' If no range is passed in then the first row of the active sheet is searched.
' AI generated
Function getColumnNumberA(ByVal searchVal As String, Optional ByVal targetRange As Range = Nothing) As Long
    Dim foundCell As Range
    If targetRange Is Nothing Then
        ' look in Active sheet ??
        Set targetRange = Range("1:1") ' ACTIVE SHEET
    End If
    Set foundCell = targetRange.Find(what:=searchVal, LookIn:=xlValues, LookAt:=xlWhole, _
                                    SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
'    Tried setting the 'After:' param like in GetColumnNumberB and it Still Doesn't work!
'    Set foundCell = targetRange.Find(what:=searchVal, After:=targetRange.Cells(targetRange.Cells.count), LookIn:=xlValues, LookAt:=xlWhole, _
                                    SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
    If Not foundCell Is Nothing Then
        getColumnNumberA = foundCell.Column
    Else
        getColumnNumberA = 0  ' Or any other value to indicate not found
        Debug.Print "GetColumnNumberA() could not find a column named '" & searchVal & "' in '" _
                & targetRange.Parent.Parent.name & "'  '" & targetRange.Parent.name & "'"
    End If
    Set targetRange = Nothing
End Function
Private Sub test2_GetColumnNumberB()
Dim col_no As Long, col_no2 As Long
    col_no = getColumnNumberA("Amount") ' Fails!!
    col_no2 = getColumnNumberB("Amount") ' Works!! ' NOW THIS FAILS TOO! DY 12-21-25
    
    MsgBox "col_no = " & col_no & "  col_no2 = " & col_no2
End Sub


Private Sub test_GetColumnNumberB()
Dim col_no As Long, col_no2 As Long
    col_no = getColumnNumberB("Amount")
    col_no2 = getColumnNumberB("NAME")
    
    MsgBox "col_no = " & col_no & "  col_no2 = " & col_no2
End Sub

' I found this and it fixed the "Amount" col search problem
' Don't know why...
' https://stackoverflow.com/questions/68531246/find-column-number-of-a-specific-text-in-vba-excel-and-use-its-address-in-rang#:~:text=Find%20column%20number%20of%20a%20specific%20text,row%20that%20I%20will%20specify%20%C2%B7%20Intersect(SValue.
Function getColumnNumberB(ByVal searchVal As String, Optional ByVal rg As Range = Nothing) As Long
Dim fCell As Range
    If rg Is Nothing Then
        ' look in Active sheet ??
        Set rg = Range("1:1") ' ACTIVE SHEET
    End If
    With rg
        Set fCell = .Find(searchVal, .Cells(.Cells.count), xlFormulas, xlWhole)
    End With
    If fCell Is Nothing Then
        Debug.Print "GetColumnNumberB() could not find a column named '" _
            & searchVal & "' in '" & rg.Parent.Parent.name & "'  '" & rg.Parent.name & "'"
        Exit Function
    End If
    getColumnNumberB = fCell.Column
End Function

Private Sub test3_GetColumnNumberC()
Dim col_no As Long, col_no2 As Long, col_no3 As Long
    col_no = getColumnNumberA("Amount") ' Fails!!
    col_no2 = getColumnNumberB("Amount") ' Works!! ' NOW THIS FAILS TOO! DY 12-21-25
    col_no3 = getColumnNumberC("Amount") ' Works!! 12-21-25
    
    MsgBox "col_no = " & col_no & "  col_no2 = " & col_no2 & "  col_no3 = " & col_no3
End Sub

' The fall back Linear Search... One by One
' This works...
Function getColumnNumberC(ByVal searchVal As String, Optional ByVal rg As Range = Nothing) As Long
Dim fCell As Range, numRows As Long, numCols As Long, i As Long, j As Long
    
    getColumnNumberC = 0
    If rg Is Nothing Then
        ' look in Active sheet ??
        Set rg = UsedRangeOfRow1
    End If
    
    If rg Is Nothing Then
        Debug.Print "GetColumnNumberC() rg is nothing... Exiting."
        Exit Function
    End If
        
    numRows = rg.Rows.count
    numCols = rg.Columns.count
        
   ' Loop through columns first
    For j = 1 To numCols
        ' Inside the column loop, loop through rows
        For i = 1 To numRows
            ' Get the current cell using Cells property relative to the top-left of the range
            Set fCell = rg.Cells(i, j)
            
            ' Check if the cell's value matches the search value
            If UCase(fCell.Value) = UCase(searchVal) Then
                ' Value found! Set the function's result to the found cell and exit the function
                getColumnNumberC = fCell.Column
                Exit Function
            End If
        Next i
    Next j
    
    If getColumnNumberC = 0 Then
        Debug.Print "GetColumnNumberC() could not find a column named '" _
            & searchVal & "' in '" & rg.Parent.Parent.name & "'  '" & rg.Parent.name & "'"
        Exit Function
    End If
    getColumnNumberC = fCell.Column
End Function

' AI generated
' Return used range of row 1 of ACTIVE SHEET
Function UsedRangeOfRow1() As Range
    
    ' Set the UsedRangeOfRow1 to the intersection of the entire first row and the used range of the sheet
    With ActiveSheet
        Set UsedRangeOfRow1 = Intersect(.Rows(1), .UsedRange)
    End With
    
    ' Check if the intersection resulted in a valid range (e.g., if row 1 is used at all)
    If UsedRangeOfRow1 Is Nothing Then
        MsgBox "UsedRangeOfRow1(): Row 1 is entirely empty."
    End If

End Function

