Attribute VB_Name = "Import_Export_IIF_Format"
Option Explicit
Global ORIGINALSH_GLOBAL As String  ' For use with User Forms
Global SELECTEDSH_GLOBAL As String  ' For use with User Forms




Sub Import_Vendor_IIF()
Dim m As String '
    Dim origWB As Workbook, wb As Workbook, answer As Long
    
    Application.Calculation = xlManual ' Turn OFF Calculation
    Application.EnableEvents = False   ' Turn OFF
    
    Set origWB = ActiveWorkbook
    Set wb = myOpenWorkbook(True, "*.IIF", origWB)  ' Open File Dialog to copy from ( TRUE for Read Only)
    If wb Is Nothing Then Exit Sub
    
    ' Copy the first sheet of the source workbook
    ' and place it before the last sheet in the target workbook
    ActiveWorkbook.Sheets(1).Copy before:=origWB.Sheets(1)
    
    m = "Activeworkbook is '" & ActiveWorkbook.name & "'" & vbCrLf
    m = m & "OPEN_WB is '" & OPEN_WB.name & "'" & vbCrLf
    m = m & "origWB is '" & origWB.name & "'"
    Debug.Print m
    
    
    
    FormatSheet_Vendor_List ActiveSheet.name
    
    
abortHere:
    Application.Calculation = xlAutomatic ' Turn ON Calculation
    Application.EnableEvents = True ' Turn Events back on
End Sub


Private Sub FormatSheet_Vendor_List(copy_SH As String) ' Called from above
Dim last_row As Long, i As Long, cell As Range, m As String
Dim answer As Long, newShName As String, template_SH  As String, delete_template As Boolean
     If Not SheetExists(copy_SH) Then
        m = "FormatSheet_Vendor_List() Error: Sheet ' & copy_SH & " ' does not exist... Aborting."
        Debug.Print m
        MsgBox m
     End If
    
    Application.Calculation = xlManual ' Turn OFF Calculation
    Application.EnableEvents = False   ' Turn OFF
    Application.ScreenUpdating = False ' Turn OFf
    
    last_row = Last(1, Sheets(copy_SH).Cells)
    
    newShName = "Vendor"
    
' Remove Unnecessary Rows
    For i = last_row To 1 Step -1  ' Step backwards from bottom
        Set cell = Cells(i, 1)
        If cell <> "VEND" And cell <> "!VEND" Then
            Rows(cell.row & ":" & cell.row).Delete ' Delete Row
        End If
    Next
    If Last(1, Sheets(copy_SH).Cells) = 0 Then
        m = "Aborting!" & vbCrLf
        m = m & " There were no VEND rows in IIF worksheet " & copy_SH
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
        template_SH = "Vendor"
        m = "FormatSheet_Vendor_List(): sheet name = '" & newShName & "' does not exist." & vbCrLf
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
   
' Copy and Paste Row 1 Columns J-J (Header Row) ' J' and on are my added Columns
    Sheets(template_SH).Range("J1:J1").Copy Sheets(copy_SH).Range("J1:J1")
' Copy and Paste Row 2 Columns J-J (Data Rows)  'J ' on are  my added formulas
    last_row = Last(1, Sheets(copy_SH).Cells)
    Sheets(template_SH).Range("J2:J2").Copy Sheets(copy_SH).Range("J2:J" & last_row)
' Format Columns ( column width, etc )
    Sheets(template_SH).Columns("A:J").Copy
    Sheets(copy_SH).Activate
    Sheets(copy_SH).Columns("A:J").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
' Format Rows 2 and on
    last_row = Last(1, Sheets(copy_SH).Cells)
    Sheets(template_SH).Range("J2:J2").Copy
    Sheets(copy_SH).Range("J2:J" & last_row).Select ' Format all Rows 2 and on
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
  
' Change Tab Color
    Call changeTabColor(copy_SH, xlThemeColorLight2) ' Change tab to Navy Blue
' Copy Print Format Setings
    Call copyPrintSettings(template_SH, copy_SH)
    Call copyPrintTitleRows(template_SH, copy_SH)
' Freeze View of Row 1 (header Row)
    Call FreezeTopRow(copy_SH)
    'MsgBox "About to rename Sheet... " & copy_SH
    If SheetExists("Old " & newShName, OPEN_WB) Then
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

Private Sub test_LookUp_ClassType()
Dim answer As String
    answer = LookUp_ClassType(ActiveSheet.Range("H2"), ActiveWorkbook.Sheets("Chart of Accounts").Range("H:T"), "BBM-OP")
    MsgBox "answer = " & answer
End Sub
Private Sub test_LookUp_ByIndex()
' Tested both LookUp_ByIndex and LookUp_ClassType. The Time differences were Imperceptible
Dim answer As String
    answer = LookUp_ByIndex(ActiveSheet.Range("H2"), ActiveWorkbook.Sheets("Chart of Accounts").Range("H:T"), 11)
    MsgBox "answer = " & answer
End Sub

'
' WildCardSearch is MUCH Slower than VLookup_
' Use VLookup_ instead
Private Sub test_WildCardSearch()
Dim searchRange2 As Range, what As String
    Set searchRange2 = ActiveWorkbook.Sheets("Chart of Accounts").Range("P:P") ' Account Names
    what = "*BBM Chestnut"
    'what = "13105*"
    ActiveWorkbook.Sheets("Chart of Accounts").Activate
    searchRange2.Select
    MsgBox ("WildCardSearch " & what & " " & WildCardSearch(what, searchRange2))
End Sub

' Search using LIKE
Function WildCardSearch(what As Variant, searchRange As Range, Optional index As Long = 1) As String
Dim cell As Range, i As Long
    If searchRange Is Nothing Then Exit Function
    i = index - 1
    For Each cell In searchRange
        If cell Like what Then
            WildCardSearch = cell.Offset(0, i)
        End If
    Next
End Function

Private Sub test_VLookup_vl()
Dim searchRange2 As Range, what As Variant
    Set searchRange2 = ActiveWorkbook.Sheets("Chart of Accounts").Range("P:P") ' Account Names
    'what = "*Dwi*"
    what = "13105*"
    'ActiveWorkbook.Sheets("Chart of Accounts").Activate
    'SearchRange2.Select
    Debug.Print ("VLookup_ " & what & " " & VLookUp_(what, searchRange2, 1))
End Sub
Private Sub test_VLookup_v2()
Dim searchRange2 As Range, what As Variant, i As Long
    Set searchRange2 = Selection ' Test ccHistorical lookup
    i = 3
    what = "cc USPS PO BOXES ONLINE"
    Debug.Print ("VLookup_ '" & what & "'  index = " & i & "   Result = '" & VLookUp_(what, searchRange2, i)) & "'"
End Sub
Private Sub test_VLookup_v4()
Dim searchRange2 As Range, what As Variant
    'what = "13105*"
    what = 10000
    ActiveWorkbook.Sheets("Chart of Accounts").Activate
    Set searchRange2 = ActiveWorkbook.Sheets("Chart of Accounts").Range("H:P")
    searchRange2.Select
    Debug.Print ("VLookup_ " & what & " " & VLookUp_(what, searchRange2, 9, , False))
End Sub

Function VLookUp_(what As Variant, searchRange As Range, index As Long, _
        Optional dbug As Boolean = False, Optional silent As Boolean = False) ' As String
Dim result As Variant, m As String
    If searchRange Is Nothing Then Exit Function
    result = Application.VLookup(what, searchRange, index, 0)
    VLookUp_ = result

    If IsError(result) Then
        If Not silent Then
            If result = CVErr(xlErrNA) Then ' #N/A error
                m = "VLookUp_(): NOT FOUND error Applicaton.VLookup  "
                m = m & "Worksheet '" & searchRange.Worksheet.name
                m = m & "'  what = '" & what & "'   Index = " & index ' Error 2042 is the VBA equivalent of #N/A
                VLookUp_ = ""
            Else
                m = "VLookUp_(): Unexpected error " & Err & ": " & Err.Description
                m = m & "Worksheet '" & searchRange.Worksheet.name
                m = m & "'  what = '" & what & "'   Index = " & index ' Error 2042 is the VBA equivalent of #N/A
            End If
            Debug.Print m
            'MsgBox M
        End If
    End If
End Function

Private Sub test_xl()
Dim searchRange As Range, returnRange As Range, what As Variant
    Set searchRange = ActiveWorkbook.Sheets("BBM COA").Range("P:P") ' Account Names
    Set returnRange = ActiveWorkbook.Sheets("BBM COA").Range("H:H") ' Account Names
    'what = "*Dwi*" ' FYI... Meta characters did NOT work with Xlookup
    what = "1082*"  ' FYI>>> Meta characters did NOT work with Xlookup
    ActiveWorkbook.Sheets("BBM COA").Activate
    returnRange.Select
    Debug.Print ("XLookup_ " & what & " " & XLookUp_(what, searchRange, returnRange))
End Sub

Private Sub test_xl3()
Dim searchRange As Range, returnRange As Range, what As Variant
    what = "1000"
    ActiveWorkbook.Sheets("BBM COA").Activate
    Set searchRange = ActiveWorkbook.Sheets("BBM COA").Range("H:H")
    Set returnRange = ActiveWorkbook.Sheets("BBM COA").Range("P:P")
    searchRange.Select
    Debug.Print ("XLookup_ " & what & " " & XLookUp_(what, searchRange, returnRange))
End Sub
Private Sub test_xl4()
Dim searchRange As Range, returnRange As Range, what As Variant
    what = 1000
    
    ActiveWorkbook.Sheets("BBM COA").Activate
    Set searchRange = ActiveWorkbook.Sheets("BBM COA").Range("H:H")
    Set returnRange = ActiveWorkbook.Sheets("BBM COA").Range("P:P")
    searchRange.Select
    Debug.Print ("XLookup_ " & what & " " & XLookUp_(what, searchRange, returnRange))
End Sub

' Using Meta characters in what did NOT work with Xlookup
Function XLookUp_(what As Variant, lookupRange As Range, returnRange As Range _
                                        , Optional silent As Boolean = False)
Dim result As Variant, m As String
    
    'On Error Resume Next
    result = Application.WorksheetFunction.XLookup(what, lookupRange, returnRange, "No Match")
    'On Error GoTo 0
    If IsError(result) Then
        If Not result = CVErr(xlErrNA) Then
            m = "XLookUp_(): Unexpected error " & Err & ": " & Err.Description
            m = m & "Worksheet '" & lookupRange.Worksheet.name
            m = m & "'  what = '" & what
        End If
        Debug.Print m
        MsgBox m
    End If
    XLookUp_ = result
End Function

Private Sub test_Compare_COAs()
Dim c1 As Worksheet
Dim c2 As Worksheet
    'Set c1 = Sheets("NEW BBM COA")
    'Set c2 = Sheets("COA_YH")
    Set c2 = Sheets("COA_TEMP")
    
    Set c1 = Sheets("COA_YH")
    'Set c2 = Sheets("COA_BBM_DY")
    
    MsgBox "Total Diffs = " & Compare_COAs(c1, c2, True)
End Sub

' Compare two COA Lists....
' ws1 worksheet is the Master
' ws2 is compared for edits  to existing names or new accounts identifed.
' Account Types OASSET, EXINC, and EXEXP are ignored if excl_DY is set
Function Compare_COAs(ws1 As Worksheet, ws2 As Worksheet, excl_DY As Boolean) As Long
Dim accNum_col As Long, accName_col As Long, accType_col As Long
Dim extra_rows As Long, last_row1 As Long, last_row2 As Long
Dim names_diff As Long, i As Long, vlkup_errors  As Long
Dim missacc1 As Long, missacc2 As Long
Dim index As Long, accname As String
Dim searchRange As Range, m As String, M2 As String, ans As Long

    If ws1 Is Nothing Or ws2 Is Nothing Then
        Compare_COAs = -1: Exit Function
    End If
    last_row1 = Last(1, ws1.Cells)
    last_row2 = Last(1, ws2.Cells)
    ' *** Assign Column Numbers ***
    accNum_col = 8 ' H column
    accName_col = 15 ' P column
    accType_col = 5 ' D column
    index = accName_col - accNum_col + 1
    
    ' 1. *** Test if number of rows of each Account Type are the same. ***
    extra_rows = Count_Accounts(ws1, ws2, excl_DY)
    If extra_rows <> 0 Then m = "ABS Extra Rows  = " & extra_rows & "." & vbCrLf
    
    ' 2. *** Test if Account Names have been modified. ***
    Set searchRange = ws2.UsedRange.Range(ws2.Columns(accNum_col), ws2.Columns(accName_col))
    For i = 2 To last_row1
        If ws1.Cells(i, accNum_col) = 0 Then GoTo next_row
        If excl_types(excl_DY, ws1.Cells(i, accType_col)) Then GoTo next_row
        If excl_accnt(excl_DY, ws1, i) Then GoTo next_row
        
        accname = VLookUp_(ws1.Cells(i, accNum_col), searchRange, index)
        If accname = "" Then vlkup_errors = vlkup_errors + 1: GoTo next_row ' Skip missing account no's ?
        
        ' *** Check names that have been modified and print in the immediate window***
        If accname <> ws1.Cells(i, accName_col) Then
            names_diff = names_diff + 1
            M2 = M2 & "Account " & ws1.Cells(i, accNum_col) & " changed to '" _
                & ws1.Cells(i, accName_col) & "' from '" & accname & "'." & vbCrLf
        End If
next_row:
    Next i
    If names_diff > 0 Or vlkup_errors > 0 Then
        Debug.Print M2
        m = m & "Vlookups not found = " & vlkup_errors
        m = m & ".  Account Names That Changed = " & names_diff & "." & vbCrLf
    End If
    
    ' 3. *** Test if PARENT Account Names are missing from ws1. ***
    Set searchRange = ws1.UsedRange.Range(ws1.Columns(accName_col), ws1.Columns(accName_col))
    For i = 2 To last_row1
        If ws2.Cells(i, accName_col) = "" Then GoTo nextParent1
        If ws2.Cells(i, accNum_col) > 0 Then GoTo nextParent1 ' Not a parent with a accNum
        If excl_types(excl_DY, ws2.Cells(i, accType_col)) Then GoTo nextParent1
        If excl_name(excl_DY, ws2, i) Then GoTo nextParent1
        
        ' *** If name is missing from ws1, bump count and print to the immediate window ***
        accname = VLookUp_(ws2.Cells(i, accName_col), searchRange, 1, False, True) ' index = 1 (only 1 col in search range)
        If accname = "" Then
            missacc1 = missacc1 + 1
            Debug.Print "Parent Account '" & ws2.Cells(i, accName_col) & "' is missing from '" & ws1.name & "'."
        End If
nextParent1:
    Next i
    If missacc1 > 0 Then
        m = m & "Parent Accounts missing from '" & ws1.name & "' = " & missacc1 & vbCrLf
    End If
    
    ' 4. *** Test if PARENT Account Names are missing from ws2. ***
    Set searchRange = ws2.UsedRange.Range(ws2.Columns(accName_col), ws2.Columns(accName_col))
    For i = 2 To last_row1
        If ws1.Cells(i, accName_col) = "" Then GoTo nextParent2
        If ws1.Cells(i, accNum_col) > 0 Then GoTo nextParent2 ' Not a parent with a accNum
        If excl_types(excl_DY, ws1.Cells(i, accType_col)) Then GoTo nextParent2
        If excl_name(excl_DY, ws1, i) Then GoTo nextParent2
        
        ' *** If name is missing from ws2, bump count and print to the immediate window ***
        accname = VLookUp_(ws1.Cells(i, accName_col), searchRange, 1, False, True) ' index = 1 (only 1 col in search range)
        If accname = "" Then
            missacc2 = missacc2 + 1 '
            Debug.Print "Parent Account '" & ws1.Cells(i, accName_col) & "' is missing from '" & ws2.name & "'."
        End If
nextParent2:
    Next i
    If missacc2 > 0 Then
        m = m & "Parent Accounts missing from '" & ws2.name & "' = " & missacc2 & vbCrLf
    End If
    
    If m <> "" Then
        Debug.Print m
        MsgBox "'" & ws1.name & "' to '" & ws2.name & "':" & vbCrLf & "'" & ws2.name & "' has " & m
    End If
    Compare_COAs = Abs(extra_rows) + names_diff + vlkup_errors
End Function
' Ai generated_
Private Function Count_Accounts(ws1 As Worksheet, ws2 As Worksheet, excl_DY As Boolean) As Long
Dim i As Long, uniq_types As String, typeArr() As String, cnt As Long
    ' Assign the comma-separated string of ACCNT TYPE column
    uniq_types = GetUniqueCommaSeparatedList(ws1.UsedRange.Range("E:E"))
    ' Use Split to convert comma separated strings into an array
    typeArr = Split(uniq_types, ",")
    ' Loop through each item in the array
    ' LBound(typeArr) gives the lower bound (usually 0)
    ' UBound(typeArr) gives the upper bound (last index)
    For i = LBound(typeArr) To UBound(typeArr)
        'Debug.Print "Type " & i & ": " & typeArr(i)
        cnt = compWs_acctype(ws1, ws2, Trim(typeArr(i)), False, excl_DY)
        Count_Accounts = Count_Accounts + Abs(cnt)
    Next i
End Function

Private Function compWs_acctype(ws1 As Worksheet, ws2 As Worksheet, accType As String, _
        verbose As Boolean, excl_DY As Boolean) As Long
Dim cnt_ws1 As Long, cnt_ws2 As Long
Dim m As String, ans As Long
    cnt_ws1 = cntTypes(ws1, accType, excl_DY)
    cnt_ws2 = cntTypes(ws2, accType, excl_DY)
    compWs_acctype = cnt_ws2 - cnt_ws1
    If compWs_acctype <> 0 Then
        m = "'" & ws1.name & "' has " & -compWs_acctype
        If compWs_acctype < 0 Then
            m = m & " MORE "
        End If
        If compWs_acctype > 0 Then
            m = m & " LESS "
        End If
        m = m & accType & " types than '" & ws2.name & "'."
        If excl_types(excl_DY, accType) Then
            m = m & " BUT, these Are Ignored."
            compWs_acctype = 0
        End If
        Debug.Print m
        If verbose Then MsgBox m
    End If
End Function
Private Function cntTypes(ws As Worksheet, accType As String, excl_DY As Boolean) As Long
Dim i As Long
    ' Column 5  ACCNT TYPE
    For i = 2 To ws.UsedRange.Rows.count
        If ws.Cells(i, 5) = accType Then
        If excl_accnt(excl_DY, ws, i) Then GoTo next_row
        If excl_name(excl_DY, ws, i) Then GoTo next_row
           cntTypes = cntTypes + 1
        End If
next_row:
    Next i
End Function
Private Sub test_GetUniqueCommaSeparatedList()
   MsgBox GetUniqueCommaSeparatedList(ActiveSheet.UsedRange.Range("E:E"))
End Sub
 
 ' AI Generated
Function GetUniqueCommaSeparatedList(inputRange As Range) As String
    Dim cell As Range
    Dim uniqueItems As Object
    Dim outputList As String
    
    ' Create a new Scripting.Dictionary object
    Set uniqueItems = CreateObject("Scripting.Dictionary")
    
    ' Loop through each cell in the input range
    For Each cell In inputRange
        ' Check if the cell is not empty
        If Trim(cell.Value) <> "" Then
            ' Add the cell's value to the dictionary as a key.
            ' If the key already exists, it won't be added again,
            ' ensuring uniqueness.
            On Error Resume Next ' Handle error if key already exists
            uniqueItems.Add Trim(cell.Value), 1 ' Value doesn't matter, just need a key
            On Error GoTo 0 ' Resume normal error handling
        End If
    Next cell
    
    ' Build the comma-separated string from the dictionary keys
    If uniqueItems.count > 0 Then
        outputList = Join(uniqueItems.Keys, ", ")
    Else
        outputList = ""
    End If
    
    ' Return the unique comma-separated list
    GetUniqueCommaSeparatedList = outputList
End Function
Private Function Max(L1 As Long, L2 As Long) As Long
    If L2 > L1 Then
        Max = L2
    Else
        Max = L1
    End If
End Function
Private Function excl_accnt(excl_DY As Boolean, ws As Worksheet, row_num As Long) As Boolean
Const accNum = 8 ' H col
excl_accnt = False
   If excl_DY Then
     Select Case ws.Cells(row_num, accNum)
     Case 10080:    ' DY CIBC Savings Acct
        excl_accnt = True
     Case 10082:    ' DY O2 MM Acct
        excl_accnt = True
     Case 10090:    ' Ruth's Schwab Acnt
        excl_accnt = True
     Case 10201:    ' Schwab Acct #1
        excl_accnt = True
     Case 10202:    ' Schwab Acct #2
        excl_accnt = True
     Case 10203:    ' Schwab Acct #3
        excl_accnt = True
     Case 10204:    ' Schwab Acct #4
        excl_accnt = True
     Case 10205:    ' Schwab Acct #5
        excl_accnt = True
     Case 10080:
        excl_accnt = True
     Case 27000:    ' PH Liability #1
        excl_accnt = True
     End Select
   End If
End Function
Private Function excl_name(excl_DY As Boolean, ws As Worksheet, row_num As Long) As Boolean
excl_name = False
   If excl_DY Then
     Select Case ws.Cells(row_num, 15) ' 15 Column P - Account
        Case "PH LIABILITIES":    '  · PH LIABILITIES"
        excl_name = True
     End Select
   End If
End Function
Private Function excl_types(excl_DY As Boolean, typ As String) As Boolean
excl_types = False
    If excl_DY Then
        Select Case typ
        'Case "BANK":
        'Case "AR":
        'Case "OCASSET":
        'Case "FIXASSET":
        'Case "AP":
        'Case "CCARD":
        'Case "OCLIAB":
        'Case "LTLIAB":
        'Case "EQUITY":
        'Case "INC":
        'Case "EXP":
        Case "OASSET":
            excl_types = True
        Case "EXINC":
            excl_types = True
        Case "EXEXP":
            excl_types = True
        End Select
   End If
End Function

