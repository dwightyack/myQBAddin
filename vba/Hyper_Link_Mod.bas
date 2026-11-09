Attribute VB_Name = "Hyper_Link_Mod"
Option Explicit
Option Compare Text ' Makes any string compare 'case insensitive' in  Module

'
'  Mod contains VBA associated with Hyper Links for K-1's, Qb Accounts, and Other files
'

Public Sub Clear_Cells(rng As Range, Optional prompt As Boolean = False)
' Clear rng Row 2 and On
'
Dim answer As Long
    If rng Is Nothing Then Exit Sub

    Set rng = rng.Rows(2 & ":" & rng.Rows.count) ' Clear 2 and on
    If prompt Then
        rng.Select  ' debug
        answer = MsgBox("Clear ?", vbOKCancel)
        If answer <> vbOK Then Exit Sub
    End If
    rng.Clear ' Clear Contents
    Call ClearFormat(rng) ' Clear Format
End Sub
Private Sub test_CheckFileLinks()
    Call CheckFileLinks(Selection)
End Sub
Sub CheckFileLinks(rng As Range, Optional rebuild As Boolean = True)
Dim cell As Range, addr As String
    On Error Resume Next
    If rng Is Nothing Then Exit Sub
    For Each cell In rng
        If cell.row <= 1 Then GoTo nextOne ' Row one is always a Header
        If cell.row > Last(1, rng) Then Exit For
        Call ClearFormat(cell)
        If cell.Value = "" Then GoTo nextOne
        addr = getLinkAddress(cell)
        ' ".\" not needed for file exists test.
        addr = StripDotBackSlash(addr)
        If InStr(addr, "\") = 0 Then addr = ActiveWorkbook.path & "\" & addr
        If Not myFileExists(addr) Then
            If rebuild Then
                Call Rebuild_Links(cell, rng)
            Else
                Call ErrorFormat(cell)
            End If
        End If
nextOne:
    Next cell
End Sub

Private Sub test_getLinkAddess()
    Debug.Print getLinkAddress(Selection)
End Sub

Function getLinkAddress(myCell) As String  ' Returns the Hyper Link Address
'https://stackoverflow.com/questions/27100422/user-defined-function-to-get-hyperlink-address-from-cell
    On Error Resume Next
    If myCell Is Nothing Then Exit Function
    Application.Volatile  ' Volatile used for UDF's that need recalc EVERY time a calc is done
    getLinkAddress = myCell.Hyperlinks(1).Address
End Function

Private Sub test_getLinkSubAddress()
    Debug.Print getLinkSubAddress(Selection)
End Sub

Function getLinkSubAddress(myCell As Range) As String ' Returns the Hyper Link Address
'https://stackoverflow.com/questions/27100422/user-defined-function-to-get-hyperlink-address-from-cell
    On Error Resume Next
    If myCell Is Nothing Then Exit Function
    Application.Volatile  ' Volatile used for UDF's that need recalc EVERY time a calc is done
    getLinkSubAddress = myCell.Hyperlinks(1).SubAddress
End Function
Sub ClearFormat(rng As Range)
    With rng.Interior
        .pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    rng.Font.Italic = False
End Sub
Private Sub ErrorFormat(rng As Range) ' Fill rng with Error Red
    With rng.Interior
        .Color = 8750591 ' pink-red
    End With
    rng.Font.Italic = True
End Sub
Sub Build_FileTags(rng As Range, put_rng As Range)
' Takes File Names in (rng), strips them and places File Tags in (put_rng) column
Dim cell As Range
Dim tag_col As Long, link_col As Long, lastRow As Long
    On Error Resume Next
    If rng Is Nothing Or put_rng Is Nothing Then Exit Sub
    tag_col = put_rng.Column '  Column for Placing File Tags
    If tag_col < 1 Then Exit Sub ' Invalid Range Name
    ' Take File Name from (rng), strip it and place File Tag in tag_col
    For Each cell In rng
        If cell.row <= 1 Then GoTo nextOne ' Row one is always a Header
        lastRow = Last(1, rng)
        If cell.row > lastRow Then Exit For
        If cell.Value = "" Then GoTo nextOne
        Cells(cell.row, tag_col).Value = FileTag(cell.Value)
nextOne:
    Next cell
End Sub
Private Sub test_FileTag()
    MsgBox FileTag("2018 K-1 AND  off and the .pdf")
End Sub
Private Function FileTag(name As String) As String
' Strips the "2018 K-1 " (9) off and the ".pdf" (4) off
FileTag = Right(Left(name, Len(name) - 4), Len(Left(name, Len(name) - 4)) - 9)
End Function
' Need more testing on Absolute and Relative links....
' It appears from dbgPrt that My K-1 List file links NO LONGER have ANY ./ or ../ !!!
Function StripDotBackSlash(name As String, Optional dbgPrt As Boolean = False) As String
' Strips the ".\" off the front
Dim fso As Object
StripDotBackSlash = name
If Left(name, 2) = ".\" Then
    StripDotBackSlash = Right(name, Len(name) - 2)
    If dbgPrt Then Debug.Print ".\ Convert '" & name & "' to '" & StripDotBackSlash & "'"
ElseIf Left(name, 3) = "..\" Then
    ' Convert Relative path to Absolute path
       Set fso = CreateObject("Scripting.FileSystemObject")
       StripDotBackSlash = fso.GetAbsolutePathName(name)
       If dbgPrt Then Debug.Print "..\ Convert '" & name & "' to '" & StripDotBackSlash & "'"
End If
If dbgPrt Then Debug.Print "No Change...  '" & name & "' to '" & StripDotBackSlash & "'"
End Function

Sub Clear_Acct_Links(rng As Range)
    rng.Hyperlinks.Delete
    rng.HorizontalAlignment = xlCenter ' Need to recenter after Hyperlinks.delete for some reason.
End Sub
Private Sub test_()
    Call SelectWorkbook_Click ' Bring up the Open Workbook List Box"
End Sub
Private Sub test_Build_Acct_Links()
    Call Build_Acct_Links(Selection)
End Sub
Sub Build_Acct_Links(rng As Range)
' Put a Hyperlink on each QB account number cell in range rng
' HyperLink the account number to a row or range of rows is a QB worksheet
Dim acct As Range
Dim srchRange As Range, fndrange As Range
Dim srow As Variant, erow As Variant ' Need variant for .Find
Dim end_col As Long, thisWs As Worksheet, eWs As Worksheet, eWb As Workbook
Dim last_row As Long, found_cnt As Long, notfound_cnt As Long, m As String
Dim qbHeaders As Long
Const modAndFunc = "myQBAddin.Build_Acct_Links"

    rng.Parent.Activate
    rng.Select
    
    MsgBox "You will now choose a QB report workbook and sheet to link the QB account numbers in this range." _
    , , "Choose a workbook and sheet..."

'    ufrmWorkBooksOpen.ufrmWorkbooksOpenCaption = "Choose a Workbook..."
'    Set eWb = SelectWorksheet_Click  ' Bring up the Open Workbook List Box"
'
'    If eWb Is Nothing Then
'        MsgBox "External Workbook is nothing.... Exiting", , modAndFunc
'        Exit Sub
'    End If
'
'    ' MODIFY THIS LATER TO SELECT A WS
'    If eWb.Worksheets.count > 1 Then
'        MsgBox "Warning: " & eWb.name & " has more than one worksheet." _
'            & vbCrLf & "Taking the first one.", , modAndFunc
'    End If
'
'    Set eWs = eWb.Worksheets(1)
    
   Set eWs = SelectWorksheet_Click  ' Bring up the select Workbook and sheet List Box"
   If eWs Is Nothing Then
        MsgBox "External Worksheet is nothing.... Exiting", , modAndFunc
        Exit Sub
    End If
    If rng Is Nothing Then
        MsgBox "rng is nothing.... Exiting", , modAndFunc
        Exit Sub
    End If
    Set thisWs = rng.Parent
       
    'thisWs.Activate
    'srchrange.Select
    'MsgBox "Search here...",,modAndFunc
    
    qbHeaders = qbHowManyRowLevels(eWs)
    If qbHeaders = 0 Then
        Debug.Print "qbHowManyRowLevels could not find anything... Setting qbHeaders = " & qbHeaders
        qbHeaders = 7 'Set default if "QBROWHEADERS" in not a defined Name in eWs
    End If
    On Error Resume Next
    ' Define search range from column 1 (A) to column set by qbHeaders
     Set srchRange = eWs.Range(eWs.Cells(1, 1), eWs.Cells(eWs.Rows.count, qbHeaders))
    If srchRange Is Nothing Then
        MsgBox "srchRange is nothing.... Exiting", , modAndFunc
        Exit Sub
    End If
    
    last_row = Last(1, rng)
    
    ' Build a Hyperlink for each acct number in rng
    For Each acct In rng
        ' Remove any previous link
        acct.Hyperlinks.Delete
        ' Skip Row one is always a Header
        If acct.row <= 1 Then GoTo nextOne
        ' Exit if last of Active cells in the range
        If acct.row > last_row Then Exit For
        ' Skip if Null
        If acct.Value = "" Then GoTo nextOne
        ' acct.value is account number. Search for account number in first 4 columns
        srow = srchRange.Find(what:=acct.Value, _
                After:=srchRange.Cells(1), _
                LookAt:=xlPart, _
                LookIn:=xlFormulas, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlNext, _
                MatchCase:=False).row
        erow = srchRange.Find(what:=acct.Value, _
                After:=srchRange.Cells(1), _
                LookAt:=xlPart, _
                LookIn:=xlFormulas, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious, _
                MatchCase:=False).row
        
        If srow = 0 And erow = 0 Then
            notfound_cnt = notfound_cnt + 1
            GoTo nextOne
        End If
        end_col = Last(2, eWs.Cells)
        
        ' Set the range to the matched account
        Set fndrange = Range(eWs.Cells(srow, 1), eWs.Cells(erow, end_col))
        If fndrange Is Nothing Then
            notfound_cnt = notfound_cnt + 1
            GoTo nextOne
        End If
        'ActiveWorkbook.Worksheets(SheetName).Activate
        'fndrange.Select
        'MsgBox ("Found This for " & acct)
        
        '  OLD CODE .............. for local work sheet
        'ActiveWorkbook.Worksheets(rng.Parent.name).Hyperlinks.Add _
        '    anchor:=acct, _
        '    Address:="", SubAddress:=sheetName & "!A1"
        'acct.Hyperlinks(1).SubAddress = sheetName & "!" & fndrange.AddressLocal
        
        '  NEW CODE .............. for external workbooks
        found_cnt = found_cnt + 1
        thisWs.Hyperlinks.Add _
            Anchor:=acct, _
            Address:=eWs.Parent.fullName, _
            SubAddress:="'" & eWs.name & "'!" & fndrange.AddressLocal
            'ScreenTip:="Click to open external workbook."
        acct.HorizontalAlignment = xlCenter
nextOne:
        srow = 0: erow = 0
    Next acct
    
    m = "Cells Linked = " & found_cnt & vbCrLf
    m = m & "Cells NOT Linked = " & notfound_cnt & vbCrLf
    
    MsgBox m, , modAndFunc
    
End Sub

' Any cells in (rng) marked as "CLOSED", rename the respective File Name linked in put_rng cell
' create a new Hyperlink with the new File Name in respective put_rng cell
' This function works with multiple rows in rng and put_rng
' rng and put_rng are single column ranges and must have the same respective row numbers
Sub Change_EntryWB_Name(rng As Range, put_rng As Range, Optional strPath As String = ".\")
Dim cell As Range, c As Long, old_fn As String, new_fn As String, Crng As Range
Dim fso As New FileSystemObject, fname As String, fpath As String

    On Error Resume Next
    If rng Is Nothing Or put_rng Is Nothing Then Exit Sub
    c = put_rng.Column ' column to find and replace File Name Link
    If c < 1 Then Exit Sub ' Invalid Range Name
    For Each cell In rng
        'If cell.row > Last(1, rng) Then Exit For  ' Not needed?
        If Right(strPath, 1) <> "\" Then Exit Sub  ' Must have '\' and at end
        If cell.row <= 1 Then GoTo nextOne ' Row one is always a Header
        If cell <> "CLOSED" Then GoTo nextOne
        old_fn = getLinkAddress(Cells(cell.row, c)) ' get file name from Hyperlink
        If old_fn = "" Then Exit Sub ' no linked file
        fname = fso.GetFileName(old_fn)
        fpath = fso.GetParentFolderName(old_fn)
        If Left(fname, 6) = "closed" Then GoTo nextOne
        new_fn = strPath & "closed " & fname
        fpath = fpath & "\"
        If fpath <> strPath Then
            Debug.Print ("Change_EntryWB_Name: path was '" & fpath & "'  fname = '" & fname)
            Debug.Print ("Changing to: fn = '" & new_fn & "'")
        End If
        'Name old_fn As new_fn   ' renames a file
        Name old_fn As fpath & "\closed " & fname   ' renames a file
        Cells(cell.row, c).Clear
        ClearFormat (Cells(cell.row, c))
        ActiveSheet.Hyperlinks.Add Cells(cell.row, c), new_fn  ', TextToDisplay:=new_fn
        Cells(cell.row, c).Value = new_fn  ' Not needed if using TextToDisplay:
        'fullPath = strPath  'full path not used any more
        'If strPath = ".\" Then fullPath = ActiveWorkbook.path & "\" ' WAS Needed for FileExists tests on '.\'  before changes
nextOne:
    Next cell
    ' Check the links to see they work
    Set Crng = rng.Offset(0, c - rng.Column)  ' This ALSO Works!
    Call CheckFileLinks(Crng, False) ' Check the File Links in destination Range using FileExists
End Sub
' Takes File Tags from (rng), builds a Hyperlink File Name and places File Name in (put_rng)
' This functins works with multiple rows in rng and put_rng
' rng and put_rng must have the same row numbers
Sub Build_Links_From_FileTags(rng As Range, yearrng As Range, put_rng As Range, Optional strPath As String = ".\")
Dim cell As Range, c As Long, FN As String, fullpath As String, Crng As Range, last_row As Long
    On Error Resume Next
    If rng Is Nothing Or put_rng Is Nothing Then Exit Sub
    c = put_rng.Column ' Place to put the File Name
    If c < 1 Then Exit Sub ' Invalid Range Name
    last_row = Last(1, rng)
    For Each cell In rng
        If cell.row <= 1 Then GoTo nextOne ' Row one is always a Header
        If cell.row > last_row Then Exit For
        FN = getLinkAddress(cell) ' test FileTag is NOT a Hyperlink
        If FN <> "" Then Exit Sub ' Selection was NOT a FileTag
        Cells(cell.row, c).Clear
        ClearFormat (Cells(cell.row, c))
        If cell.Value = "" Then GoTo nextOne
        If yearrng(cell.row, 1) = "Closed" Then
            FN = "CLOSED"
        Else
            FN = yearrng(cell.row, 1) & " K-1 " & cell.Value & ".pdf"
        End If
        
        ActiveSheet.Hyperlinks.Add Cells(cell.row, c), strPath & FN   ', TextToDisplay:=fn
        Cells(cell.row, c).Value = FN  ' Not needed if using TextToDisplay:
        fullpath = strPath
        If strPath = ".\" Then fullpath = ActiveWorkbook.path & "\" ' Needed for FileExists tests on '.\'
nextOne:
    Next cell
    Set Crng = rng.Offset(0, c - rng.Column)  ' This ALSO Works!
    Call CheckFileLinks(Crng, False) ' Check the File Links in destination Range using FileExists
End Sub

' Display Hyperlink properties of ALL links in Active worksheet
' Hyperlink object Properties:
'
'    Address: The target URL, file path, or email address of the hyperlink.
'    SubAddress: Used for linking to a specific location within the target, such as a bookmark in a document or a named range in a workbook.
'    ScreenTip: The text that appears when the user hovers over the hyperlink.
'    TextToDisplay: The text that is visible in the cell or object that contains the hyperlink.
'    Anchor: The range, shape, or control that serves as the clickable element for the hyperlink.
'
Sub DisplayHyperlinkProperties()
Dim ws As Worksheet
Dim hl As Hyperlink
Dim i As Long

    ' Set the worksheet where your hyperlinks are located
    Set ws = ActiveSheet ' Change "Sheet1" to your sheet name

    ' Check if there are any hyperlinks on the sheet
    If ws.Hyperlinks.count = 0 Then
        MsgBox "No hyperlinks found on " & ws.name, vbInformation
        Exit Sub
    End If

    ' Loop through each hyperlink on the worksheet
    For i = 1 To ws.Hyperlinks.count
        Set hl = ws.Hyperlinks(i)

        ' Display the properties
        Debug.Print "--- Hyperlink " & i & " ---"
        Debug.Print "Address: " & hl.Address
        Debug.Print "SubAddress: " & hl.SubAddress
        Debug.Print "TextToDisplay: " & hl.TextToDisplay
        Debug.Print "ScreenTip: " & hl.ScreenTip
        Debug.Print "Type: " & TypeName(hl.Parent) ' Type of object the hyperlink is attached to (e.g., Range)
        Debug.Print "Anchor Address: " & hl.Range.Address ' The cell or range the hyperlink is in
    Next i
End Sub

Private Sub test_is_File_Linked()
    is_File_Linked Selection
End Sub
Function is_File_Linked(cell As Range) As Boolean
Dim hlink As Hyperlink
    If cell.Hyperlinks.count < 1 Then
        Debug.Print "NO Hyperlink exists for cell address: '" & cell.Address & "'"
        is_File_Linked = False
        Exit Function
    End If
    ' Assuming hlink refers to a specific hyperlink object
    ' e.g., Set hlink = ActiveSheet.Hyperlinks(1)
    Set hlink = cell.Hyperlinks(1)
    
    If newFileOrFolderExists(hlink.Address) Then
        Debug.Print "Hyperlink exists: '" & hlink.Address & "'"
        is_File_Linked = True
    Else
        Debug.Print "Hyperlink does NOT exist: '" & hlink.Address & "'"
        is_File_Linked = False
    End If
End Function


' Ai Generated
Function ExtractFileNameFromHyperlink(cell As Range) As String
Dim fullpath As String
Dim LastSlashPos As Long
    On Error GoTo ErrorHandler ' Handle potential errors
    ' Check if the cell contains a hyperlink
    If cell.Hyperlinks.count > 0 Then
        fullpath = cell.Hyperlinks(1).Address
        
        ' Find the position of the last backslash or forward slash
        LastSlashPos = InStrRev(fullpath, "\")
        If LastSlashPos = 0 Then ' If no backslash, check for forward slash (for URLs)
            LastSlashPos = InStrRev(fullpath, "/")
        End If
        If LastSlashPos > 0 Then
            ' Extract the substring after the last slash
            ExtractFileNameFromHyperlink = Mid(fullpath, LastSlashPos + 1)
        Else
            ' If no slash is found, the entire address might be the filename (e.g., local file in current directory)
            ExtractFileNameFromHyperlink = fullpath
        End If
    Else
        ExtractFileNameFromHyperlink = "" ' Return empty string if no hyperlink
    End If
    Exit Function

ErrorHandler:
    ExtractFileNameFromHyperlink = "Error" ' Return "Error" on error
End Function
    
Private Sub test_ReplaceHyperlinkInCell()
    
End Sub

' Replace Hyperlink (if any ) in a cell with a new Hyperlink
' AI Generated
Sub ReplaceHyperlinkInCell(cell As Range, newAddress As String, newTextToDisplay As String, _
            Optional newScreenTip As String = "")
    ' If the cell currently contains a hyperlink, then delete
    If cell.Hyperlinks.count > 0 Then cell.Hyperlinks.Delete

    ' Add the new hyperlink to the cell
    cell.Hyperlinks.Add _
        Anchor:=cell, _
        Address:=newAddress, _
        TextToDisplay:=newTextToDisplay, _
        ScreenTip:=newScreenTip
End Sub

' Withing the range rg passed in:
'  Verify that the file name from old link exists in folder path passed in.
'  Replace existing link folder path with the folder path passed in.

Sub replaceLinksinRange(rg As Range, fpath As String)
Dim cell As Range, fname As String, newAddress As String
Dim m As String, ans As Long, lastRow As Long

    rg.Worksheet.Activate ' Needed warning msage
    lastRow = Last(1, rg)
    
    For Each cell In rg
        If cell.row > lastRow Then Exit For
        ' Get file name from old link
        fname = ExtractFileNameFromHyperlink(cell)
        If fname = "" Then GoTo next_cell
        ' Test  for validity of new fpath and fname
        newAddress = fpath & "\" & fname
        If Not newFileOrFolderExists(newAddress) Then
            cell.Select
            m = "Warning: New link NOT a valid path for this cell " & cell.Address & "." & vbCrLf & vbCrLf
            m = m & newAddress & vbCrLf & vbCrLf
            m = m & "Set the link anyway?  Yes to set, No to skip this one, Cancel to exit... "
            ans = MsgBox(m, vbYesNoCancel, "Cell Address: " & cell.Address)
            If ans = vbCancel Then Exit Sub
            If ans = vbNo Then GoTo next_cell
        End If
        ' remove and replace link in this cell
        Call ReplaceHyperlinkInCell(cell, newAddress, fname)
next_cell:
    Next cell

End Sub


' THIS NEEDS MORE TESTING
Private Sub test_networkpaths()
' Map the network share to a drive letter (e.g., "Z:")
Dim objShell As Object
Set objShell = CreateObject("WScript.Shell")
    objShell.Run "net use Z: \\nas\bbmdata /persistent:no", 0, True

' Now use the mapped drive with FileSystemObject
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Dim path1 As String
Dim path2 As String
Dim relativePath As String

    path1 = "Z:\Dwight\PFS\" '0 FS STATEMENTS"
    path2 = "Z:\Dwight\K-1s\" 'K-1 Entry Workbooks"
    
    ' The function will now work correctly
    Debug.Print fso.GetRelativePath(path2, path1)
    
    ' Clean up the mapped drive
    objShell.Run "net use Z: /delete", 0, True

End Sub

'  NEEDS MORE TESTING
Private Sub Test_GetRelativePath()
    Dim BasePath As String
    Dim TargetPath As String
    Dim relativePath As String

'    BasePath = "\\NAS\bbmdata\Dwight\PFS\0 FS STATEMENTS"
'    TargetPath = "\\NAS\bbmdata\Dwight\K-1s\K-1 Entry Workbooks"

'    BasePath = "N:\Dwight\PFS\0 FS STATEMENTS"
'    TargetPath = "N:\Dwight\K-1s\K-1 Entry Workbooks"

    BasePath = "C:\Users\Dwight.YACKLEY-ALIENWA\AppData\Roaming\Microsoft\AddIns"
    TargetPath = "C:\myBRC"
    
    relativePath = GetRelativePath(BasePath, TargetPath)
    Debug.Print "Relative Path: " & relativePath
End Sub


' AI genereated  Edge CoPilot
Function GetRelativePath(ByVal BasePath As String, ByVal TargetPath As String) As String
Dim fso As Object
Dim BaseFolder As Object
Dim TargetFolder As Object
Dim relativePath As String
    If Not newFileOrFolderExists(BasePath) Then
        Debug.Print "BasePath does NOT exist: '" & BasePath & "'"
        Exit Function
    End If
    If Not newFileOrFolderExists(TargetPath) Then
        Debug.Print "TargetPath does NOT exist: '" & TargetPath & "'"
        Exit Function
    End If
      
    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set BaseFolder = fso.GetFolder(BasePath)
    Set TargetFolder = fso.GetFolder(TargetPath)
    
    If BaseFolder Is Nothing Or TargetFolder Is Nothing Then
        Debug.Print "Error: BaseFolder or TargetFolder is Nothing!"
        Exit Function
    End If
    
    Debug.Print "fso.getrelativepath: " & fso.GetRelativePath(BaseFolder, TargetFolder)
    
    ' Calculate relative path
    relativePath = fso.BuildPath(BaseFolder.path, fso.GetRelativePath(BaseFolder, TargetFolder))
    
    ' Return the relative path
    GetRelativePath = relativePath
End Function



Private Sub test_Rebuild_Links()
    Call Rebuild_Links(Selection, Selection, "")
End Sub
Sub Rebuild_Links(rng As Range, put_rng As Range, Optional strPath As String = "")
' Rebuild a Hyperlink for each File Name in rng
Dim cell As Range, c As Long, savedstr As String
    On Error Resume Next
    If rng Is Nothing Or put_rng Is Nothing Then Exit Sub
    c = put_rng.Column ' Place to put the Rebuilt Links
    If c < 1 Then Exit Sub ' Invalid Range Name
    If strPath = "" Then strPath = ActiveWorkbook.path & "\"
    
    For Each cell In rng
        If cell.row <= 1 Then GoTo nextOne ' Row one is always a Header
        If cell.row > Last(1, rng) Then Exit For
        savedstr = cell.Value
        Cells(cell.row, c).Clear
        ClearFormat (Cells(cell.row, c))
        ActiveSheet.Hyperlinks.Add Cells(cell.row, c), strPath & savedstr
        Cells(cell.row, c).Value = savedstr
        If Not myFileExists(strPath & cell.Value) Then ErrorFormat (Cells(cell.row, c))
nextOne:
   Next cell
End Sub
Private Sub JUNKtest_Compare_To_K1_List()
    Call JUNKCompare_To_K1_List(Selection)
End Sub
Sub JUNKCompare_To_K1_List(compareTo As Range)
Dim headerArray As Variant, hdr As Variant
Dim i As Long
    ' Compare Header Row First
    headerArray = Array("Year", "Partnership Name", "Partnership TAX ID", _
    "Opened", "Closed", "Member Name", "Member ID", "Tax Ent", "QB #", _
    "QB Balance", "Comments", "LINK", "File Tags", "REBUILD LINKS", "Scratch")
    'i = 1
    For Each hdr In headerArray
        i = Header_Index(compareTo.EntireRow(1), hdr)
        If i < 1 Then MsgBox ("Missing Header '" & hdr & "'"): GoTo nextOne
        If StrComp(hdr, StripAltEnter(compareTo(1, i)), vbTextCompare) <> 0 Then
            compareTo(1, i).Select
            MsgBox ("Mismatch: " & vbCrLf & vbCrLf _
            & hdr & "   " & compareTo(1, i))
        End If
        'i = i + 1
nextOne:
    Next hdr
End Sub
Function Strip_NonPrintableCharacters(str As String) As String
'Removes non-printable characters from a string
'https://stackoverflow.com/questions/41882525/vba-how-to-remove-non-printable-characters-from-data
Dim cleanString As String
Dim i As Integer
    cleanString = str
    For i = Len(cleanString) To 1 Step -1
        'Debug.Print Asc(Mid(Str, i, 1))
        Select Case Asc(Mid(str, i, 1))
            Case 1 To 31, Is >= 127
                'Bad stuff
                'https://www.ionos.com/digitalguide/server/know-how/ascii-codes-overview-of-all-characters-on-the-ascii-table/
                cleanString = Left(cleanString, i - 1) & Mid(cleanString, i + 1)
            Case Else
                'Keep
        End Select
    Next i
    Strip_NonPrintableCharacters = cleanString
End Function
Private Sub test_StripAltEnter()
    MsgBox (StripAltEnter(Range("C1")))
End Sub
Function StripAltEnter(sdata As String) As String
Dim i As Long
    For i = 1 To Len(sdata) Step 1
            If Mid(sdata, i, 1) <> Chr(10) Then
                StripAltEnter = StripAltEnter & Mid(sdata, i, 1)
            Else
                StripAltEnter = StripAltEnter & " " ' Replace chr(10) with SPACE
            End If
    Next i
    StripAltEnter = Trim(StripAltEnter)
End Function
