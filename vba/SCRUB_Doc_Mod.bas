Attribute VB_Name = "SCRUB_Doc_Mod"
Option Explicit
Sub Scrub_Doc()
Dim m As String, ans As Long
    
    MsgBox "Removing Document Info..."
    ActiveWorkbook.RemoveDocumentInformation (99) ' Removes ALL Document Information
    '==============================================================
    ' Purpose: Toggle BACK "Remove personal information from file
    '          properties on save".
    '          This needs to be reset after (99) call above.
    '==============================================================
    ActiveWorkbook.RemovePersonalInformation = False

    MsgBox "Breaking External Links..."
    Call BreakExternalLinks
    
    m = "Unhide Hidden Worksheets = YES"
    ans = MsgBox(m, vbYesNo, "Unhide HIDDEN WORKSHEETS?")
    
    If ans = vbYes Then Call UnhideAllSheets
    
    MsgBox "Unhiding Hidden Rows and Columns..."
    Call UnhideAllRowsColumns
    
    Call UnhideAllNames
    
    ans = MsgBox("Run Workbook Inspector?", vbOKCancel)
    If ans = vbCancel Then Exit Sub
    Call Examine_Workbook
End Sub
Sub Examine_Workbook()
Dim docStatus As MsoDocInspectorStatus
Dim insp As DocumentInspector, n As name
Dim result As String, i As Long, m As String, ans As Long
Dim ExternalLinks As Variant, X As Long, t As String
    t = "DOCUMENT INSPECT    " & ActiveWorkbook.name
      
    ' Count Hidden Names
    i = 0
    For Each n In ActiveWorkbook.Names
         If n.Visible = False Then i = i + 1
    Next n
    m = Format(i, "#,##0") & "  Hidden Name(s) found." & vbCrLf & vbCrLf
    ans = MsgBox(m, vbOKOnly, t)
    
    ' Count External Links
    i = 0
    'Create an Array of all External Links stored in Workbook
    ExternalLinks = ActiveWorkbook.LinkSources(Type:=xlLinkTypeExcelLinks)
    If IsEmpty(ExternalLinks) = True Then GoTo no_links
    For X = 1 To UBound(ExternalLinks)
      i = i + 1
    Next X
no_links:
    m = Format(i, "#,##0") & "  External Link(s) found." & vbCrLf & vbCrLf
    ans = MsgBox(m, vbOKOnly, t)

    ' Run thru Inspectors Collection
    i = 1
    For Each insp In ActiveWorkbook.DocumentInspectors
        insp.Inspect docStatus, result
        m = i & "   Document Inspector Number " & vbTab & docStatus & vbCrLf
        ans = MsgBox(m & vbCrLf & _
        "Result: " & result, vbOKOnly, t)
        i = i + 1
    Next
    
    Call see_DocumentBuiltinProperties
    Call Open_DocumentInspector
    '==============================================================
    ' Purpose: Toggle BACK "Remove personal information from file
    '          properties on save". = FALSE
    '          This needs to be reset after (99) call above.
    '==============================================================
    ActiveWorkbook.RemovePersonalInformation = False
End Sub
Sub Open_DocumentInspector()
    ' Displays Excel's Document Inspector
    Application.Dialogs(xlDialogDocumentInspector).Show
End Sub
'https://docs.microsoft.com/en-us/office/vba/api/excel.workbook.builtindocumentproperties
' Loads all properties in a worksheet
' MAKE SURE TO HAVE A BLANK WORKSHET ACTIVE
Sub see_DocumentBuiltinProperties()
Dim rw As Long, p As Variant
Dim m As String, result As Long
On Error Resume Next

    m = "Are you on a BLANK worksheet??" & vbCrLf & vbCrLf
    result = MsgBox(m, vbYesNoCancel, ActiveSheet.name)
    If result = vbCancel Then Exit Sub
    If result = vbNo Then Sheets.Add
    
    Columns("A:A").ColumnWidth = 30
    Columns("B:B").ColumnWidth = 20

    rw = 1
    For Each p In ActiveWorkbook.BuiltinDocumentProperties
        Cells(rw, 1).Value = p.name
        Cells(rw, 2).Value = p
        rw = rw + 1
    Next
    ' Format Date and Time Cells
    Range("B10:B12").Select
    Selection.NumberFormat = "[$-en-US]m/d/yy h:mm AM/PM;@"
End Sub

Sub RemoveDocInfo_AndLinks()
Dim m As String, ans As Long
    
    ActiveWorkbook.RemoveDocumentInformation (99) ' Removes ALL Document Information
    
    Call BreakExternalLinks
    
    'Application.Dialogs(xlDialogDocumentInspector).Show
End Sub
'https://www.thespreadsheetguru.com/vba/2014/10/10/break-all-external-links-excel
Sub BreakExternalLinks()
Dim m As String
'PURPOSE: Breaks all external links that would show up in Excel's "Edit Links" Dialog Box
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim ExternalLinks As Variant
Dim X As Long
Dim BreakCounter As Long

'Create an Array of all External Links stored in Workbook
  ExternalLinks = ActiveWorkbook.LinkSources(Type:=xlLinkTypeExcelLinks)
  If IsEmpty(ExternalLinks) = True Then GoTo ReportResults

'Loop Through each External Link in ActiveWorkbook and Break it
  For X = 1 To UBound(ExternalLinks)
    Debug.Print " Breaking Link: " & ExternalLinks(X)
    ActiveWorkbook.BreakLink name:=ExternalLinks(X), Type:=xlLinkTypeExcelLinks
    BreakCounter = BreakCounter + 1
  Next X

'Report to User how many External Links were Broken
ReportResults:
  If BreakCounter = 0 Then
    m = "No External Links found..."
  Else
    m = "External Links broken: " & BreakCounter & vbCrLf & vbCrLf
    m = m & "See Immediate Window for more details."
  End If
  MsgBox m
  

End Sub
Sub UnhideAllSheets()
Dim ws As Worksheet, count As Long
Dim m As String
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
            count = count + 1
            Debug.Print " UnHiding worksheet: " & ws.name & vbCrLf
        End If
    Next ws
    If count = 0 Then
        m = "No Hidden Sheets found."
    Else
        m = Format(count, "#,##0") & " Hidden Worksheet(s) are now visible." & vbCrLf & vbCrLf
        m = m & "See Immediate Window for more details."
    End If
    MsgBox m
End Sub
Sub UnhideAllRowsColumns()
Dim ws As Worksheet, m As String
Dim countRows As Long, countColumns As Long
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Columns.EntireColumn.Hidden = True Then
            ws.Columns.EntireColumn.Hidden = False
            countColumns = countColumns + 1
            Debug.Print " UnHiding Column in Worksheet: " & ws.name & vbTab & vbCrLf
        End If
        If ws.Rows.EntireRow.Hidden = True Then
            ws.Rows.EntireRow.Hidden = False
            countRows = countRows + 1
        End If
    Next ws
    If (countColumns + countRows = 0) Then
        m = "No Hidden Rows or Columns found."
    Else
        m = Format(countColumns, "#,##0") & " Hidden Columns(s) are now visible." & vbCrLf
        m = m & Format(countRows, "#,##0") & " Hidden Rows(s) are now visible." & vbCrLf
        m = m & "See Immediate Window for more details."
    End If
  MsgBox m
End Sub
Sub test_UnhideAllNames()
    Call UnhideAllNames
End Sub
' https://excelhelphq.com/how-to-unhide-or-delete-name-ranges-in-excel/
' Show HIDDEN names useful when using Excel Document Inspector
'
Sub UnhideAllNames(Optional prompt As Boolean = True)
Dim n As name, answer As Long, count As Long, m As String
    For Each n In ActiveWorkbook.Names
        If n.Visible = False Then count = count + 1
    Next n
    
    m = Format(count, "#,##0") & "  Hidden Name(s) found." & vbCrLf & vbCrLf
    Debug.Print m
    If count = 0 Then
        If prompt Then answer = MsgBox(m, vbInformation + vbOKOnly, ActiveSheet.name)
        Exit Sub
    End If
    
    m = m & "Unhide ALL Hidden Names?"
    If prompt Then
        answer = MsgBox(m, vbYesNoCancel, ActiveWorkbook.name)
        Select Case answer
            Case vbCancel
                Exit Sub
            Case vbNo
                Exit Sub
        End Select
    End If
    
    count = 0
    For Each n In ActiveWorkbook.Names
        If n.Visible = False Then
            n.Visible = True
            count = count + 1
        End If
    Next n
    m = Format(count, "#,##0") & "  Hidden Name(s) are now Shown." & vbCrLf & vbCrLf
    If prompt Then answer = MsgBox(m, vbInformation + vbOKOnly, ActiveWorkbook.name)
    Application.Dialogs(xlDialogNameManager).Show
End Sub
' Remove hidden names Interactively on active workbook
Sub Remove_Hidden_Names()
       Dim n As Variant, result As Variant
       Dim count As Long, m As String
       ' Loop once for each name in the workbook.
       count = 0
       For Each n In ActiveWorkbook.Names
            If n.Visible = False Then count = count + 1
       Next n
        m = Format(count, "#,##0") & "  Hidden Name(s) found." & vbCrLf & vbCrLf
        result = MsgBox(m, vbInformation + vbOKOnly, ActiveWorkbook.name)
        
       If count = 0 Then Exit Sub ' No Hidden Names
       
       ' Loop thru each hidden name in the workbook.
       For Each n In ActiveWorkbook.Names

           'If a name is not visible (it is hidden)...
           If n.Visible = True Then GoTo Next_Name

           ' ...ask whether or not to delete Hidden name.
           result = MsgBox(vbTab & n.name & "?" & vbCrLf & vbCrLf & _
               "Refers to: " & Chr(10) & vbCrLf & n.RefersTo & vbCrLf _
               , vbYesNoCancel, "DELETE HIDDEN NAME in '" & ActiveWorkbook.name & "'")

           ' If the result is Yes, then delete the name. Cancel Exits sub
           If result = vbYes Then n.Delete
           If result = vbCancel Then Exit Sub
Next_Name:
       Next n
End Sub

' Debug Print hidden names on active workbook
Sub Dbg_Print_Hidden_Names()
       Dim n As name, result As Variant
       Dim count As Long, m As String
       ' Loop once for each name in the workbook.
       count = 0
       For Each n In ActiveWorkbook.Names
            If n.Visible = False Then
                count = count + 1
                Debug.Print count & ": " & n.name & "    VALUE:" & n.Value & "     RefersTo:" & n.RefersTo & " "
            End If
       Next n
       'M = Format(count, "#,##0") & "  Hidden Name(s) found." & vbCrLf & vbCrLf
       'result = MsgBox(M, vbInformation + vbOKOnly, ActiveWorkbook.name)
        
End Sub


' https://www.msofficeforums.com/word-vba/8227-batch-run-inspect-documents.html
' Needs Microsoft Doc Source References checked
'Sub Anonymizer()
'Dim strInFold As String, strOutFold As String, strFile As String, strOutFile As String, DocSrc As Document
'    Application.ScreenUpdating = False
'
'    'Call the GetFolder Function to determine the folder to process
'    strInFold = GetFolder
'    If strInFold = "" Then Exit Sub
'    strFile = Dir(strInFold & "\*.doc", vbNormal)
'
'    'Check for documents in the folder - exit if none found
'    If strFile <> "" Then strOutFold = strInFold & "\Output\"
'
'    'Test for an existing outpfolder & create one if it doesn't already exist
'    If Dir(strOutFold, vbDirectory) = "" Then MkDir strOutFold
'    strFile = Dir(strInFold & "\*.doc", vbNormal)
'
'    'Process ALL Documents in the chosen folder
'    While strFile <> ""
'      Set DocSrc = Documents.Open(FileName:=strInFold & "\" & strFile, AddTorecentFiles:=False, Visible:=False)
'      With DocSrc
'        'Remove personal information
'        .RemoveDocumentInformation (wdRDIDocumentProperties)
'        'String variable for the output filenames
'        strOutFile = strOutFold & Split(.name, ".")(0)
'       'Save and close the document
'        .SaveAs FileName:=strOutFile
'        .Close
'      End With
'      strFile = Dir()
'    Wend
'    Set rng = Nothing: Set DocSrc = Nothing
'    Application.ScreenUpdating = True
'End Sub
'Function GetFolder(Optional Title As String, Optional RootFolder As Variant) As String
'On Error Resume Next
'GetFolder = CreateObject("Shell.Application").BrowseForFolder(0, Title, 0, RootFolder).Items.Item.path
'End Function


' https://docs.microsoft.com/en-us/office/vba/api/office.documentinspector.inspect
' Lists All Document Inspector Object names
Private Sub List_Document_Inspectors()
Dim DocI As DocumentInspector
Dim m As String
    For Each DocI In ActiveWorkbook.DocumentInspectors
        m = DocI.name
        Debug.Print m
    Next
End Sub
' https://docs.microsoft.com/en-us/office/vba/api/office.documentinspector.inspect
' Run Document Inspector Object number 1
Private Sub DI_InspectDocument()
Dim docStatus As MsoDocInspectorStatus
Dim Results As String
ActiveWorkbook.DocumentInspectors(1).Inspect docStatus, Results
 
MsgBox ("The inspection returned the following status " & ActiveWorkbook.DocumentInspectors(1).name & " " & docStatus & _
" with this result " & Results)
End Sub
' https://docs.microsoft.com/en-us/office/vba/api/office.documentinspector.inspect
' Run All Document Inspector Objects
Sub Dbg_Print_Document_Inspectors()
Dim DocI As DocumentInspector
Dim docStatus As MsoDocInspectorStatus
Dim Results As String
Dim m As String
    For Each DocI In ActiveWorkbook.DocumentInspectors
        DocI.Inspect docStatus, Results
        m = "Doc Inspection: " & DocI.name & " Status:" & docStatus & " Result:" & Results
        Debug.Print vbNewLine & m
    Next
End Sub

' Display Meta Data for Active work book
Sub Dbg_Print_FileProperties()
Dim p
On Error Resume Next
For Each p In ActiveWorkbook.BuiltinDocumentProperties
    Debug.Print p.name & " : " & p.Value
Next
End Sub

' Displays Meta Data and windows extended attributes of a directory

'https://stackoverflow.com/questions/5651890/using-vba-to-get-extended-file-attributes
Sub test_MetaData_Dir()
Dim sFile As Variant
Dim oShell: Set oShell = CreateObject("Shell.Application")
'Dim oDir:   Set oDir = oShell.Namespace("n:\dwight\junk")
Dim oDir:   Set oDir = oShell.Namespace(ActiveWorkbook.path) ' Directory of ACTIVE WORKBOOK
Dim XXX As Long, m As String

For Each sFile In oDir.Items
    For XXX = 0 To 40
        m = sFile.name & " " & XXX & ": " & oDir.GetDetailsOf(sFile, XXX)
        Debug.Print m
    Next
Next
End Sub

'==============================================================
' Module: PrivacySettings
' Purpose: Toggle "Remove personal information from file
'          properties on save" for the active workbook
'==============================================================
Sub DisableRemovePersonalInfo()
    ActiveWorkbook.RemovePersonalInformation = False
End Sub

Sub EnableRemovePersonalInfo()
    ActiveWorkbook.RemovePersonalInformation = True
End Sub
