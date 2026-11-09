Attribute VB_Name = "my_ribbon_Ctx_MISC"
Option Explicit
Option Private Module

' **********************************************
'
' Stuff I Added
'
' **********************************************

'Callback for Toggle Comments
Dim myQBAddin_CommentsToggle As Boolean ' Toggle State stored in the scope of this module
Sub ToggleComments_click()
    'MsgBox "Show Comments value: " & returnedVal
    If Not myQBAddin_CommentsToggle Then  ' COMMENTS ON
        Application.DisplayCommentIndicator = xlCommentAndIndicator
    Else  '   COMMENTS OFF
        Application.DisplayCommentIndicator = xlCommentIndicatorOnly
    End If
    myQBAddin_CommentsToggle = Not myQBAddin_CommentsToggle
End Sub
'Callback for ClearMyFill
Sub ClearMyFill_click()
Dim s As String
    On Error Resume Next
    Call ClearFillandComments_Click
End Sub

'Callback for CheckFileLinks
Sub CheckFileLinks_click()
    On Error Resume Next
    Call ActiveWorkbook.Test_MyFileLinks
End Sub

'Callback for ClearAcctLinks
Sub ClearAcctLinks_click()
    On Error Resume Next
    Call ActiveWorkbook.Clear_MyAcct_Links
End Sub
'Callback for BuildAcctLinks
Sub BuildAcctLinks_click()
    On Error Resume Next
    ' The active workbook will likely
    ' call myQBAddin.Build_Acct_Links
    ' with a range
    Call ActiveWorkbook.Build_MyAcct_Links
End Sub

'Callback for BuildLinksFromFileTags
Sub BuildLinksFromFileTags_click()
    On Error GoTo error1
    Call ActiveWorkbook.Build_MyLinks_From_FileTags
    Exit Sub
error1:
    ' Application.VBE.ActiveCodePane.CodeModule.name IS SUPPPOSED TO GIVE THE MODULE NAME WHERE THIS VBA
    MsgBox "Macro 'Build_MyLinks_From_FileTags'" & vbCrLf & _
        "does NOT exist" & "in ActiveWorkbook." & _
        vbCrLf & vbCrLf & "'" & ActiveWorkbook.name & "'.", _
        vbInformation, Application.VBE.ActiveCodePane.CodeModule.name
End Sub

'Callback for BuildFileTags
Sub BuildFileTags_click()
    On Error Resume Next
    Call ActiveWorkbook.Build_MyFileTags
End Sub

'Callback for Clear_Cells
Sub Clear_Cells_click()
    On Error Resume Next
    Call ActiveWorkbook.Clear_MyCells
End Sub
'https://www.mrexcel.com/board/threads/easy-way-to-check-if-a-macro-exists-or-ignore-the-call-if-it-doesnt.393551/
Sub CallTest1_click()
Dim ans As Long
    On Error GoTo Error_Test1
    CallByName ActiveWorkbook, "Test1", VbMethod
    Exit Sub
Error_Test1:
    ans = MsgBox("Test1 DOES NOT EXIST - OR FAILED in ActiveWorkbook '" & ActiveWorkbook.name & "'.", vbExclamation)
End Sub
Sub CallTest2_click()
Dim ans As Long
    On Error GoTo Error_Test2
    CallByName ActiveWorkbook, "Test2", VbMethod
    Exit Sub
Error_Test2:
    ans = MsgBox("Test2 DOES NOT EXIST - OR FAILED in ActiveWorkbook '" & ActiveWorkbook.name & "'.", vbExclamation)
End Sub
Sub CallTest3_click()
Dim ans As Long, m As String
    On Error GoTo Error_Test3
    CallByName ActiveWorkbook, "Test3", VbMethod
    Exit Sub
Error_Test3:
    m = "Test3 FAILED. Error " & Err & " in ActiveWorkbook '" & ActiveWorkbook.name & "'." & vbCrLf & vbCrLf
    m = m & Err.Description & "."
    ans = MsgBox(m, vbExclamation)
    Err.Clear
End Sub
Sub CallTest4_click()
Dim ans As Long
    On Error GoTo Error_Test4
    CallByName ActiveWorkbook, "Test4", VbMethod
    Exit Sub
Error_Test4:
    ans = MsgBox("Test4 DOES NOT EXIST - OR FAILED in ActiveWorkbook '" & ActiveWorkbook.name & "'.", vbExclamation)
End Sub

' **********************************************
'
' MACROS FROM ORIGINAL DD LIST
'
' https://www.contextures.com/excelribbondynamictab.html
' **********************************************
' Misc Macros from ORIGINAL DD List

Sub ToggleGridlines()
With ActiveWindow
  .DisplayGridlines = Not .DisplayGridlines
End With
End Sub

Sub ToggleHeadings()
With ActiveWindow
  .DisplayHeadings = Not .DisplayHeadings
End With
End Sub

Sub ToggleRefStyle()
On Error Resume Next
Application.ReferenceStyle _
  = xlA1 + xlR1C1 _
  - Application.ReferenceStyle
End Sub

Sub ToggleZeros()
On Error Resume Next
With ActiveWindow
  .DisplayZeros = Not .DisplayZeros
End With
End Sub
'https://www.contextures.com/excelribbonmacrostab.html

Sub ColourList()
'colours cells in column A, and lists the colours
Dim wsList As Worksheet
Dim lRow As Long
'On Error Resume Next

Set wsList = Sheets.Add
For lRow = 1 To 56  '1=black, 2=white
 With wsList
     .Cells(lRow + 1, 2) = lRow
     .Cells(lRow + 1, 1).Interior.ColorIndex = lRow
  End With
Next lRow
End Sub

Sub EventsOn()
  Application.EnableEvents = True
  Application.ScreenUpdating = True
  'MsgBox "Events On"
End Sub

Sub PivotCacheCount()
MsgBox "Pivot Cache Count: " _
  & ActiveWorkbook.PivotCaches.count
End Sub

Sub ShowAllSheets()
Dim sh
For Each sh In ActiveWorkbook.Sheets
    If Not sh.Visible Then
        'MsgBox "Unhiding sheet name '" & sh.name & "'"
        Debug.Print "Unhiding sheet name '" & sh.name & "'"
        sh.Visible = True
    End If
Next sh
End Sub

Sub CtxMainPage(control As IRibbonControl)
Dim strLink As String
strLink = "https://www.contextures.com"
On Error GoTo ErrHandler
ThisWorkbook.FollowHyperlink _
  Address:=strLink, NewWindow:=True
Exit Sub

ErrHandler:
    MsgBox "Cannot open " & strLink
End Sub

Sub CtxHelpPage(control As IRibbonControl)
Dim strLink As String
strLink = "https://www.contextures.com/excelribbonmacrostab.html"
On Error GoTo ErrHandler
ThisWorkbook.FollowHyperlink _
  Address:=strLink, NewWindow:=True
Exit Sub

ErrHandler:
    MsgBox "Cannot open " & strLink

End Sub

