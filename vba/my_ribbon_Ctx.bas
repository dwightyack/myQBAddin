Attribute VB_Name = "my_ribbon_Ctx"
Option Explicit
Option Private Module

'https://www.contextures.com/excelribbondynamictab.html
Public myQBAddin_Ribbon As IRibbonUI

Sub myQBAddin_Onload(ribbon As IRibbonUI)
'Create a ribbon instance for use in this project
Set myQBAddin_Ribbon = ribbon
End Sub
Sub refreshRibbon()
Dim retval
 ' MACRO NOT TESTED DY 1-17-26
  ' Invalidates the caches of all of this add-in's controls
  'retval = myQBAddin_Ribbon.Invalidate() ??
End Sub

    
'<customUI onLoad="RibbonLoaded_myAddin" xmlns="http://schemas.microsoft.com/office/2006/01/customui">

'DropDown onAction
Sub myQBAddin_MyMacroDD(ByVal control As IRibbonControl, _
                selectedID As String, _
                selectedIndex As Integer)
On Error Resume Next
Dim strList As String
Dim strMacro As String
Select Case control.ID
  Case "myQBAddin_DD1"
    strList = "Macro_Ctx_List"
  Case "myQBAddin_DD2"
    strList = "Macro2List"
  Case "myQBAddin_DD3"
    strList = "Macro3List"
End Select

strMacro = ThisWorkbook.Names(strList) _
  .RefersToRange.Rows(CLng(selectedIndex + 1)).Value

Application.Run (strMacro)
'Restore control to original state
myQBAddin_Ribbon.InvalidateControl control.ID
End Sub

'Callback for DropDown GetItemCount
Sub myQBAddin_GetItemCount(ByVal control As IRibbonControl, ByRef count)
Dim strList As String
Select Case control.ID
  Case "myQBAddin_DD1"
    strList = "Macro_Ctx_List"
  Case "myQBAddin_DD2"
    strList = "Macro2List"
  Case "myQBAddin_DD3"
    strList = "Macro3List"
End Select
count = ThisWorkbook.Names(strList) _
          .RefersToRange.Rows.count
End Sub

'Callback for DropDown GetItemLabel
Sub myQBAddin_GetItemLabel(ByVal control As IRibbonControl, _
    index As Integer, ByRef label)
Dim rngML As Range
Dim strList As String
Select Case control.ID
  Case "myQBAddin_DD1"
    strList = "Macro_Ctx_List"
  Case "myQBAddin_DD2"
    strList = "Macro2List"
  Case "myQBAddin_DD3"
    strList = "Macro3List"
End Select

Set rngML = ThisWorkbook.Names(strList) _
  .RefersToRange
label = rngML.Cells(index + 1)

End Sub

'Callback for DropDown getSelectedItemIndex
Sub myQBAddin_GetSelItemIndex(ByVal control As IRibbonControl, ByRef index)
'Ensure first item in dropdown is displayed.
Select Case control.ID
  Case Is = "myQBAddin_DD1"
    index = 0
  Case Is = "myQBAddin_DD2"
    index = 0
  Case Is = "myQBAddin_DD3"
    index = 0
  Case Else
End Select
End Sub

'Callback for Button onAction
Sub myQBAddin_MyMacroBtn(ByVal control As IRibbonControl)
Dim rngBL As Range
Dim c As Range
Dim strMacro As String
Dim lBtn As Long
Dim strBtn As String
Dim m As String
On Error Resume Next
    ' The main list - the first table of Sheet 'Admin'
    Set rngBL = ThisWorkbook.Names("BtnIdList").RefersToRange
    
    'lBtn is number 1 to xx
    lBtn = WorksheetFunction.Match(control.ID, rngBL, 0)
    If lBtn >= 1 Then
       ' get the  Macro Name - two columns to the right
        strBtn = rngBL.Cells(lBtn, 1).Offset(0, 2).Value
        If strBtn = "" Then
              m = "MyMacroBtn(): Blank Cell in Macros Table... Can't run " & lBtn & " " & rngBL.Cells(lBtn, 1)
              Debug.Print m
              MsgBox ThisWorkbook.name & vbCrLf & vbCrLf & m, vbCritical, ThisWorkbook.name
              Exit Sub
        End If
        Application.Run (strBtn)
    End If
End Sub

'Callback for Button GetLabelBtn
Sub myQBAddin_GetLabelBtn(control As IRibbonControl, ByRef returnedVal)
Dim rngBL As Range
Dim c As Range
Dim strMacro As String
Dim lBtn As Long
On Error Resume Next
Set rngBL = ThisWorkbook.Names("BtnIdList") _
        .RefersToRange
lBtn = WorksheetFunction.Match(control.ID, rngBL, 0)
returnedVal = rngBL.Cells(lBtn, 1).Offset(0, 1).Value
End Sub

'Callback for Button getImage
Sub myQBAddin_GetImageBtn(control As IRibbonControl, ByRef returnedVal)
Dim rngBL As Range
Dim c As Range
Dim strMacro As String
Dim lBtn As Long
On Error Resume Next
Set rngBL = ThisWorkbook.Names("BtnIdList").RefersToRange
lBtn = WorksheetFunction.Match(control.ID, rngBL, 0)
returnedVal = rngBL.Cells(lBtn, 1).Offset(0, 3).Value
End Sub
