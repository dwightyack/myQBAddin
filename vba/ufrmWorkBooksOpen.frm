VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufrmWorkBooksOpen 
   Caption         =   "List of OPEN WORKBOOKs"
   ClientHeight    =   4670
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   4100
   OleObjectBlob   =   "ufrmWorkBooksOpen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufrmWorkBooksOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Public variable to store the string passed in
Public ufrmWorkbooksOpenCaption As String

' The last User Selected Workbook
Private SelectedWorkbook As Workbook

' ----------------------------------------------------------------
' Procedure Name: UserForm_Initialize
' ----------------------------------------------------------------
Private Sub UserForm_Initialize()
    ' Use the Caption set by the User in the public variable
    Me.Caption = ufrmWorkbooksOpenCaption
    Call populateOpenWbNamesToListbox
End Sub

' ----------------------------------------------------------------
' Procedure Name: populateOpenWbNamesToListbox
'-----------------------------------------------------------------
Sub populateOpenWbNamesToListbox()
 Dim wb As Workbook
    ' Clear any previous items from the listbox
    lstWorkbooks.Clear

    ' Loop through all open workbooks and add their names to the listbox
    For Each wb In Application.Workbooks
        lstWorkbooks.AddItem wb.name
    Next wb

    ' Pre-select the first item if the list is not empty
    If lstWorkbooks.ListCount > 0 Then
        lstWorkbooks.ListIndex = 0
    End If
End Sub
' ----------------------------------------------------------------
' Procedure Name: cmbWbNamesCancel_Click
' ----------------------------------------------------------------
Private Sub cmbWbNamesCancel_Click()
    Unload Me
End Sub
' ----------------------------------------------------------------
' Procedure Name: cmbRefreshWBList_Click
' ----------------------------------------------------------------
Private Sub cmbRefreshWBList_Click()
    populateOpenWbNamesToListbox
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' THIS IS WHERE THE WORK IS DONE.
'
'  User has Selected a Workbook.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmbSelectWB_Click()
Dim wbName As String
    Debug.Print "cmbSelectWB_Click(): ..."
    
    ' Check if an item is selected
    If lstWorkbooks.ListIndex >= 0 Then
        wbName = lstWorkbooks.Value
        ' Get the workbook object
        Set SelectedWorkbook = Application.Workbooks(wbName)
    End If
    
    ' Hide the form, but don't unload it yet
    Me.Hide

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  User wants to Open another Workbook.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmbOpenAnotherWB_Click()
    Debug.Print "cmbOpenAnotherWB_Click(): ..."
    Set SelectedWorkbook = myOpenWorkbook(True, "*.xls*", , "")
    ' Refresh the list
    cmbRefreshWBList_Click
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' This is cool! I didn't know I could retreive properties from a form!!
'
' Property to allow the standard module to retrieve the selected workbook
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get GetSelectedWorkbook() As Workbook
    Set GetSelectedWorkbook = SelectedWorkbook
End Property



