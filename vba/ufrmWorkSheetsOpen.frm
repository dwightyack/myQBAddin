VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufrmWorkSheetsOpen 
   Caption         =   "List of Worksheets"
   ClientHeight    =   4670
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   4100
   OleObjectBlob   =   "ufrmWorkSheetsOpen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufrmWorkSheetsOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

' Public variable to store the workbook
Private SelectedWorkbook As Workbook

' The last User Selected Worksheet
Private SelectedWorksheet As Worksheet

' ----------------------------------------------------------------
' Procedure Name: UserForm_Initialize
' ----------------------------------------------------------------
Private Sub UserForm_Initialize()
    ' Populate worksheets if workbook is set
    If Not SelectedWorkbook Is Nothing Then
        Call populateWorksheetsToListbox
    End If
End Sub

' ----------------------------------------------------------------
' Procedure Name: SetWorkbook
' ----------------------------------------------------------------
Public Sub SetWorkbook(wb As Workbook)
    Set SelectedWorkbook = wb
    Me.Caption = "Worksheets in " & wb.name
    Call populateWorksheetsToListbox
End Sub

' ----------------------------------------------------------------
' Procedure Name: populateWorksheetsToListbox
'-----------------------------------------------------------------
Sub populateWorksheetsToListbox()
    Dim ws As Worksheet
    ' Clear any previous items from the listbox
    lstWorkbooks.Clear

    If SelectedWorkbook Is Nothing Then Exit Sub

    ' Loop through all worksheets in the workbook and add their names to the listbox
    For Each ws In SelectedWorkbook.Worksheets
        lstWorkbooks.AddItem ws.name
    Next ws

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
    populateWorksheetsToListbox
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' THIS IS WHERE THE WORK IS DONE.
'
'  User has Selected a Worksheet.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmbSelectWB_Click()
    Dim wsName As String
    Debug.Print "cmbSelectWB_Click(): ..."
    
    ' Check if an item is selected
    If lstWorkbooks.ListIndex >= 0 Then
        wsName = lstWorkbooks.Value
        ' Get the worksheet object
        Set SelectedWorksheet = SelectedWorkbook.Worksheets(wsName)
    End If
    
    ' Hide the form, but don't unload it yet
    Me.Hide
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Property to allow the standard module to retrieve the selected worksheet
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get GetSelectedWorksheet() As Worksheet
    Set GetSelectedWorksheet = SelectedWorksheet
End Property
