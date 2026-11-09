VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufrmSheetNames2 
   Caption         =   "List of sheets"
   ClientHeight    =   4770
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   4560
   OleObjectBlob   =   "ufrmSheetNames2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufrmSheetNames2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'https://www.excelcise.org/how-to-listpopulate-sheet-names-to-a-listbox-on-a-userform-vba/
'
' MY FORM FOR SELECTING A WORKSHEET AND PERFORMING SOME ACTION
'
' *** Variables used after events in this mod

' ----------------------------------------------------------------
' Procedure Name: UserForm_Initialize
' Purpose: Initialize list when user form is loaded.
' ----------------------------------------------------------------
Private Sub UserForm_Initialize()
    Call populateSheetNamesToListbox
End Sub
' ----------------------------------------------------------------
' Purpose: Populate sheet names to user form listbox.
' ----------------------------------------------------------------
Sub populateSheetNamesToListbox()
    Dim sh As Worksheet
    
    ' *** If only ONE work sheet in workbook. Go straight to process Action ***
    If ActiveWorkbook.Worksheets.count = 1 Then
        Debug.Print "Only One worksheet in workbook... Calling Process_Selection."
        Call Process_Selection(ActiveWorkbook.Worksheets(1).name)
    End If
    
    ' *** Otherwise List all the sheets in the workbook and add their name to listbox.
    For Each sh In ActiveWorkbook.Worksheets ' Changed to ActiveWorkbook from ThisWorkBook
        lstbSheetNamesNames.AddItem sh.name
        lstbSheetNamesNames.ListIndex = lstbSheetNamesNames.ListCount - 1 ' selects last Item in listbox (I added this)
    Next sh
End Sub
' ----------------------------------------------------------------
' Purpose: Unload user form when Cancel button is clicked.
' ----------------------------------------------------------------
Private Sub cmbSheetNamesCancel_Click()
    Unload Me
End Sub
' ----------------------------------------------------------------
' Purpose: Refresh the sheet names list in the user form, called by Refresh button.
' ----------------------------------------------------------------
Private Sub cmbSheetNamesRefresh_Click()
    lstbSheetNamesNames.Clear
    Call populateSheetNamesToListbox
End Sub
' *** The User Has Selected a Sheet and Pressed the GoTo BUTTON ****
'    Navigate to the selected sheet, called by GoTo button.
'-------------------------------------
Private Sub cmbSheetsNamesGoTo_Click()
    Dim selectedSheet As String
    Dim i As Long
    'Find the selected item in the listbox
    For i = 0 To lstbSheetNamesNames.ListCount - 1
        If lstbSheetNamesNames.Selected(i) = True Then
            selectedSheet = lstbSheetNamesNames.List(i)
            Exit For
        End If
    Next i
    ' If any item was selected, GoTo THAT SHEET
    If selectedSheet <> "" Then
        ThisWorkbook.Worksheets(selectedSheet).Activate
    Else
        MsgBox "No sheet is selected.", vbCritical, "SELECTION ERROR"
    End If
End Sub
' *** The User Has Selected a Sheet and Pressed the SELECT BUTTON ****
'   Process to the sheet selected by the SELECT button.
'-------------------------------------
Private Sub SelectSheetName_Click()
' List the Sheet Names of original wb
    Dim selectedSheet As String
    Dim i As Long
    'Find the selected item in the listbox
    For i = 0 To lstbSheetNamesNames.ListCount - 1
        If lstbSheetNamesNames.Selected(i) = True Then
            selectedSheet = lstbSheetNamesNames.List(i)
            Exit For
        End If
    Next i
    'Check if any item was selected.
    If selectedSheet = "" Then
        MsgBox "No sheet is selected.", vbCritical, "SELECTION ERROR"
        Unload Me: Exit Sub
    End If
    
    ' Call ACTIONS HERE !!!!! ......................
    
    ' Process the selected sheet
    Call Process_Selection(selectedSheet)
    
    Unload Me ' I Added this too
End Sub



