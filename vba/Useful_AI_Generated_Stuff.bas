Attribute VB_Name = "Useful_AI_Generated_Stuff"
Option Explicit


' DY 1-30-26  Move this ai generated macro to a "misc usefull func" mod LATER
Function Copy_Sheet(destWB As Workbook, destName As String, srcWb As Workbook, srcsheetName As String)
    Dim m As String
    
    If Not myQBAddin.SheetExists(srcsheetName, srcWb) Then
        m = "Sheet '" & srcsheetName & "' does not exist in '" & srcWb.name & "'."
        GoTo ErrorHandler
    End If
    
    'On Error Resume Next
   
    ' Copy the new sheet from the source workbook to the destination workbook
    '    It will be placed before existing sheet in the destination workbook.
    
    If myQBAddin.SheetExists("old_" & destName, destWB) Then
        srcWb.Sheets(srcsheetName).Copy After:=destWB.Sheets("old_" & destName)
    Else
        srcWb.Sheets(srcsheetName).Copy After:=destWB.Sheets(destWB.Sheets.count)
    End If
    ' The newly copied sheet becomes the ActiveSheet. Rename it to the original name.
    destWB.Sheets(srcsheetName).name = destName
    'MsgBox "Dest Wb = " & destWB.name
    
    ' Clean up error handler
    On Error GoTo 0
    Copy_Sheet = True
    Exit Function

ErrorHandler:
    m = m & " Copy_Sheet(): An error occurred. " & Err.Description
    Debug.Print m
    'MsgBox M
    Copy_Sheet = m
End Function


' AI Generated
'   saveAndConfirm
'
'   Saves a workbook.
'   Returns True if success.
'   Displays and Returns errors if failed.
Function saveAndConfirm(wb As Workbook)
    ' Define a variable to hold the message for the user.
    Dim confirmationMessage As String
    ' Define a variable to hold the workbook name.
    Dim wbName As String
    
    ' Get the name of the current workbook.
    wbName = wb.name
    
    ' Attempt to save the workbook.
    On Error GoTo ErrorHandler
    
    wb.Save
    
    ' If the save is successful, set the confirmation message.
    confirmationMessage = "Successfully saved the workbook: " & vbCrLf & wbName
    ' Display a success message box (vbInformation icon).
    MsgBox confirmationMessage, vbInformation, "Save Successful"
    
    ' Exit the sub to prevent running the error handler code if successful.
    saveAndConfirm = True
    Exit Function
    
ErrorHandler:
    ' If an error occurs during the save process, this code runs.
    ' Set the message to indicate failure and show the error description.
    confirmationMessage = "Could not save the workbook '" & wb.name & "'. " & vbCrLf & "Error: " & Err.Description
    ' Display an error message box (vbCritical icon).
    Debug.Print confirmationMessage
    MsgBox confirmationMessage, vbCritical, "saveAndConfirm(): Save Failed"
    
    ' Clear the error object.
    Err.Clear
End Function
