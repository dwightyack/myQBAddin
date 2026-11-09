Attribute VB_Name = "BRC_Import_Lists"
Option Explicit

Sub BRC_COA_Click()  ' Run macro from a shortcut
    
    ' Very Important to have QB setup right and BRC List Reporter installed
        
    MsgBox "This Code is not done yet."
    
End Sub

Sub BRC_List_Click()  ' Run macro from a shortcut
    ' Very Important to have QB setup right and BRC List Reporter installed
    Call BRC_List_Report_To_WB(ActiveWorkbook, "", "Vendors", True)
End Sub


' Very Important to have QB setup right and BRC List Reporter installed
Sub BRC_List_Report_To_WB(dest_Wb As Workbook, QB_NAME As String, _
                            dest_sheetName As String, remove_Old As Boolean)
Dim m As String, ans As Long, brcLoaded As Boolean
Dim wb As Workbook, resultSuccess, companyFilesList As String, OpenNames_array() As String
Dim st_col As Long, lastCol As Long, prevState As Boolean
    
    ' 1. Get List of Open QB Company Files
    MsgBox "Exec pwsh Get-OpenQBWCountAndNames.ps1." & vbCrLf & vbCrLf & "This will take a few seconds..."
    companyFilesList = myQBAddin.myOpenCompanyFiles ' Calls Pwsh script, to return comma separated list
    OpenNames_array = Split(companyFilesList, ",") '
    
    ' 2. If QB_NAME not passed in, Select QB_NAME from List of open QB Company Files
    If QB_NAME = "" Then
        resultSuccess = Application.InputBox("SELECT A QB COMPANY FILE FROM OPEN FILES...", "WHERE IS THIS", _
                        companyFilesList, , , , , 2)
        If resultSuccess = "False" Then Exit Sub ' User Canceled
        ' Set Company Name
        QB_NAME = resultSuccess
        m = "You selected '" & QB_NAME & "'"
        MsgBox m
    End If
    
    ' 3. Make sure user selected QB_NAME in in the list List of open QB Company Files
    If IsStringInArray(QB_NAME, OpenNames_array) Then
        Debug.Print """" & QB_NAME & """ IS in QB OpenNames_array."
    Else
        m = """" & QB_NAME & """ was not found in QB OpenNames_array."
        Debug.Print m
        MsgBox m
        Exit Sub
    End If
    
    ' 2. Is BRC List Reporter Installed (Loaded)
    brcLoaded = myQBAddin.is_BRC_Addin_Installed("C:\myBRC\List Reporter\List Reporter.xla")
    If brcLoaded Then
        ' BRC List Reporter is Installed (Loaded)
        ' Run BRC 'Create List' Macro, select 'dest_sheetName' from List, click 'Create Report'
        m = "OK good!" & vbCrLf & vbCrLf & "BRC List Reporter IS loaded." & vbCrLf & vbCrLf
        m = m & "Open it in the ribbon and Click 'Create List Report'." & vbCrLf
        m = m & "Select '" & dest_sheetName & "' from List, click 'Create Report'." & vbCrLf
        MsgBox m
    Else
        m = "OOOPS...!" & vbCrLf & vbCrLf & "The BRC List Reporter is NOT loaded..." & vbCrLf & vbCrLf
        m = m & "Open the addins list.  Check the box 'List Reporter' and rerun this macro." & vbCrLf
        m = m & "Aborting..." & vbCrLf
        MsgBox m
        Exit Sub
    End If
    
    ' 3. Find the existing sheet in the destination workbook
    On Error Resume Next ' Turn off error handling
    If Not myQBAddin.SheetExists(dest_sheetName, dest_Wb) Then
        If Not myQBAddin.SheetExists("old_" & dest_sheetName, dest_Wb) Then
            Application.DisplayAlerts = prevState ' Turn alerts to prev state
    
            m = m & "Adding a template named 'old_" & dest_sheetName & "' from '" & ThisWorkbook.name & "' to make this work."
            If remove_Old Then m = m & vbCrLf & vbCrLf & "The sheet will removed when finished."
            Debug.Print m
            MsgBox m
            ' Add template ws from this workbook
            prevState = Application.DisplayAlerts
            Application.DisplayAlerts = False ' Turn off alerts to suppress name conflict warnings
            resultSuccess = myQBAddin.Copy_Sheet(dest_Wb, "old_" & dest_sheetName, ThisWorkbook, "BRCt_" & dest_sheetName)
            Application.DisplayAlerts = prevState ' Turn alerts to prev state
             
            If resultSuccess <> True Then
                m = "Failed to a ws named 'old_" & dest_sheetName & "' to '" & dest_Wb.name & "'." & vbCrLf & vbCrLf
                m = m & resultSuccess & vbCrLf & vbCrLf
                m = m & "Aborting..."
                MsgBox m, vbCritical, "BRC_List_Report_To_WB()"
                Exit Sub
            End If
        End If
    End If
    On Error GoTo 0 ' Turn on error handling again
    
    ' 4. Find the new workbook created by List Reporter
    ' function to select Which Open  file
    Set wb = whichOpen_WB_Dialog(dest_sheetName, dest_Wb)
    If wb Is Nothing Then Exit Sub ' error or user canceled
    
    
    ' 5. Confirm user's wb choice
    wb.Activate
    m = "Use this sheet for " & dest_sheetName & vbCrLf & vbCrLf
    m = m & "Click Yes to Continue."
    m = m & "Click No or Cancel to Abort and try again..."
    ans = MsgBox(m, vbYesNoCancel + vbQuestion, wb.name)
    If ans <> vbYes Then Exit Sub
        
    
    ' 6. Import BRC sheet into this workbook replacing existing dest_sheetName
    resultSuccess = myQBAddin.Import_BRC_List(dest_Wb, dest_sheetName, wb, "List Report")
    If resultSuccess <> True Then
        m = "Error occured on importing a worksheet! " & dest_Wb.name & vbCrLf & vbCrLf
        Debug.Print m
        MsgBox m, vbCritical, "Copy_List_Report_To_dest_Wb()..."
        Exit Sub
    End If
    
        
    ' 7. Save this workbook and confirm it saved.
    resultSuccess = myQBAddin.saveAndConfirm(dest_Wb)
    If resultSuccess <> True Then
        m = "Error occured on saving this workbook! " & dest_Wb.name & vbCrLf & vbCrLf
        Debug.Print m
        MsgBox m, vbCritical, "Copy_List_Report_To_dest_Wb()..."
    End If
    
    '
    ' 8. Now prompt user to close the List Reporter's  wb with out saving and NO Messages.
    If resultSuccess Then
        m = "Workbook '" & dest_Wb.name & "' is now saved!" & vbCrLf & vbCrLf
    Else
        m = ""
        Debug.Print m
        MsgBox m & vbCrLf & vbCrLf & "Aborting...", vbCritical, "BRC_List_Report_To_WB()"
        Exit Sub
    End If
    m = m & "Now closing with out saving the temp workbook '" & wb.name & "'." & vbCrLf & vbCrLf
    m = m & "Click Yes to Continue."
    m = m & "Click No or Cancel to NOT close workbook '" & wb.name & "'."
    ans = MsgBox(m, vbYesNoCancel + vbQuestion, wb.name)
    If ans <> vbYes Then Exit Sub
        
    
    On Error Resume Next ' resume next In case the workbook is already closed or invalid
        ' Set the Saved property to True to trick Excel into thinking
        ' the workbook is already saved, preventing "Do you want to save?" messages.
        wb.Saved = True
        ' Close the workbook. The SaveChanges argument is explicitly set
        ' to False, although the above line makes this redundant for suppression.
        wb.Close SaveChanges:=False
    On Error GoTo 0 ' Reset error handling

End Sub


'
'   Import_BRC_List
'
'    '  Move destsheet to 'old_' & destsheet
'    '  Confirm srcWB header matches 'old_' & destsheet
'    '  Copy srcSheet and Rename to destSheet in destWB
'    '  Format the new destSheet like the old_ destSheet
'
'       This function uses new AI generated code for copying and formatting
'           Vendor Lists
'           Chart of Accounts
'           Class Lists
'

Function Import_BRC_List(destWB As Workbook, destSheet As String, srcWb As Workbook, srcSheet As String)
Dim m As String

    '  Confirm destSheet exists
    If Not myQBAddin.SheetExists(destSheet, destWB) And _
       Not myQBAddin.SheetExists("old_" & destSheet, destWB) Then
        m = "Sheets '" & destSheet & "' and '" & "old_" & destSheet & "' missing in '" & destWB.name & "'.  Aborting..."
        Debug.Print m
        MsgBox m, , srcWb.name
        Import_BRC_List = m
        Exit Function ' Failed destSheet exist test
    End If
    
    '  Move dest sheet to 'old_' & dest sheet
    Call Move_wS_to_Old_wS(destWB, destSheet)

    '  Confirm new wb header matches
    If Not do_List_Hdrs_Match(destWB, srcWb, "old_" & destSheet) Then
        m = "Format of the new sheet in " & srcWb.name & " is incorrect..." & vbCrLf & vbCrLf
        m = m & "New Sheet should look like sheet '" & "old_" & destSheet & "' in this workbook." & vbCrLf
        m = m & "Aborting..." & vbCrLf
        MsgBox m, , srcWb.name
        Import_BRC_List = m
        Exit Function ' Failed header Match test
    End If
    
    '  Copy and Rename sheet to dest WB
    Import_BRC_List = myQBAddin.Copy_Sheet(destWB, destSheet, srcWb, srcSheet)
    If Import_BRC_List <> True Then Exit Function  ' Failed to copy sheet


    '  Format the new destSheet like the old_ destSheet
    Call Format_BRC_List(destWB, destSheet, destWB, "old_" & destSheet)

End Function ' End of Import_BRC_List function


'  Move_wS_to_Old_wS
'    '  Delete old_ + sheetname worksheet
'    '  Rename sheetname to old_ + sheetname
'
Sub Move_wS_to_Old_wS(wb As Workbook, sheetName As String)
   On Error Resume Next
    '  Delete Existing Sheet old_ + sheetname
    If myQBAddin.SheetExists("old_" & sheetName, wb) _
            And myQBAddin.SheetExists(sheetName, wb) Then
        ' delete the existing old_ sheet first
        Application.DisplayAlerts = False
        ' wb.Sheets("old_" & sheetName).Select  ' select for user confirmation
        wb.Sheets("old_" & sheetName).Delete  ' delete old_ ws first
        Application.DisplayAlerts = True
        '  Rename Existing Sheet to old_ + sheetname
        wb.Sheets(sheetName).name = "old_" & sheetName
    End If
   On Error GoTo 0
End Sub

' AI Generated
' Determine if Header of other workbook matches the header in sheetname
Function do_List_Hdrs_Match(dest_Wb As Workbook, ByVal OtherWb As Workbook, ByVal sheetName As String) As Boolean
    Dim wsCurrent As Worksheet
    Dim wsOther As Worksheet
    Dim lastColCurrent As Long
    Dim lastColOther As Long
    Dim colIndex As Long
    Dim m As String

    On Error Resume Next ' Handle potential errors like sheet not found
    Set wsCurrent = dest_Wb.Worksheets(sheetName)
    Set wsOther = OtherWb.Sheets(1) ' Reference the first sheet of the other workbook
    On Error GoTo 0 ' Turn off error handling

    If wsCurrent Is Nothing Then
        m = "Sheet not found in THIS workbook. '" & sheetName & "'   '" & dest_Wb.name & "'."
        Debug.Print m
        do_List_Hdrs_Match = False ' sheetname not found in this workbook. "
        Exit Function
    End If
    If wsOther Is Nothing Then
        m = "Sheets(1) not found in other workbook. '" & OtherWb.name & "'."
        Debug.Print m
        do_List_Hdrs_Match = False ' Sheets(1) not found in other workbook
        Exit Function
    End If

    ' Find the last used column in the first row for both sheets
    lastColCurrent = wsCurrent.Cells(1, wsCurrent.Columns.count).End(xlToLeft).Column
    lastColOther = wsOther.Cells(1, wsOther.Columns.count).End(xlToLeft).Column

    ' Determine the maximum column to compare up to
    Dim maxCol As Long
    'maxCol = IIf(lastColCurrent >= lastColOther, lastColCurrent, lastColOther)
    ' compare up to the last column in Otherwb
    maxCol = lastColOther  ' Use BRC List Reports (other wb) here because we may have added columns  at end in dest_Wb

    ' Loop through each cell in the first row and compare values
    For colIndex = 1 To maxCol
        If Not IsEmpty(wsCurrent.Cells(1, colIndex)) Or Not IsEmpty(wsOther.Cells(1, colIndex)) Then
             If wsCurrent.Cells(1, colIndex).Value <> wsOther.Cells(1, colIndex).Value Then
                do_List_Hdrs_Match = False ' Mismatch found
                Exit Function
            End If
        End If
    Next colIndex

    do_List_Hdrs_Match = True ' Headers Match. No mismatch found
End Function


Sub Format_BRC_List(destWB As Workbook, destSheet As String, srcWb As Workbook, srcSheet As String)
Dim st_col As Long
    
    
    Call myQBAddin.myFav_SheetSetup(destWB.Sheets(destSheet))
    ActiveWindow.Zoom = 100 ' Set the zoom level  100%
    
    ' Copy any addl header columns
    st_col = CopyAddlHdrColumnsAndFormat(destWB, destSheet, srcWb, srcSheet)
    
    ' Copy the data rows of the addl header columns starting in st_col
    Call CopyAddlDataColumnsAcrossSheets(st_col, destWB, destSheet, srcWb, srcSheet)
    
    ' Format Columns
    Call CopyColumnFormatting(destWB, destSheet, srcWb, srcSheet)
    
    ' Format Rows
    Call CopyRowFormatting(destWB, destSheet, srcWb, srcSheet)
End Sub


'Ai generated.
'   1. Copy additional column headers (if any) from srcsheetName ( ws2 )
'   and place after the last column in sheetname ( ws1 )
'   2. Copy and Paste Formatting of Entire row 1 srcWB to destWB
'
'   3. Return the column number of the starting column of the additonal columns (if any)
Function CopyAddlHdrColumnsAndFormat(destWB As Workbook, sheetName As String, _
            srcWb As Workbook, srcsheetName As String) As Long
    ' Declare variables
    Dim ws1 As Worksheet ' Destination
    Dim ws2 As Worksheet ' Source
    Dim startCol As Long
    Dim endCol As Long
    Dim targetRange As Range
    Dim m As String
    
    CopyAddlHdrColumnsAndFormat = 0 ' set  default
    
    ' Set references to the worksheets
    Set ws1 = destWB.Sheets(sheetName) ' Destination
    Set ws2 = srcWb.Sheets(srcsheetName) ' Source
    
    ' Disable screen updating for faster execution
    Application.ScreenUpdating = False
    
    ' 1. Find the last used column in Row 1 of ws1 PLUS ONE
    '    This is equivalent to selecting the last cell in the row and pressing Ctrl + Left
    startCol = ws1.Cells(1, ws1.Columns.count).End(xlToLeft).Column + 1
    
    ' Find the last used column in Row 1 of ws2
    endCol = ws2.Cells(1, ws2.Columns.count).End(xlToLeft).Column
    
    ' Check if the start column is a valid column number (e.g., if ws1 was empty)
    If startCol > ws1.Columns.count Then
        m = "Sheet1 row 1 is empty. Cannot determine a valid starting column for Sheet2."
        GoTo abortHere
    End If
    
    ' Set the target range on Sheet2 using the dynamic start and end column numbers
    ' The range will be from the calculated startCol to the calculated endCol in Row 1
    If endCol >= startCol Then
        Set targetRange = ws2.Range(ws2.Cells(1, startCol), ws2.Cells(1, endCol))
        'M = "Range set: " & targetRange.Address
        ' targetrange.parent.activate
        'targetRange.Select
        'MsgBox M
        
    ' 1. Copy columns from srcsheetName and place after the last column in sheetName
        ' Copy the column
        targetRange.Copy
        ' Paste into the next available column in destination (Sheet1)
        ws1.Cells(1, startCol).PasteSpecial xlPasteAll
        CopyAddlHdrColumnsAndFormat = startCol
    ElseIf endCol + 1 < startCol Then
        m = "CopyAddlHdrColumnsAndFormat() Error: The ending column is before the starting column. Adjust logic or data."
        Debug.Print m
        MsgBox m & vbCrLf & vbCrLf & "Aborting...", vbCritical, "CopyAddlHdrColumnsAndFormat():"
        GoTo abortHere
    Else
        m = "CopyAddlHdrColumnsAndFormat(): No additional columns in '" & srcsheetName & "' to append to '" & sheetName & "'."
        Debug.Print m
        ' continue with copy formatting
    End If
  
    ' 2. Copy and Paste Row 1 Formatting from srcsheetName to sheetName
    
    ' Copy the Format of the entire row 1 in Sheet2
    ws2.Rows(1).Copy
    
    ' Paste the Format to the entire row 1 in Sheet1
    ws1.Rows(1).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
abortHere:
    
    ' Clear the clipboard
    Application.CutCopyMode = False
    
    ' Re-enable screen updating
    Application.ScreenUpdating = True
    
    ' Optional: Autofit column widths in Sheet1 for better display
    'ws1.Columns.AutoFit
End Function



' Copy the data cells (ie. columns I added) from the original wsSource worksheet
'    to the end of each data row in the wsDestination
'
' AI generated
'
' Copies the contents (VALUES ONLY) of cells in row 2 wsSource,
' starting from StartCol through to the end column of the row
' and pastes them to the same Startcol in row 2 of wsDestination
' Then copies the cells  from row 2 and pastes(VALUES ONLY) down to the last row in wsDestination

Sub CopyAddlDataColumnsAcrossSheets(startCol As Long, destWB As Workbook, sheetName As String, _
            srcWb As Workbook, srcsheetName As String)
    ' Define variables for the worksheets
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    ' Define variables for the source and destination ranges
    Dim sourceRange As Range
    Dim destRange  As Range
    ' Define a variable for the last used column in the source sheet
    Dim lastCol As Long
    ' Define a variable for the last used row in the destination sheet
    Dim lastRowDest As Long
    ' Define the row index (2)
    Const secondRow As Long = 2

    ' if startCol is zero then nothing to do.
    If startCol = 0 Then Exit Sub
 
 
    ' Set references to the worksheets
    ' Ensure the sheet names match your actual sheet names in the workbook
    Set wsSource = srcWb.Sheets(srcsheetName)
    Set wsDest = destWB.Sheets(sheetName)

    ' Find the last used column in the source row (Row 1)
    lastCol = wsSource.Cells(1, wsSource.Columns.count).End(xlToLeft).Column

    ' Check if there is data to copy starting from the specified column
    If lastCol >= startCol Then
        ' Define the source range (Row 2, starting from startCol to the last used column)
        Set sourceRange = wsSource.Range(wsSource.Cells(secondRow, startCol), wsSource.Cells(secondRow, lastCol))

        ' Find the last used row in the destination sheet's column l
        ' This determines how many rows the source data needs to be pasted into
        lastRowDest = wsDest.Cells(wsDest.Rows.count, 1).End(xlUp).row
        
        ' Define the destination range (from Row 2 to the last used row in startCol)
        ' The Resize property ensures the destination range has the same number of columns as the source
        Set destRange = wsDest.Range(wsDest.Cells(2, startCol), wsDest.Cells(lastRowDest, lastCol))
 
        ' *** Paste VALUES ONLY (no formatting) ***
        destRange.Formula = sourceRange.Formula

        ' Optional: Clear the clipboard mode
        Application.CutCopyMode = False

    End If
End Sub

' AI generated

' Copy Column Formatting from SourceSheet to DestinationSheet
'
Sub CopyColumnFormatting(destWB As Workbook, sheetName As String, _
            srcWb As Workbook, srcsheetName As String)
    ' Declare variables for source and destination worksheets
    Dim SourceSheet As Worksheet
    Dim DestinationSheet As Worksheet
    ' Declare variable for the last used column number in the source sheet
    Dim lastCol As Long
    ' Declare loop counter
    Dim colIndex As Long
    
    ' Set the worksheet objects
    Set SourceSheet = srcWb.Sheets(srcsheetName)
    Set DestinationSheet = destWB.Sheets(sheetName)
    
    ' Disable screen updating for faster macro execution
    Application.ScreenUpdating = False
    
    ' Find the last used column in the source sheet (checks the first row, adjust if needed)
    lastCol = SourceSheet.Cells(1, SourceSheet.Columns.count).End(xlToLeft).Column
    
    ' Loop through each column that has data in the source sheet
    For colIndex = 1 To lastCol
        ' Copy the entire column in the source sheet
        SourceSheet.Columns(colIndex).Copy
        
        ' Paste only the formats to the corresponding column in the destination sheet
        DestinationSheet.Columns(colIndex).PasteSpecial Paste:=xlPasteFormats
        
        ' Optional: Also copy column widths for an exact match
        DestinationSheet.Columns(colIndex).PasteSpecial Paste:=xlPasteColumnWidths
    Next colIndex
    
    ' Clear the clipboard
    Application.CutCopyMode = False
    
    ' Re-enable screen updating
    Application.ScreenUpdating = True
    
    ' Confirmation message
    'MsgBox "Column formatting copied successfully from Sheet 2 to Sheet 1."
End Sub

' AI Generated

' Copy row formatting from row 2 of  source sheet
' Paste to all Data rows in destination sheet
'
Sub CopyRowFormatting(destWB As Workbook, sheetName As String, _
            srcWb As Workbook, srcsheetName As String)
    ' Define variables for worksheets and last row
    Dim wsSource As Worksheet
    Dim wsDestination As Worksheet
    Dim lastRow As Long
    
    ' Set references to the source and destination worksheets
    Set wsSource = srcWb.Sheets(srcsheetName)
    Set wsDestination = destWB.Sheets(sheetName)
    
    ' Find the last row with data in Sheet1 to determine the range for formatting
    ' This ensures only rows with existing data get formatted
    lastRow = wsDestination.Cells(wsDestination.Rows.count, 1).End(xlUp).row
    
    ' Copy the formatting from the entire row 2 of Sheet2
    wsSource.Rows(2).Copy
    
    ' Apply the copied formatting to the target range in Sheet1 (from row 2 to the last data row)
    ' xlPasteFormats ensures only the formatting is pasted
    With wsDestination.Range("A2:A" & lastRow).EntireRow
        .PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    End With
    
    ' Clear the clipboard
    Application.CutCopyMode = False
        
    'MsgBox "Formatting  ROWS from Sheet2 Row 2 applied to all rows after row 1 in Sheet1."
End Sub


' AI Generated

Function IsStringInArray(ByVal targetString As String, ByRef stringArray() As String) As Boolean
    Dim element As Variant
    IsStringInArray = False ' Default to False
    
    On Error GoTo ErrorHandler ' Handle case where the array might be empty or not initialized

    For Each element In stringArray
        ' Use StrComp for case-insensitive comparison (vbTextCompare)
        If StrComp(targetString, element, vbTextCompare) = 0 Then
            IsStringInArray = True
            Exit Function ' Exit the function as soon as a match is found
        End If
    Next element
    
    Exit Function ' Ensure normal exit if no match is found

ErrorHandler:
    ' Handle the error here, for example, if the array is empty
    If Err.Number = 9 Then ' Subscript out of range error (common for uninitialized arrays)
        Debug.Print "Error: The string array is empty or not initialized."
    Else
        Debug.Print "An unexpected error occurred: " & Err.Description
    End If
End Function


