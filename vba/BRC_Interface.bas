Attribute VB_Name = "BRC_Interface"
Option Explicit


Private Sub test_is_BRC_Addin_Installed()
    MsgBox is_BRC_Addin_Installed("C:\myBRC\List Reporter\List Reporter.xla")
End Sub

Function is_BRC_Addin_Installed(fullName As String) As Boolean
' Used by my BRC Exec and IIF Mod
Dim addinWB As Workbook, ff As Long, ErrNo As Long
Dim m As String, adName As String
Dim LastSlashPos As Integer
    
    ' Find the position of the last backslash
    LastSlashPos = InStrRev(fullName, "\")
    
    ' If a backslash is found, extract the part of the string after it
    If LastSlashPos > 0 Then
        adName = Mid(fullName, LastSlashPos + 1)
    Else
        ' If no backslash is found, the input is likely just the filename
        adName = fullName
    End If
    
    is_BRC_Addin_Installed = False
    On Error Resume Next
    
    ' See is Addin is in workbooks...
    Set addinWB = Workbooks(adName)
    If Err = 0 Then
        is_BRC_Addin_Installed = True ' Addin or WB Is Loaded
        Error 55  ' Pass 'File Already Open' Error back to calling function
        Exit Function
    End If
    
' Test if file open perhaps by another application
' https://stackoverflow.com/questions/9373082/detect-whether-excel-workbook-is-already-open
    Err.Clear
    ff = FreeFile()
    Open fullName For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    Select Case ErrNo
        Case 0:    is_BRC_Addin_Installed = False ' Not loaded
        Case 53:   ' File not found
        Case 70:   ' File opened by another application
                m = fullName & " ADDIN IS LOADED BY ANOTHER APP!! "
        Case 76:   ' Path not found
        Case Else: is_BRC_Addin_Installed = True ' Other error
    End Select
    If ErrNo <> 0 Then Debug.Print "is_BRC_List_Reporter_Installed(): " & m & Error, ErrNo  ' print error
    If m <> "" Then MsgBox m
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Create IIF from ACTIVEWORKBOOK and ACTIVE SHEET
'
' args: output_file - Make sure output_file is a FULL PATH NAME
'
' Uses: Addin C:\myBRC\IIF Transaction Creator\"
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function IIF_EXPORT_ActiveSheet(Optional output_file As String = "", Optional result As Boolean = True)
Dim ofile As String
    If output_file = "" Then
        ' o_file HERE OUT_PUT_FILE WILL BE ACTIVE WORKBOOK NAME & ".IIF"
        ' strips the '.xlsx' off ofile name
        ofile = Left(ActiveWorkbook.fullName, Len(ActiveWorkbook.fullName) - 5)
    Else
        ofile = output_file
    End If
    
    'Make sure a FULL PATH NAME IS USED for ofile
    IIF_EXPORT_ActiveSheet = brcExportFunction(ActiveWorkbook.fullName, ofile, ActiveSheet.name)
    If result Then MsgBox IIF_EXPORT_ActiveSheet, , "IIF_EXPORT_ActiveSheet"
End Function

Private Sub test_brcExportSub()
Dim o_file As String
    ' o_file SHOULD BE workbook name & ".IIF" & Date Time String BUT IT IS NOT!!!
    o_file = brcExportSub(ActiveWorkbook.fullName, ActiveWorkbook.fullName, ActiveSheet.name)
    MsgBox "Exported to '" & o_file & "'"
End Sub

'**********************************************************************
'
' brcExportSub
'  Creates IIF files by calling BRC IIF Creator Appliaciton
'  Must have Addin 'IIF Transaction Creator' installed for this to work
'  args: i_file and o_file MUST BE FULL PATH NAMES
'
'***********************************************************************
Function brcExportSub(i_file As String, o_file As String, Optional sheet As String = "Sheet1") As String
Dim myResult As Variant, m As String
Const path = "C:\myBRC\IIF Transaction Creator\"
Const brcname = "Transaction Creator for QB.xla" ' No single quotes
Const brcAddin = "'Transaction Creator for QB.xla'"
    If Not my_Addins_Mod.IsAddinLoaded(path, brcname) Then
        m = "Addin 'IIF Transaction Creator' is NOT LOADED." & vbCrLf & vbCrLf
        m = m & "You must open the Addins list and check the box for 'IIF Transaction Creator'." & vbCrLf & vbCrLf
        m = m & "Aborting..."
        MsgBox m, vbCritical, brcAddin
        Exit Function
    End If
    Debug.Print "brcExportSub: '" & i_file & "' '" & sheet & "' '" & o_file & "'"
    myResult = Application.Run(brcAddin & "!Export_To_IIF", i_file, sheet, o_file)
    If myResult(1) = True Then
        m = "IIF Export Success!! "
        brcExportSub = myResult(3)
    Else
        m = "IIF Export FAILED! "
    End If
    MsgBox m & vbCrLf & myResult(2) & vbCrLf & myResult(3)
End Function

Private Sub test_brcExportFunction()
    ' o_file SHOULD BE workbook name & ".IIF" & Date Time String BUT IT IS NOT!!!
    MsgBox brcExportFunction(ActiveWorkbook.fullName, ActiveWorkbook.fullName, ActiveSheet.name)
End Sub

'**********************************************************************
'
' brcExportFunction
'   Returns  file name of the LATEST IIF Created
'
'  Creates IIF files by calling BRC IIF Creator Appliaciton
'  Must have Addin 'IIF Transaction Creator' installed for this to work
'  args: i_file and o_file MUST BE FULL PATH NAMES
'
'************************************************************************
Function brcExportFunction(i_file As String, o_file As String, Optional sheet As String = "Sheet1") As String
Dim myResult As Variant, m As String
Const path = "C:\myBRC\IIF Transaction Creator\"
Const brcname = "Transaction Creator for QB.xla" ' No single quotes
Const brcAddin = "'Transaction Creator for QB.xla'"
Dim latest_IIF As String, opath As String, ofn As String
    latest_IIF = ""
    If Not my_Addins_Mod.IsAddinLoaded(path, brcname) Then
        m = "Addin 'IIF Transaction Creator' is NOT LOADED."
        brcExportFunction = m
        m = vbCrLf & vbCrLf & m & "You must open the Addins list and check the box for 'IIF Transaction Creator'." & vbCrLf & vbCrLf
        m = m & "Aborting..."
        MsgBox m, vbCritical, brcAddin
        Exit Function
    End If
    Debug.Print "brcExportFunction: '" & i_file & "' '" & sheet & "' '" & o_file & "'"
    myResult = Application.Run(brcAddin & "!Export_To_IIF", i_file, sheet, o_file)
    If myResult(1) = True Then
        ' Find and return the latest IIF matching file name in path
        opath = PathNameOnly(myResult(3))
        ofn = FilenameFromPath(myResult(3))
        ofn = Strip_Ext(ofn, ".IIF")
        latest_IIF = FullNameOfLatestSave(opath, ofn)
        
        m = "IIF Export Success! " & latest_IIF
    Else
        m = "IIF Export FAILED! "
    End If
'
'   On Error Resume Next
'   M = M & Join(myResult, ", ")
'   ' Release memory
'   Erase myResult
'   On Error GoTo 0

    Debug.Print m
    
    brcExportFunction = latest_IIF
End Function

Private Sub test_FullNameOfLatestSave()
    Dim fullFoundName As String
    
    ' Looks for any file in the folder starting with "Report_2026"
    fullFoundName = FullNameOfLatestSave("N:\BBM_OFFICE\BANK Statements\SCHW RP 373\", "ENT_SwRP 373.xlsm *.IIF")
    
    If fullFoundName <> "" Then
        MsgBox "Found file: " & fullFoundName
    Else
        MsgBox "No file found with that prefix."
    End If
End Sub


Function FullNameOfLatestSave(ByVal folderPath As String, ByVal filePrefix As String) As String
    Dim fileName As String
    Dim latestName As String
    Dim fileDate As Date
    Dim latestDate As Date
    
    ' Ensure the folder path ends with a backslash
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    ' Find the first matching file
    fileName = Dir(folderPath & filePrefix & "*")
    
    ' Loop through all matching files to find the newest one
    Do While fileName <> ""
        fileDate = FileDateTime(folderPath & fileName)
        
        ' If this file is newer (or it is the first one found), record it
        If fileDate > latestDate Then
            latestDate = fileDate
            latestName = fileName
        End If
        
        ' Get the next matching file
        fileName = Dir()
    Loop
    
    ' Return the result (will be blank if no match was found)
    FullNameOfLatestSave = latestName
End Function


