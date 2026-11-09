Attribute VB_Name = "ZIP_FILE_FOLDER_EMAIL"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public Sub ExportZip_VB()
    ' Prompts for temp folder and zip folder
    On Error Resume Next
     Call Zip_All_Files_in_Folder(ExportVisualBasicCode, "", ActiveWorkbook.name, True)
End Sub
Sub ExportZip_My_VB()
    ' No Prompts... and using my folders
    On Error Resume Next
    ' Empty temp folder with Kill
    Kill "N:\Dwight\Junk\VBA\*.*"
    Call Zip_All_Files_in_Folder(ExportVisualBasicCode("N:\Dwight\Junk\VBA", False), "N:\Dwight\VBA Stuff\My VBA BackUps", ActiveWorkbook.name, True)
End Sub
Sub Open_NotePadPlusPlus(Optional ByVal fp As String = "")
If fp <> "" Then fp = Chr(34) & fp & Chr(34)
Dim retval
    retval = Shell("C:\Program Files (x86)\Notepad++\notepad++.exe " & fp, vbNormalFocus)
End Sub

Private Sub test_Zip_File_Or_Files()
    Call Zip_File_Or_Files("N:\Dwight\Junk\")
End Sub
' https://www.rondebruin.nl/win/s7/win001.htm
' Browse to the folder you want and select the file or files
Sub Zip_File_Or_Files(Optional defPath As String = "")
    Dim strDate As String, sFName As String
    Dim oApp As Object, iCtr As Long, i As Integer
    Dim fname, vArr, FileNameZip

    defPath = BrowseToFolder(defPath)

    If defPath = "" Then MsgBox ("No Folder Chosen.     Exiting..."):   Exit Sub
    If Right(defPath, 1) <> "\" Then
        defPath = defPath & "\"
    End If

    strDate = Format(Now, " dd-mmm-yy h-mm-ss")
    FileNameZip = defPath & "MyFilesZip " & strDate & ".zip"

    'Browse to the file(s), use the Ctrl key to select more files
    fname = Application.GetOpenFilename(, _
            ButtonText:="CHOOSE FILES... ", MultiSelect:=True, Title:="Select the files you want to zip")
    If IsArray(fname) = False Then
        'do nothing
    Else
        'Create empty Zip File
        NewZip (FileNameZip)
        Set oApp = CreateObject("Shell.Application")
        i = 0
        For iCtr = LBound(fname) To UBound(fname)
            vArr = Split97(fname(iCtr), "\")
            sFName = vArr(UBound(vArr))
            If bIsBookOpen(sFName) Then
                MsgBox "You can't zip a file that is open!" & vbLf & _
                       "Please close it and try again: " & fname(iCtr)
            Else
                'Copy the file to the compressed folder
                i = i + 1
                oApp.Namespace(FileNameZip).CopyHere fname(iCtr)

                'Keep script waiting until Compressing is done
                On Error Resume Next
                Do Until oApp.Namespace(FileNameZip).Items.count = i
                    Application.Wait (Now + TimeValue("0:00:01"))
                Loop
                On Error GoTo 0
            End If
        Next iCtr

    MsgBox oApp.Namespace(FileNameZip).Items.count & " Items Zipped.   You find the zipfile here: " & FileNameZip
    End If
End Sub

Private Sub test_BrowseToFolder()
Dim f As String
    f = BrowseToFolder("N:\Dwight\Junk\VBA")
    MsgBox " You Chose   '" & f & "'"
End Sub
Function BrowseToFolder(Optional initialFolder As String = "") As String
Dim fd As FileDialog, vrtSelectedItem
Dim fileChosen As Integer 'the number of the button chosen
    If initialFolder = "" Then
        initialFolder = Application.DefaultFilePath
    End If
    If Right(initialFolder, 1) <> "\" Then
        initialFolder = initialFolder & "\"
    End If
    
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    'Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Choose a FOLDER"
        .InitialFileName = initialFolder  ' make sure to add the last "\"
        .InitialView = msoFileDialogViewDetails
        .ButtonName = "USE THIS FOLDER... "
        .Filters.Clear
        '.Filters.Add "Excel files", "*.xlsx" ' For msoFileDialogFilePicker Only
        '.Filters.Add "All files", "*.*" ' For msoFileDialogFilePicker Only
        'MsgBox .SelectedItems(1)
        fileChosen = .Show
        If fileChosen = -1 Then
            For Each vrtSelectedItem In .SelectedItems
            'vrtSelectedItem is a string that contains the path of each selected item.
            'Use any file I/O functions that you want to work with this path.
            'This example displays the path in a message box.
            'MsgBox "The path selected is: " & vrtSelectedItem
            BrowseToFolder = .SelectedItems(1)
            Next vrtSelectedItem
        End If
    End With
 'Set the object variable to Nothing.
 Set fd = Nothing
End Function
Private Sub test_mySaveAS()
Dim FN As String
    'fn = "N:\Dwight\Misc_Spreadsheet\ABCDE.xls"
    FN = ActiveWorkbook.fullName
    
    FN = mySaveAS(FN)
    If FN <> "" Then
        MsgBox "Success.  " & FN
    Else
        MsgBox "Failed!  "
    End If
End Sub
Private Sub test_CopyFileWithFSO()
    'Call CopyFileWithFSO("\\NAS\BBMDATA\TEMPLATE\AddIn Backups\myQBAddin.xlam 09Jan2022_10PM.BBM-ALIEN.Dwight", "C:\Users\dwight\AppData\Roaming\Microsoft\AddIns\", "myQBAddin.xlam.JUNK", True)
End Sub
'  SAFE GUARD METHOD of Copying a FILE from the internet - Awesome
' https://wellsr.com/vba/2018/excel/vba-copy-a-file-with-fso-copyfile/#safeguards
Sub CopyFileWithFSO(SourceFilePath As String, DestPath As String, Optional destName As String = "", Optional overwrite As Boolean = True)
' (1) copies one file from one folder to another folder with the VBA FileSystemObject
' (2) contains extensive error handling (safeguards)
' (3) requires a reference to the object library "Microsoft Scripting Runtime" under Options > Tools > References... in the Visual Basic Editor.

Dim blFileExists As Boolean, blSourceErr As Boolean
Dim strFileName As String, strSuccessMsg As String, strNewDestPath As String, strNewSourcePath As String
Dim fso As Scripting.FileSystemObject
Dim strErrMsg As String
    
    Set fso = New Scripting.FileSystemObject
    With fso
        strNewDestPath = .BuildPath(.GetAbsolutePathName(DestPath), "\")
        strFileName = .GetFileName(SourceFilePath)
    
        'check if the source file exists
        If Not .FileExists(SourceFilePath) Then
            
            ' check if the root drive was specified
            If .DriveExists(Left(SourceFilePath, 2)) Then
                blSourceErr = True
            
            ' the provided source path is incomplete
            ' build new path and ask the user if he accepts the suggestion
            Else
                strNewSourcePath = .BuildPath(.GetAbsolutePathName(SourceFilePath), "")
                If Not MsgBox("The source path " & Chr(34) & SourceFilePath & Chr(34) & _
                    " is incomplete. Will you accept the following suggestion: " _
                    & Chr(34) & strNewSourcePath & Chr(34) & "?", vbYesNo, "Confirm new source path") = vbYes Then _
                        blSourceErr = True
            End If
            
            ' error
            If blSourceErr Then _
                strErrMsg = "The source file," & Chr(34) & strFileName & Chr(34) & _
                        " does not exist, or the specified path to the file, " & Chr(34) & _
                        Replace(SourceFilePath, strFileName, "") & Chr(34) & " is incorrect."
        
        ' check if the destination folder already exists
        ElseIf Not .FolderExists(strNewDestPath) Then
        
            ' prompt the user if the destination folder should be created
            If MsgBox("The destination folder, " & Chr(34) & strNewDestPath & Chr(34) & ", does not exist. Do you want to create it?", vbYesNo, _
                "Create new folder?") = vbYes Then
                .CreateFolder (strNewDestPath)
            Else
                strErrMsg = "The destination folder could not be created."
            End If
        
        ' check if the file already exists in the destination folder
        Else
            If destName = "" Then destName = strFileName
            blFileExists = .FileExists(strNewDestPath & destName)
            If Not overwrite Then
                If blFileExists Then _
                    strErrMsg = "The file, " & Chr(34) & destName & Chr(34) & _
                        ", already exists in the destination folder, " & Chr(34) & _
                        strNewDestPath & Chr(34) & "."
            End If
        End If
        
        ' attempt to copy file
        If strErrMsg = vbNullString Then
            On Error Resume Next
            If strNewSourcePath = vbNullString Then strNewSourcePath = SourceFilePath
            Call .CopyFile(strNewSourcePath, strNewDestPath & destName, overwrite)  ' FSO.CopyFile ??
            If Err.Number <> 0 Then strErrMsg = "Run-time error " & Err.Number & Chr(10) & Err.Description
            On Error GoTo 0
        End If
        
        ' succesful copy
        If strErrMsg = vbNullString Then
            strSuccessMsg = "The file" & Chr(34) & strFileName & Chr(34) & " was copied to " & _
                Chr(34) & strNewDestPath & destName & Chr(34) & "."
            If blFileExists Then strSuccessMsg = strSuccessMsg & Chr(10) & _
                "(Note, the existing file in the destination folder was overwritten)."
            MsgBox strSuccessMsg, vbInformation, "File copied"
        
        ' error
        Else
            MsgBox strErrMsg, vbCritical, "Error!"
        End If
        
    End With

End Sub

' Saves active workbook to filename passed in.
' Uses BrowseToFolder if folder not passed in
Function mySaveAS(Optional ByVal FN As String = "") As String
Dim full_fname As String, fpath As String
Dim m As String, ans As Long
    If FN = "" Then
        full_fname = BrowseToFolder("N:\Dwight\Misc_Spreadsheet\")
        full_fname = full_fname & "\" & ActiveWorkbook.name
        If myFileExists(full_fname) Then
            m = full_fname & vbCrLf & vbCrLf
            m = m & "Overwrite this EXISTING file?"
            ans = MsgBox(m, vbYesNo, "mySaveAS(): SAVING FILE ALREADY EXISTS...")
            If ans <> vbYes Then MsgBox full_fname, , "ABORTING mySaveAS ": Exit Function
            If ActiveWorkbook.fullName = full_fname Then
                ' Save file and go to end.
                ActiveWorkbook.Save  '
                GoTo File_Update_Check
            End If
        End If
    Else
        If myFileExists(FN) Then
            m = FN & vbCrLf & vbCrLf
            m = m & "Overwrite this EXISTING file?"
            ans = MsgBox(m, vbYesNo, "mySaveAS(): SAVING FILE ALREADY EXISTS...")
            If ans <> vbYes Then MsgBox FN, , "ABORTING mySaveAS ": Exit Function
            full_fname = FN
            If ActiveWorkbook.fullName = full_fname Then
                ' Save file and go to end.
                ActiveWorkbook.Save  '
                GoTo File_Update_Check
            End If
        ElseIf FolderExists(FN) Then
            full_fname = BrowseToFolder(FN)
            If full_fname = "" Then MsgBox FN, , "ABORTING mySaveAS ": Exit Function
            full_fname = full_fname & "\" & ActiveWorkbook.name
        Else
            '? Folder or file passed in does not exist. What to do?
            fpath = PathNameOnly(FN)
            If FolderExists(fpath) Then
                m = FN
                ans = MsgBox(m, vbYesNoCancel, "SAVE TO THIS NAME?")
                If ans <> vbYes Then MsgBox FN, , "ABORTING mySaveAS ": Exit Function
                full_fname = FN
            Else
                m = FN & vbCrLf & vbCrLf
                m = m & "PATH DOES NOT EXIST!!   ABORTING!"
                MsgBox m, , "ABORTING mySaveAS ": Exit Function
            End If
        End If
    End If
    
    ActiveWorkbook.SaveAs fileName:=full_fname, AddToMru:=True, FileFormat:=xlOpenXMLWorkbook '    xlOpenXMLWorkbook (51) format = .xlsx
File_Update_Check:
    ' Check to see lastmodified date is updated
    If myFileUpdated(full_fname) Then
         mySaveAS = full_fname ' Success
    End If
End Function
Private Sub test_Zip_All_Files_in_Folder()
End Sub

'==============================================================================
' Function:     Zip_All_Files_in_Folder
' Purpose:      Zips all files in FolderName into a .zip file, saving the
'               result to defPath. Prompts for folders if not supplied.
' Notes:        Uses PowerShell's Compress-Archive instead of
'               Shell.Application.Namespace/CopyHere. The Shell "Compressed
'               Folder" extension was failing to register the zip as a valid
'               namespace even on a local drive (likely because a 3rd-party
'               zip tool has taken over the .zip file association, or the
'               shell extension is disabled by policy). Compress-Archive
'               uses .NET's zip library directly and works reliably on both
'               local and UNC destination paths with no shell dependency.
' Updated:      Replaced Shell.Application zip approach entirely with a
'               PowerShell Compress-Archive call; sanitized ZipName to
'               prevent double ".zip" extensions.
'==============================================================================
Function Zip_All_Files_in_Folder(FolderName As Variant, Optional defPath As String = "", Optional ZipName = "", Optional prompt = True) As String
    Dim strDate As String
    Dim FileNameZip As String
    Dim fso As Object
    Dim itemCount As Long
    Dim psCommand As String
    Dim shellCmd As String
    Dim wsh As Object
    Dim retCode As Long

    If defPath = "" Then
        defPath = BrowseToFolder(defPath)
    Else
        If prompt Then defPath = BrowseToFolder(defPath)
    End If
    If defPath = "" Then Exit Function
    If Right(defPath, 1) <> "\" Then defPath = defPath & "\"

    If FolderName = "" Then FolderName = BrowseToFolder(defPath)
    If Right(FolderName, 1) <> "\" Then
        FolderName = FolderName & "\"
    End If
    If LCase(FolderName) = LCase(defPath) Then
        MsgBox ("Origination and Destination Folders cannot be the same.")
        Exit Function
    End If

    If ZipName = "" Then ZipName = "MyFilesZip"
    ' Sanitize: strip a pre-existing ".zip" extension so we don't double it up
    If LCase(Right(ZipName, 4)) = ".zip" Then
        ZipName = Left(ZipName, Len(ZipName) - 4)
    End If
    ZipName = Trim(ZipName)

    strDate = Format(Now, " dd-mmm-yy h-mm-ss")
    FileNameZip = defPath & ZipName & " " & strDate & ".zip"

    ' If a zip with this exact name somehow already exists, remove it first
    ' (Compress-Archive -Force will overwrite, but this keeps things clean)
    On Error Resume Next
    If Len(Dir(FileNameZip)) > 0 Then Kill FileNameZip
    On Error GoTo 0

    ' Build the PowerShell command. Wrap paths in escaped double-quotes to
    ' handle spaces. "FolderName\*" zips the CONTENTS of the folder rather
    ' than the folder itself.
    psCommand = "Compress-Archive -Path '" & FolderName & "*' -DestinationPath '" & FileNameZip & "' -Force"

    shellCmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command """ & psCommand & """"

    Set wsh = CreateObject("WScript.Shell")
    ' Run synchronously (wait = True) and hidden (windowStyle = 0)
    retCode = wsh.Run(shellCmd, 0, True)

    If retCode <> 0 Then
        MsgBox "PowerShell Compress-Archive failed (exit code " & retCode & ") zipping:" & vbCrLf & FolderName, vbCritical
        Exit Function
    End If

    If Len(Dir(FileNameZip)) = 0 Then
        MsgBox "Compress-Archive reported success but no zip file was found at: " & FileNameZip, vbCritical
        Exit Function
    End If

    ' Count items that were zipped (count files in source folder, non-recursive
    ' to match what Compress-Archive -Path "Folder\*" grabs by default;
    ' adjust if your folder has subfolders you want included/excluded)
    Set fso = CreateObject("Scripting.FileSystemObject")
    itemCount = fso.GetFolder(FolderName).Files.count

    Zip_All_Files_in_Folder = FileNameZip

    If prompt Then
        MsgBox itemCount & " Items Zipped. " _
        & " You find the zipfile here: " & FileNameZip _
        , vbInformation, ZipName
    End If

    Debug.Print (itemCount & " Items Zipped. " & FileNameZip)
End Function



Sub NewZip(sPath)
'Create empty Zip File
'Changed by keepITcool Dec-12-2005
    If Len(Dir(sPath)) > 0 Then Kill sPath ' If ZipFile aleready exists Delete it
    Open sPath For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
End Sub
Function bIsBookOpen(ByRef szBookName As String) As Boolean
' Rob Bovey
    On Error Resume Next
    bIsBookOpen = Not (Application.Workbooks(szBookName) Is Nothing)
End Function
Function Split97(sStr As Variant, sdelim As String) As Variant
'Tom Ogilvy
    Split97 = Evaluate("{""" & _
                       Application.Substitute(sStr, sdelim, """,""") & """}")
End Function

Private Sub test_CreateEmptyFile()
    CreateEmptyFile "N:\Dwight\JUNK\tmp\testCreate\jnk.txt"
    ' createTextFile REQUIRES the folder to exist. Otherwise it errors out.
    'createTextFile "N:\Dwight\JUNK\tmp\testCreate\jnk.txt"
End Sub
' Creates empty file ONLY if nonexistant.
Sub CreateEmptyFile(fpath As String)
' Early Binding...  No Microsoft Scripting Runtime Needed
Dim fso As New FileSystemObject
Dim ts As TextStream
'On Error Resume Next
    ' Create ONLY if Non Existent
    If Not fso.FileExists(fpath) Then
        ' doesn't exist. so create
        If Not fso.FolderExists(fpath) Then
            ' folder doesn't exist, so create the folder
                  fso.CreateFolder PathNameOnly(fpath)
        End If
        ' create the file.
        Set ts = fso.createTextFile(fpath)
        'Close the file
        ts.Close
    End If
    'Clean up memory
    Set fso = Nothing
    Set ts = Nothing
End Sub


Sub Test_SelectFilePath()
    Dim result As String
    ' --- Choose a test pattern ---
    result = SelectFilePath("C:\Users\Dwight.YACKLEY-ALIENWA\Downloads\*.zip")
    'result = SelectFilePath("C:\tmp\*.*")
    'result = SelectFilePath("C:\tmp\*.xlsx")
    If result = "" Then
        MsgBox "No file selected (user cancelled).", vbExclamation
    Else
        MsgBox "Selected file:" & vbCrLf & result, vbInformation
    End If
End Sub
'==============================================================================
' Function : SelectFilePath
' Purpose  : Prompts the user to select a file, using a path pattern (e.g.
'            "C:\tmp\*.zip" or "C:\tmp\*.*") to set the initial folder and
'            file filter. Returns the full path of the selected file, or
'            an empty string ("") if the user cancels.
' Usage    : Dim f As String
'            f = SelectFilePath("C:\tmp\*.zip")
'            If f <> "" Then MsgBox f
'==============================================================================
Function SelectFilePath(ByVal PathPattern As String) As String

    Dim fd As FileDialog
    Dim folderPath As String
    Dim filePattern As String
    Dim lastSlash As Long

    ' Determine if pattern passed in resolves to a Folder
    ' If so, make sure "\" is placed at the end of the folder PathPattern
    If FileOrFolderTypePath(PathPattern) = 2 Then ' it's a folder
           If Right(PathPattern, 1) <> "\" Then PathPattern = PathPattern & "\"
    End If
    
    
    ' --- Split the pattern into folder path and file filter ---
    lastSlash = InStrRev(PathPattern, "\")

    If lastSlash > 0 Then
        folderPath = Left(PathPattern, lastSlash)      ' e.g. "C:\tmp\"
        filePattern = Mid(PathPattern, lastSlash + 1)   ' e.g. "*.zip"
    Else
        folderPath = ""
        filePattern = PathPattern
    End If

    If filePattern = "" Then filePattern = "*.*"

    ' --- Set up the file picker dialog ---
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = "Select a " & filePattern & " file."
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Filtered Files", filePattern
        .Filters.Add "All Files", "*.*"

        If folderPath <> "" Then
            If Len(Dir(folderPath, vbDirectory)) > 0 Then
                .InitialFileName = folderPath
            End If
        End If

        If .Show = -1 Then
            SelectFilePath = .SelectedItems(1)
        Else
            SelectFilePath = ""   ' user cancelled
        End If
    End With

End Function
Private Sub test_CleanTmpFolder()
  ' Note: Folder N:\Dwight\JUNK\tmp\tmpjnk is created ONLY IF N:\Dwight\JUNK\tmp exists
'    CleanTmpFolder ("N:\Dwight\JUNK\tmp\tmpjnk")
    CleanTmpFolder ("\\NAS\bbmdata\Dwight\JUNK\tmpZip")
End Sub

' AI Generated
Sub CleanTmpFolder(folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(folderPath) Then
        ClearFolderContents fso, folderPath
        Exit Sub
    End If

    Dim attempt As Long
    Dim errNum As Long
    Dim errDesc As String
    Const max_Attempts As Long = 5

    For attempt = 1 To max_Attempts
        On Error Resume Next
        Err.Clear
        fso.CreateFolder folderPath
        errNum = Err.Number
        errDesc = Err.Description
        On Error GoTo 0

        If errNum = 0 Then
            Exit Sub ' success
        End If

        If fso.FolderExists(folderPath) Then
            ' It exists now - NAS just caught up. Clear it and finish.
            ClearFolderContents fso, folderPath
            Exit Sub
        End If

        ' Not created, not found either - NAS is still settling. Wait and retry.
        Sleep 500 * attempt ' 0.5s, 1s, 1.5s, 2s, 2.5s
    Next attempt

    MsgBox "CleanTmpFolder error occurred: " & errDesc & _
        " (Error " & errNum & ")" & vbCrLf & _
        "Could not create '" & folderPath & "' after " & max_Attempts & " attempts.", vbCritical
End Sub

Private Sub ClearFolderContents(fso As Object, folderPath As String)
    Dim folder As Object, subFolder As Object, file As Object
    Set folder = fso.GetFolder(folderPath)
    For Each file In folder.Files
        file.Delete True
    Next file
    For Each subFolder In folder.SubFolders
        subFolder.Delete True
    Next subFolder
End Sub


'' Creates every missing folder in the path, one level at a time
'Private Sub EnsureParentFolders(fso As Object, folderPath As String)
'    Dim parts() As String
'    Dim currentPath As String
'    Dim i As Long
'    Dim cleanPath As String
'
'    cleanPath = folderPath
'    If Right(cleanPath, 1) = "\" Then cleanPath = Left(cleanPath, Len(cleanPath) - 1)
'
'    ' Split on \ but keep the leading \\server\share intact
'    parts = Split(cleanPath, "\")
'
'    ' Rebuild starting from \\server\share (parts(0)="" , parts(1)="" , parts(2)=server, parts(3)=share)
'    If Left(cleanPath, 2) = "\\" Then
'        currentPath = "\\" & parts(2) & "\" & parts(3)
'        For i = 4 To UBound(parts) - 1 ' stop before the final folder itself
'            currentPath = currentPath & "\" & parts(i)
'            If Not fso.FolderExists(currentPath) Then
'                On Error Resume Next
'                fso.CreateFolder currentPath
'                On Error GoTo 0
'            End If
'        Next i
'    Else
'        currentPath = parts(0)
'        For i = 1 To UBound(parts) - 1
'            currentPath = currentPath & "\" & parts(i)
'            If Not fso.FolderExists(currentPath) Then
'                On Error Resume Next
'                fso.CreateFolder currentPath
'                On Error GoTo 0
'            End If
'        Next i
'    End If
'End Sub

Private Sub test_CleanTmpFolderOLD_CODE()
    CleanTmpFolderOLD_CODE ("N:\Dwight\JUNK\tmp\tmpjnk")
End Sub
' THIS CODE WAS REPLACED with CleanTmpFolder
Sub CleanTmpFolderOLD_CODE(path As String)
' Does NOT remove subfolders
On Error Resume Next
    If Right(path, 1) <> "\" Then path = path & "\"
    ' Kill DOES NOT remove SUB Folders!
    Kill path & "*.*"
    TempDir (path) ' Create folder if it doesn't exist
End Sub
Function TempDir(path As String)
' Early Binding...  No Microsoft Scripting Runtime Needed
Dim fso As New FileSystemObject
    If Not fso.FolderExists(path) Then
    ' doesn't exist, so create the folder
              fso.CreateFolder path
    End If
    ' NO Return Variable PASSED BACK???
    Set fso = Nothing
End Function
Private Sub test_FileOrFolderExists()
    MsgBox ("FileOrFolderExists = " & FileOrFolderExists("N:\Dwight\K-1s\2019 K-1 Discus US Side BBM.pdf"))
End Sub


Function FileOrFolderExists(strFilePath As String) As Boolean
' Accepts wildcard characters
' ".\" does NOT work for DIR and other File Exist test methods
' DIR returns "file.pdf", "file.xls", "file.txt", etc when "file" is passed in
'
    Err.Clear
    On Error GoTo TestingErr
    If Dir(strFilePath, vbDirectory) <> "" Then
        FileOrFolderExists = True  ' Path exists (could be a file or folder)
    ElseIf Dir(strFilePath) Then
        FileOrFolderExists = True ' Path exists (it's a file)
    Else
        FileOrFolderExists = False  ' Path noes not exist
    End If
    Exit Function
TestingErr:
    ' This is a vailid exit
    'Debug.Print "FileOrFolderExists() Error: " & Err.Number & " " & Err.Description
End Function


'===============================================================================
' Function:    FileOrFolderTypePath
' Purpose:     Determines whether a path string refers to a file, a folder,
'              or nothing that currently exists.
'
' Parameters:  pathStr (String) - Full or relative path to check. Local paths
'              (C:\...) and UNC paths (\\server\share\...) are supported.
'              A trailing backslash is ignored.
'
' Returns:     Long
'                2 (PATH_IS_FOLDER) - Path exists and is a directory
'                1 (PATH_IS_FILE)   - Path exists and is a file
'                0 (PATH_NOT_FOUND) - Path is empty, does not exist, or is
'                                     inaccessible
'
' Notes:       Uses GetAttr, so no library references are required.
'              GetAttr raises an error for missing paths; that error is
'              trapped and reported as 0.
'              Directory detection uses a bitwise test because folders may
'              also carry other attributes (hidden, read-only, etc.).
'
' Example:     If FileOrFolderTypePath("C:\Temp") = PATH_IS_FOLDER Then ...
'===============================================================================
Private Sub tst_FileOrFolderTypePath()
Dim m As String, answer As Long
    answer = FileOrFolderTypePath(ActiveWorkbook.path)
    m = "PATH_NOT_FOUND = 0" & vbCrLf
    m = m & "PATH_IS_FILE = 1" & vbCrLf
    m = m & "PATH_IS_FOLDER = 2" & vbCrLf & vbCrLf
    m = m & "Answer is " & answer
    MsgBox m
End Sub
Function FileOrFolderTypePath(ByVal pathStr As String) As Long
    Dim attr As Long
 Const PATH_NOT_FOUND As Long = 0
 Const PATH_IS_FILE As Long = 1
 Const PATH_IS_FOLDER As Long = 2

    FileOrFolderTypePath = PATH_NOT_FOUND
    pathStr = Trim$(pathStr)
    If Len(pathStr) = 0 Then Exit Function

    ' Strip a trailing backslash (but keep root like "C:\")
    If Len(pathStr) > 3 And Right$(pathStr, 1) = "\" Then
        pathStr = Left$(pathStr, Len(pathStr) - 1)
    End If

    On Error Resume Next
    attr = GetAttr(pathStr)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function          ' doesn't exist or can't be accessed
    End If
    On Error GoTo 0

    If (attr And vbDirectory) = vbDirectory Then
        FileOrFolderTypePath = PATH_IS_FOLDER
    Else
        FileOrFolderTypePath = PATH_IS_FILE
    End If
End Function

Private Sub test_newFileOrFolderExists()
    ' Test a valid folder path with backslashes
    Debug.Print "Testing 'C:\Windows\System32': " & newFileOrFolderExists("C:\Windows\System32") ' Should return True
    ' Test a valid folder path with forward slashes
    Debug.Print "Testing 'C:/Program Files': " & newFileOrFolderExists("C:/Program Files") ' Should return True
    ' Test a valid file path
    Debug.Print "Testing a file: " & newFileOrFolderExists("C:\Windows\System32\notepad.exe") ' Should return True
    ' Test a valid UNC path
    Debug.Print "Testing a UNC path '\\localhost\c$\Windows' : " & newFileOrFolderExists("\\localhost\c$\Windows") ' Should return True (if accessible)
    ' Test a path that does not exist
    Debug.Print "Testing a fake path: " & newFileOrFolderExists("C:\Fake\Folder\Name") ' Should return False
    ' Test with a forward slash at the end
    Debug.Print "Testing 'C:/Windows/': " & newFileOrFolderExists("C:/Windows/") ' Should return True
End Sub

' AI generated - More robust than using 'Dir'
Function newFileOrFolderExists(ByVal sPath As String) As Boolean
'
' Returns True if the specified file or folder path exists, otherwise False.
' Handles paths with either backslashes (\) or forward slashes (/).
'
    Dim fso As Object
    Dim sNormalizedPath As String
    
    ' Check for empty string
    If Trim(sPath) = "" Then newFileOrFolderExists = False: Exit Function
        
    ' Replace all forward slashes with backslashes for consistency
    sNormalizedPath = Replace(sPath, "/", "\")
    
    ' Create a FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Use a robust check that handles both files and folders
    If fso.FolderExists(sNormalizedPath) Or fso.FileExists(sNormalizedPath) Then
        newFileOrFolderExists = True
    Else
        newFileOrFolderExists = False
    End If
    
    ' Clean up the object
    Set fso = Nothing
End Function

Private Sub test_FolderExists()
Dim FN As String
    FN = "N:\Dwight\K-1s"
    ' Will return FALSE when fn is a FILE
    'fn = "N:\Dwight\K-1s\my K-1 List.xlsm"
    MsgBox "FolderExists = " & FolderExists(FN), , FN
End Sub
Function FolderExists(strFolderPath As String) As Boolean
'no wildcard characters
'It works both, with '\' at the end and without.
    On Error Resume Next
    FolderExists = (GetAttr(strFolderPath) And vbDirectory) = vbDirectory
End Function
Private Sub test_myFileExists()
    MsgBox "myFileExists = " & myFileExists("N:\Dwight\K-1s\2019 K-1 Discus US Side BBM.pdf")
End Sub

Private Sub test_isFileOpen()
Dim ff As Long, ErrNo As Long, m As String

Const fp = "\\NAS\bbmdata\Dwight\JUNK\my_Addins_Mod.bas"
    On Error Resume Next
    Err.Clear
    ff = FreeFile()
    Open fp For Input Lock Read As #ff
    MsgBox isFileOpen(fp)
    Close ff
End Sub
' This FUNCTION NEEDS MORE TESTING
Function isFileOpen(fp As String) As Boolean
' Test if file open perhaps by another application
' https://stackoverflow.com/questions/9373082/detect-whether-excel-workbook-is-already-open
Dim ff As Long, ErrNo As Long, m As String
    
    On Error Resume Next
    Err.Clear
    ff = FreeFile()
    Open fp For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    Select Case ErrNo
        Case 0:     ' Not loaded
        Case 53:   ' File not found
        Case 70:   ' File opened by another application
                m = fp & " IS OPENED BY ANOTHER APP!!" & vbCrLf
                isFileOpen = True
        Case 76:   ' Path not found
       Case Else: ' Other error
    End Select
    Debug.Print m & " Error: " & ErrNo & " " & Err.Description  ' print error
    MsgBox m & " Error: " & ErrNo & " " & Err.Description
End Function


Private Sub test_myFileUpdated()
    'MsgBox "myFileUpdated = " & myFileUpdated("\\NAS\bbmdata\Dwight\K-1s\My K-1 List.xlsmAAA")
    MsgBox "myFileUpdated = " & myFileUpdated(ActiveWorkbook.fullName)
End Sub
' Function to Confirm if a File Save Really did work.
' Returns TRUE if fp was updated in the last 3 seconds
Function myFileUpdated(fp As String) As Boolean ' https://www.rondebruin.nl/win/s9/win003.htm
' Late Binding... Set reference to Microsoft Scripting Runtime" because we use use Late binding here
Dim fso As Object, timediff As Double, fileobject As Object
Dim lastModifiedDate As Date
    On Error Resume Next
    Set fso = CreateObject("scripting.filesystemobject")
    ' USE ANOTHER object...Set with GetFile from fso
    Set fileobject = fso.GetFile(fp)
    If fileobject Is Nothing Then Debug.Print "myFileUpdated() fileobject is NOTHING. " & fp: GoTo Abort_or_End
    'If Not fileobject.FileExists(fp) Then Debug.Print "myFileUpdated() File Doesn't Exist " & fp: GoTo Abort_or_End
    lastModifiedDate = fileobject.DateLastModified
    If lastModifiedDate = 0 Then GoTo Abort_or_End
    timediff = DateDiff("s", lastModifiedDate, Now)
    Debug.Print "myFileUpdated() " & fp & " timediff = " & Format(timediff, "#,##0.000")
    ' Was file updated within the last 3 seconds?
    If timediff <= 3 Then myFileUpdated = True
Abort_or_End:
    Set fso = Nothing
    Set fileobject = Nothing
End Function
Private Sub test_myDeleteExistFile()
    MsgBox "myDeleteExistFile = " & myDeleteExistFile("\\NAS\bbmdata\Dwight\JUNK\tmp")
End Sub
Function myDeleteExistFile(fp As String) As Boolean ' https://www.rondebruin.nl/win/s9/win003.htm
' Late Binding... Set reference to Microsoft Scripting Runtime" because we use use Late binding here
Dim fso As Object, m As String, ans As Long
    On Error Resume Next
    Set fso = CreateObject("scripting.filesystemobject")
    If fso Is Nothing Then Exit Function
    If Not fso.FileExists(fp) Then
        m = "myDeleteExistFile(): File does not exist. " & vbCrLf & vbCrLf & fp
        ans = MsgBox(m, vbOKCancel, "FILE DOES NOT EXIST!")
        Exit Function
    End If
    fso.DeleteFile fp
    Set fso = Nothing
    myDeleteExistFile = True
End Function
' https://stackoverflow.com/questions/1743328/how-to-extract-file-name-from-path
Function FilenameFromPath(ByVal strFullPath As String) As String

    FilenameFromPath = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))

End Function

' To extract the file name from a path in VBA that
' might contain both forward slashes (/) and backslashes (`\`),
' the most robust and recommended approach is to use
' the FileSystemObject. This method handles path parsing consistently
' regardless of the specific path separators used.
' Ai generated
Function getFileNameFromMixedPath(ByVal fullpath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    getFileNameFromMixedPath = fso.GetFileName(fullpath)
    Set fso = Nothing ' Release the object
End Function

Private Sub test_GetUserDocumentsPath()
Dim docsFolderPath As String, m As String
    docsFolderPath = GetUserDocumentsPath()
    ' Print the path to the Immediate Window (Ctrl+G in VBA Editor)
    m = "The user's current Documents folder path is: " & docsFolderPath
    Debug.Print m
    MsgBox m
End Sub

'  AI Generated
'  This method uses the SpecialFolders("MyDocuments") property, which correctly
'  identifies the CURRENT path even if the user has changed the default location in Environment
Function GetUserDocumentsPath() As String
Dim WshShell As Object
    Set WshShell = CreateObject("WScript.Shell")
    GetUserDocumentsPath = WshShell.SpecialFolders("MyDocuments")
    Set WshShell = Nothing
End Function

Private Sub test_BackupAllOpenWorkbooks()
    BackupAllOpenWorkbooks
End Sub

' AI Generated
'  Back up all open workbooks using SaveCopyAs
'  while keeping the original files open.
'  The code iterates through each open workbook,
'   creates a backup in a specified folder with a timestamp and
'   then continues without closing the original files.
'
Sub BackupAllOpenWorkbooks(Optional backupFolder As String = "")
Dim wb As Workbook
Dim backupFileName As String
Dim wbName As String
Dim fileExtension As String
Dim timeStamp As String
Dim fso As Object, m As String, ans As Long

    ' !!! IMPORTANT: Change this path to your desired backup location !!!
    If backupFolder = "" Then
       backupFolder = GetUserDocumentsPath() ' my function
    End If
    
    ' Initialize FileSystemObject for path verification
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Ensure the backup directory exists. If not, consider adding code to create it.
    If Not fso.FolderExists(backupFolder) Then
        m = "Path does not exist.  Create it?" & vbCrLf & vbCrLf
        m = m & backupFolder
        ans = MsgBox(m, vbYesNoCancel, "BackupAllOpenWorkbooks()")
        If ans <> vbYes Then Exit Sub
        ' Create the Folder
        On Error Resume Next
        fso.CreateFolder backupFolder
        If Err.Number <> 0 Then
            MsgBox "BackupAllOpenWorkbooks(): Failed to create folder: " & backupFolder, vbCritical
            Exit Sub
        End If
        On Error GoTo 0
    End If
    
    ' Ensure folder path ends with a backslash
    If Right(backupFolder, 1) <> "\" Then backupFolder = backupFolder & "\"
    
    ' Format the timestamp to be file-name friendly (avoiding characters like / or :)
    timeStamp = Format(Now, "yyyy-mm-dd_HH-MM-SS") ' HH-MM-SS is 24 Hour format
    
    ' Loop through all open workbooks and NOT add-ins!!
    For Each wb In Application.Workbooks
        ' Skip Add-ins and workbooks that haven't been saved at least once
        If wb.path <> "" And Not wb.IsAddin Then
            wbName = fso.GetBaseName(wb.fullName) ' not wb.name? interesting... failsafe way maybe?
            backupFileName = backupFolder & wbName & "_Backup_" & timeStamp & "." & fso.GetExtensionName(wb.name)
            
            On Error Resume Next
            ' Save snapshot while keeping original open
            wb.SaveCopyAs backupFileName
            
            ' Proactive Code: Verify file actually exists after save
            If fso.FileExists(backupFileName) Then
                Debug.Print "BackupAllOpenWorkbooks(): Successfully backed up: " & backupFileName
            Else
                Debug.Print "BackupAllOpenWorkbooks(): FAILED to backup: " & backupFileName & " (Error: " & Err.Description & ")"
            End If
            On Error GoTo 0
        End If
    Next wb
End Sub


Private Sub test_CreateOutlook_Mail()
    Call Create_Mail("Dwight@bbmincorporated.com", "Hello", "This is the Body", "")
End Sub
' Create an Email Using Outlook
' References Microsoft Outlook Object Reference Library
' https://www.rondebruin.nl/win/s7/win001.htm
' https://www.automateexcel.com/vba/send-emails-outlook/
Sub Create_Mail(strTo As String, strSubject As String, strBody As String, strAttach As String)
'Dim OutApp As Object, OutMail As Object  ' used when replacing early bining with late binding
Dim OutApp As Outlook.Application, OutMail As Outlook.MailItem

        Set OutApp = New Outlook.Application
        Set OutMail = OutApp.CreateItem(olMailItem)

        On Error Resume Next
        With OutMail
            .To = strTo
            .CC = ""
            .BCC = ""
            .Subject = strSubject
            .Body = strBody
            If strAttach <> "" Then .Attachments.Add strAttach
            '.Send   'or use .Display
            .Display
        End With
        On Error GoTo 0
End Sub


' My function. Used in managing Outlook Junk Mail Address Lists
' Create an empty file Only if Not existent
Sub createTextFile(fp As String, Optional nonExistOnly As Boolean = False)
Dim ff As Long
    If nonExistOnly Then
        If Not FileOrFolderExists(fp) Then
            ff = FreeFile
            Open fp For Output As #ff: Close #ff
        End If
    Else  ' Make the file empty
        ff = FreeFile
        Open fp For Output As #ff: Close #ff
    End If
End Sub

Sub ClearTextFile(file As String)
Dim iFileNum As Integer
    iFileNum = FreeFile
    If file = "" Then Exit Sub '
    Open file For Output As iFileNum ' Create empty File
    Close #iFileNum
End Sub
Private Sub tst_AppendTextFile()
Dim fpath As String
    fpath = BrowseToFolder("N:\Dwight\Junk")
    fpath = fpath & "\junkText.txt"
    MsgBox " You Chose   '" & fpath & "'"
    AppendTextFile fpath, "ABCD"
    AppendTextFile fpath, "1234"
End Sub
Sub AppendTextFile(file As String, appString As String)
Dim oFileNum As Integer
    oFileNum = FreeFile
    If file = "" Then Exit Sub '
    Open file For Append As oFileNum ' Open File for Append
    ' Append data to the file
    Print #oFileNum, appString
    Close #oFileNum
End Sub

'  USES A BATCH FILE TO WORK --- NOT VERY PRACTICAL
Sub Make_ReadOnly(fp As String)
Dim m As String, ans As Long
        m = fp & vbCrLf & vbCrLf
        m = m & "Do you want to make this file Read Only now?" & vbCrLf & vbCrLf
        ans = MsgBox(m, vbQuestion + vbYesNoCancel, fp)
        If ans = vbYes Then
        ' ReadOnly.bat is a bat file to change a file attribute to Read Only
            m = "ReadOnly " & fp
            ' "/S /K" Continues without waiting
            Call Shell("cmd.exe /S /K" & m, vbNormalFocus)
        End If
End Sub
Private Sub test_myFileCopy()
    ' Copy all *.csv files
    Call myFileCopy("C:\Users\Public\Documents\DoorKing\DoorKing Account Manager\ACCT0001\*.csv", _
        "N:\DoorKing\.", True) ' True = overwrite
End Sub
' https://stackoverflow.com/questions/16943003/vba-to-copy-a-file-from-one-directory-to-another
Function myFileCopy(source As String, destination As String, Optional overwrite As Boolean = False)
Dim fso As Object, m As String, ans As Long
    On Error GoTo error_Out
    Err.Clear
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    Debug.Print "myFile Copy(): " & source & "   " & destination & " " & overwrite
    Call fso.CopyFile(source, destination, overwrite)
    GoTo Done
error_Out:
    m = "myQBAddin.myFileCopy():" & vbCrLf
    m = m & vbCrLf & "Error: " & Err.Number & " " & Err.Description & vbCrLf
    m = m & vbCrLf & "s = " & source & vbCrLf
    m = m & vbCrLf & "d = " & destination
    ans = MsgBox(m, vbInformation, "Error Copying...")
Done:
    Set fso = Nothing
End Function
' http://www.vbaexpress.com/forum/showthread.php?46362-Sort-contents-in-a-text-file
' Example Using Shell script to sort a text file
' Shell "Sort x:\sort\sort.txt /O x:\sort\sort_out.txt", vbHide
Sub my_Shell_script(s As String, focus_id As Long)
    's = "this is -  ""quote"""
   ' MsgBox s   ' Use for testing double quotes "" within strings
    Call Shell(s, focus_id)
    
    ' need to wait here for shell command to complete to output
    Application.Wait (Now + TimeValue("0:00:01"))
End Sub

' *****************************************************************************
'
'   The Folowing functions are duplicated in my_Addin's Mod as PRIVATE
'   This allows for  my_Addin's mod independence from this addin
'
' *****************************************************************************


 Function myFileExists(fp As String) As Boolean ' https://www.rondebruin.nl/win/s9/win003.htm
' Late Binding... Set reference to Microsoft Scripting Runtime" because we use use Late binding here
Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("scripting.filesystemobject")
    If fso Is Nothing Then Debug.Print "myFileExists() FSO is NOTHING!": Exit Function
' HAVING TROUBLE HERE! Sometimes returns FALSE for existing files!
    myFileExists = fso.FileExists(fp)
    Set fso = Nothing
End Function

Private Sub test_file_ext()
    MsgBox File_Ext(ThisWorkbook)
End Sub
' Returns .extention of WB passed in, plus the last '.' in a string
 Function File_Ext(wb As Workbook) As String
    File_Ext = LCase(Right(wb.fullName, Len(wb.fullName) - InStrRev(wb.fullName, ".") + 1))
End Function
Private Sub test_strip_ext()
    MsgBox Strip_Ext(ThisWorkbook.name)
End Sub
' Strips any and ALL ext 'x' strings at end of file name
 Function Strip_Ext(fname As String, Optional X As String = ".xlam") As String
    Strip_Ext = fname
    Do While LCase(Right(Strip_Ext, Len(X))) = LCase(X)
        Strip_Ext = LCase(Left(Strip_Ext, Len(Strip_Ext) - Len(X)))
    Loop
End Function
 Function FileNameOnly(pname) As String
' Returns the filename from a path/filename string
Dim temp As Variant
    If pname = "" Then Exit Function
    temp = Split(pname, Application.PathSeparator)
    FileNameOnly = temp(UBound(temp))
End Function
Private Sub test_OtherWorkbooksOpen()
    MsgBox (OtherWorkbooksOpen(""))
    MsgBox (OtherWorkbooksOpen(ThisWorkbook.name))
End Sub
 Function OtherWorkbooksOpen(myWorkbookName As String) As String
Dim i As Long, wbct As Long
    wbct = 1
    For i = 1 To Workbooks.count
        If LCase(Workbooks(i).name) = LCase(myWorkbookName) Then GoTo nextOne
        If LCase(Workbooks(i).name) Like LCase("PERSONAL.x*") Then GoTo nextOne
        OtherWorkbooksOpen = OtherWorkbooksOpen & wbct & ": " & Workbooks(i).name & vbCrLf
        wbct = wbct + 1
nextOne:
    Next
    If OtherWorkbooksOpen <> "" Then OtherWorkbooksOpen = _
        "Other Workboooks Are Open:" & vbCrLf & vbCrLf & OtherWorkbooksOpen
End Function

 Function MyDocsPath() As String
    If LCase(Environ("Username")) = "dwight" Then
        MyDocsPath = "\\NAS\BBMDATA\dwight\QUICKBOOKS Work area\" ' use for my work environ
    Else
        MyDocsPath = Environ("USERPROFILE") & "\My Documents\"
    End If
End Function
 Function PathNameOnly(pname) As String
' Returns the Path from a path/ string
Dim temp As Variant
Dim fname As String
Dim l As Integer
    l = Len(pname)
    temp = Split(pname, Application.PathSeparator)
     fname = temp(UBound(temp))
    l = l - Len(fname) ' leave off filename
    PathNameOnly = Left(pname, l)
    ' MsgBox "PathNameOnly(): result " & PathNameOnly
End Function

' Returns a one letter string with drive letter
Private Sub test_GetDriveLetter()
    Debug.Print ("Drive = " & GetDriveLetter("N:\abc")) ' N
    Debug.Print ("Drive = " & GetDriveLetter(PathNameOnly(ThisWorkbook.fullName))) ' C
    Debug.Print ("Drive = " & GetDriveLetter("\\NAS\BBMDATA")) ' Fails
    Debug.Print ("Drive = " & GetDriveLetter("\\NAS\BBMDATA\DWIGHT")) ' Fails
End Sub
' Returns a one letter string with drive letter
Function GetDriveLetter(drvpath) As String
' Returns NULL if drive letter Not Exist
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/driveletter-property
Dim fs, d
    On Error Resume Next
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(fs.GetDriveName(drvpath))
    GetDriveLetter = d.driveLetter
End Function
Private Sub test_fso_GetDrive()
Dim fso, d
Set fso = CreateObject("Scripting.FileSystemObject")
Set d = fso.GetDrive("C:") ' MUST BE LIKE be C, C:, or C:\, But not anything longer!
 
'Now you can use the Drive object to get drive properties such as below
Debug.Print d.AvailableSpace 'Available space on drive
Debug.Print d.driveLetter   ' Result: "C"
Debug.Print d.drivetype     ' Result 2
Debug.Print d.FileSystem    ' Result probably NTFS
Debug.Print d.FreeSpace     ' A large number
Debug.Print d.IsReady       ' True
Debug.Print d.RootFolder    ' Result: "C:\"
Debug.Print d.SerialNumber  ' A large number
Debug.Print d.ShareName     ' nothing
Debug.Print d.TotalSize     ' a LARGE NUMBER
Debug.Print d.VolumeName    ' Result "OS"
End Sub
' Not a very useful function...
Private Sub test_GetDriveType()
Dim t
    On Error Resume Next
    Debug.Print "Drive Type = " & GetDriveType("C")                 ' 2
    Debug.Print "Drive Type = " & GetDriveType("N:")                ' 3
    Debug.Print "Drive Type = " & GetDriveType("\\NAS\BBMDATA\")    ' 3
    Debug.Print "Drive Type = " & GetDriveType("\\NAS\")    ' FAILS -1
    Debug.Print "Drive Type = " & GetDriveType("N:\dwight") ' FAILS -1
End Sub
' Not a very useful function...
 Function GetDriveType(drvpath) As Long
Dim fs, d, s, t
' Returns 0: "Unknown", 1: "Removable", 2: "Fixed", 3: "Network", 4: "CD-ROM", 5: "RAM Disk"
    On Error Resume Next
    GetDriveType = -1 ' default for errors set HERE!
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(drvpath)
    If d.drivetype >= 0 Then GetDriveType = d.drivetype
End Function
' Not a very useful function...
Private Sub test_ShowDriveType()
    Debug.Print ShowDriveType("c:\")            'Result = Drive letter C Fixed
    Debug.Print ShowDriveType("\\nas\\bbmdata") ' Invalid Drive
    Debug.Print ShowDriveType("c:\users")       ' Invalid Drive
    Debug.Print ShowDriveType(PathNameOnly(ThisWorkbook.fullName)) ' Invalid Drive
    Debug.Print ShowDriveType(ThisWorkbook.fullName)    ' Invalid Drive
End Sub
' Not a very useful function...
 Function ShowDriveType(drvpath)
Dim fs, d, t
    On Error Resume Next
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(drvpath)
    ShowDriveType = "Invalid Drive letter in '" & drvpath & "'" ' Default message
    Select Case d.drivetype
        Case 0: t = "Unknown"
        Case 1: t = "Removable"
        Case 2: t = "Fixed"
        Case 3: t = "Network"
        Case 4: t = "CD-ROM"
        Case 5: t = "RAM Disk"
    End Select
    ShowDriveType = "Drive letter " & d.driveLetter & " " & t
End Function

' REMOVED FROM PUBLIC DY 1-11-26
Sub Show_Excel_Version()
    MsgBox Get_Excel_Version(), vbInformation
End Sub
 Function Get_Excel_Version() As String
    Dim version As Variant, excel As String
    'Excel 97 = 8
    'Excel 2000 = 9
    'Excel 2002 = 10
    'Excel 2003 = 11
    'Excel 2007 = 12
    'Excel 2010 = 14
    'Excel 2013 = 15
    'Excel 2016 = 16
    version = Val(Application.version)
    'version = 17 ' for testing purposes
    'If version < 14 Then
    '    excel = "Older than Excel 2010!... Which is too old to work!!"
    'Else
    '    excel = "Newer than Excel 2016!"
    'End If
    Select Case version
        Case 8:
            excel = "Excel 97"
            excel = excel & "...  which is older than Excel 2010 and too old to work here!!"
        Case 9:
            excel = "Excel 2000"
            excel = excel & "...  which is older than Excel 2010 and too old to work here!!"
        Case 10:
            excel = "Excel 2002"
            excel = excel & "...  which is older than Excel 2010 and too old to work here!!"
        Case 11:
            excel = "Excel 2003"
            excel = excel & "...  which is older than Excel 2010 and too old to work here!!"
        Case 12:
            excel = "Excel 2007"
            excel = excel & "...  which is older than Excel 2010 and too old to work here!!"
        Case 14:
            excel = "Excel 2010"
        Case 15:
            excel = "Excel 2013"
        Case 16:
            excel = "Excel 2016"
        Case Else:
            excel = "Unknown Excel version!!!  version = " & version
    End Select
    If version > 16 Then excel = "Newer than Excel 2016!"
    Get_Excel_Version = "This Excel version is '" & version & "   which is... " & excel
End Function


