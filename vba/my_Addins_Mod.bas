Attribute VB_Name = "my_Addins_Mod"
Option Explicit
Option Private Module
 Const ADDINAME As String = "myQBAddin.xlam" ' The name for this addin

' By default, subs and functions not explicitly declared as Private are Public by default.
' So, to prevent conflicts with other projects, using Option Private Module here.
' Then, this module can be copied and used in another project with out nameing conflicts.
' Option Private Module makes public members within this module PRIVATE TO THIS PROJECT.

' This module can be used with any Excel Add-In for Add-In management.
 Const NETWORK_ADDIN_FOLDER As String = "\\NAS\BBMDATA\TEMPLATE\AddIns_Shared\" ' Folder with updates and backups
 Const NETWORK_ADDIN_BACKUPS As String = "\\NAS\BBMDATA\TEMPLATE\AddIn_Backups\" ' Folder with updates and backups
 Const RETURN_SUCCESS = 0
 Const RETURN_FAILED = True
 Const RETURN_CANCELED = 2
 Const HELP_MESSAGE As String = "Use the following commands to use this Add-In:" & vbCrLf & _
"   Ctrl + A    Accounts with balances that are" & vbCrLf & _
"                                not allowed for a particular Class using MyClassType" & _
"                                  in the Class List sheet." & vbCrLf & _
"   Ctrl + N    To find if any Negative numbers in a sheet." & vbCrLf & _
"   Ctrl + C    To Clear all color filled cells on a sheet." & vbCrLf & _
"   Ctrl + Shift + C    To Clear color filled cells AND all Comments on a sheet." & vbCrLf & _
"   Ctrl + H    To Display Help Message." & vbCrLf & _
"   Ctrl + W    To see which version of Excel you are running." & vbCrLf & _
"   Ctrl + Shift + V to view Workbook and Worksheets for this Addin)."
 
 Dim DummyWorkbookName As String '
 Dim INSTALL_ADDIN_UPDATE As Boolean ' used at wb opening to flag user to install newer versions of the addin

' For referencing the constant in other Modules
Function GetMyADDINNAME() As String
    GetMyADDINNAME = ADDINAME
End Function
' For referencing the constant in other Modules
Function GetMyNETWORK_ADDIN_FOLDER() As String
    GetMyNETWORK_ADDIN_FOLDER = NETWORK_ADDIN_FOLDER
End Function
' For referencing the constant in other Modules
Function GetMyNETWORK_ADDIN_BACKUPS() As String
    GetMyNETWORK_ADDIN_BACKUPS = NETWORK_ADDIN_BACKUPS
End Function
' For referencing the variable in other Modules
Function GetMyINSTALL_ADDIN_UPDATE() As Boolean
    GetMyINSTALL_ADDIN_UPDATE = INSTALL_ADDIN_UPDATE
End Function
' For setting the variable from other Modules
Sub SetMyINSTALL_ADDIN_UPDATE(new_state As Boolean)
    INSTALL_ADDIN_UPDATE = new_state
End Sub
Sub Display_Help_Message()
    Dim ans As Integer
    ans = MsgBox(HELP_MESSAGE, vbInformation, ThisWorkbook.name)
End Sub
Sub ViewHideAddin()
    If ThisWorkbook.IsAddin = True Then
         View_Addin
    Else
         Hide_Addin
    End If
End Sub
' REMOVED FROM PUBLIC DY 1-11-26
Sub View_Addin(Optional verbose = False) 'Use from any Open workbook  Ctrl + Shift + V
Dim wb As Workbook, ans As Long
    If verbose Then _
    ans = MsgBox("About to VIEW Worksheets of Add-In '" & ThisWorkbook.fullName & "'!!!", vbInformation, ThisWorkbook.name)
    ThisWorkbook.IsAddin = False
End Sub
' REMOVED FROM PUBLIC DY 1-11-26
Sub Hide_Addin(Optional verbose = False) 'Use from any Open workbook Ctrl + Shift + H
Dim wb As Workbook, ans As Long
    If verbose Then _
    ans = MsgBox("About to HIDE Worksheets of Add-In '" & ThisWorkbook.fullName & "'!!!", vbInformation, ThisWorkbook.name)
    ThisWorkbook.IsAddin = True
End Sub
' REMOVED FROM PUBLIC DY 1-11-26
Sub Save_Addin() 'Use from any Open workbook  Ctrl + Shift + S
    Dim wb As Workbook
    Dim ans As Long, m As String
    Dim ai As AddIn, FN As String
    Dim addin_in_list As Boolean, ret As Long, i As Long
    On Error Resume Next
    ' Check for duplicate in ADDINS First
    If Duplicate_Addin Then
        Exit Sub ' Abort
    End If
    ' CONFIRM THE NAME OF THIS WORKBOOK IS THIS ADDIN NAME
    ' This can happen when add-in is copied or renamed or Saveas in the Add-in's folder.
    If ThisWorkbook.name <> ADDINAME Then
        m = "WARNING!!   The Addin NAME has CHANGED!" & vbNewLine
        m = m & vbNewLine & "ADDIN NAME = '" & ADDINAME & "'"
        m = m & vbNewLine & "This Workbook NAME = '" & ThisWorkbook.name & "'" & vbNewLine
        m = m & vbNewLine & "Do you want to continue?..  Enter Yes to Continue."
        m = m & vbNewLine & "No or Cancel to Abort. "
        Debug.Print m
        ans = MsgBox(m, vbCritical + vbYesNoCancel, ThisWorkbook.fullName)
        If ans <> vbYes Then Exit Sub ' Abort
    End If
    
    For Each ai In AddIns
        If LCase(ai.fullName) = LCase(ThisWorkbook.fullName) Then
           ' You are editing the Addin file in the Add-ins List
           addin_in_list = True ' Is in Add-Ins
' BACK UP COPY OF OLD ADDIN FIRST
           ret = BkUp_File(ai.fullName, NETWORK_ADDIN_BACKUPS, "." & Environ("computername") & "." & Environ("username"), False)   ' backup old copy first
           Exit For
        End If
    Next
      
    Err.Clear ' reset before next file operation
    If addin_in_list Then
        Set wb = Application.Workbooks(ai.name)
        If wb Is Nothing Then
            m = "Critical Error! " & Err & ai.fullName & _
                " file is NOT in the Application.Workbooks! " & vbCrLf
            Debug.Print m
            m = vbCrLf & m & " Aborting Add-in Save!"
            ans = MsgBox(m, vbCritical, ThisWorkbook.fullName)  '
            Exit Sub
        End If
    End If  ' An Addin_in_list


' NOW SAVE ADDIN or WORKBOOK
        FN = ThisWorkbook.fullName
        ret = my_Addins_Mod.SaveAddin(ThisWorkbook, False)    ' Fail Safe method
        If ret <> RETURN_SUCCESS Then Exit Sub ' If not success, don't publish
        
' NOW PUBLISH ADDIN TO NETWORK FOLDER
        'm = " Publishing Saved Add-in to Network Addin Folder!" & vbCrLf & _
        '    "   Folder = '" & NETWORK_ADDIN_FOLDER & "'"
        'ans = MsgBox(m, vbOKOnly + vbExclamation, ThisWorkbook.name)  '
        ret = BkUp_File(ThisWorkbook.fullName, NETWORK_ADDIN_FOLDER, ".Saved", False, "Published ")
End Sub

' ***********************************************************
'
' The following code is useful for DEBUGGING AND CRACKING ADDINS
'
' ***********************************************************

' I found a LOT of Junk in my AI list when I Run this macro.
' For any Addins with FALSE in the immediate window, the file doesn't exist.
' You can OPEN the ADDINS LIST in EXCEL and CHECK the box associated that Name
' This will prompt the system to ask if you want it removed from the list.
'
' If .exe or other files exist in the same folder and the addin's,
' those FiLES WILL BE ADDED INTO the Addins object List.
' So, BE CARRFUL what files you leave in a addins folder!
' FYI, I put htis marco in a separate workbook called DumpAIs_testing.xlsm

Private Sub Dump_AIs()
Dim ai As AddIn, i As Long, file_exist As Boolean
    i = 0
    For Each ai In AddIns
        i = i + 1
        file_exist = A_myFileExists(ai.fullName)
        
        Debug.Print i & " " & file_exist & " " & ai.name & "    '" & ai.fullName & "'"
    Next ai
End Sub


' CHECK ADDIN LIST for ANY duplicated ai names
Private Sub CheckAddinListForAnyDuplicates()
    Dim ai As AddIn, ai2 As AddIn
    Dim i As Long, ans As Long, m As String
    For Each ai In AddIns
        i = 0
        For Each ai2 In AddIns
            If A_Strip_Ext(LCase(A_FileNameOnly(ai2.fullName))) = A_Strip_Ext(LCase(A_FileNameOnly(ai.fullName))) Then
               i = i + 1
               If i > 2 Then
                      'CheckAddinListForAnyDuplicates = True
                      m = "DUPLICATE ADDIN ERROR!....." & vbCrLf & vbCrLf _
                      & "'" & ai.name & "' IS IN THE ADDIN LIST " & i & " TIMES!!!"
                      Debug.Print m
                      Debug.Print "ai = " & ai.fullName & " ai2 = " & ai2.fullName
                      'ans = MsgBox(M, vbCritical, ai.name)
               End If
            Else ' Skip other Addins
               'Debug.Print (i & ". Skipping... " & ai.FullName)
            End If
        Next ai2
    Next ai
End Sub

' CHECK ADDIN LIST for Duplicate Names of THIS workbook
Private Function Duplicate_Addin() As Boolean
    Dim ai As AddIn
    Dim i As Long, ans As Long, m As String
    i = 0
    For Each ai In AddIns
        If A_Strip_Ext(LCase(A_FileNameOnly(ai.fullName))) = A_Strip_Ext(LCase(A_FileNameOnly(ThisWorkbook.name))) Then
           i = i + 1
           If i > 1 Then
                  Duplicate_Addin = True
                  m = "DUPLICATE ADDIN ERROR!....." & vbCrLf & vbCrLf _
                  & "'" & ai.name & "' IS IN THE ADDIN LIST " & i & " TIMES!!!"
                  Debug.Print m
                  ans = MsgBox(m, vbCritical, ThisWorkbook.name)
           End If
        Else ' Skip other Addins
           'Debug.Print (i & ". Skipping... " & ai.FullName)
        End If
    Next
End Function

' Useful while CRACKING BRC and Other Projects
Private Sub SaveASAnyAddin()
Dim ai As AddIn
    Set ai = NumberFromAddinList
    Call SaveASAnyAddin_(ai, "")
End Sub

' Allows user to select an Addin by inputting number
' from a numbered List if INSTALLED Addins
Private Function NumberFromAddinList() As AddIn
Dim m As String, answer As Variant, ans As Long
Dim ai As AddIn
    Set NumberFromAddinList = Nothing  'default
    
    ' Dump all INSTALLED AIs to Immediate window
    Dump_Installed_AIs
    
    ' Select addin number from debug.print
    answer = Application.InputBox("Look in Immediate Window for list of loaded Addins." & _
            vbCrLf & "Enter the number of the one to Select." & vbCrLf, _
    "SELECTING AN ADDIN", _
    "Enter a number", , , , , 1) 'Type = 1 is number , 2 is string
    'Debug.Print answer
    
    If answer <= 0 Or answer = "False" Then ' User Canceled
        Exit Function
    End If
    
    Set NumberFromAddinList = AI_Num(CLng(answer))
            
    If NumberFromAddinList Is Nothing Then
        m = answer & " is an invalid selection." & vbCrLf & vbCrLf
        m = m & "No Addin found matching number '" & answer & "'." & vbCrLf & vbCrLf
        m = m & "Aborting..."
        MsgBox m, , "SELECTING AN ADDIN"
        Exit Function
     End If
    
    m = "You selected: " & answer & vbCrLf & vbCrLf
    m = m & NumberFromAddinList.fullName & vbCrLf & vbCrLf
    m = m & "Correct?" & vbCrLf & "Enter Yes to continue.    No, Cancel to Cancel."
        
    ans = MsgBox(m, vbYesNoCancel + vbQuestion, "SELECTING AN ADDIN")
    If ans <> vbYes Then ' User canceled
        Set NumberFromAddinList = Nothing
        Exit Function
    End If
    
    
End Function
' Numbered List of Running AI's
Private Sub Dump_Installed_AIs()
Dim ai As AddIn, i As Long, m As String
    For Each ai In AddIns
        If ai.Installed Then
            i = i + 1
            m = ""
            Debug.Print i & " " & ai.fullName & m
        End If
    Next ai
End Sub
' Return AI from the Numbered List of Running AI's
Private Function AI_Num(num As Long) As AddIn
Dim ai As AddIn, i As Long
    For Each ai In AddIns
        If ai.Installed Then
            i = i + 1
            Debug.Print i & " " & ai.fullName
            If i = num Then
                Set AI_Num = ai
                Debug.Print "AI_Num = " & i & " " & ai.fullName
                Exit Function
            End If
        End If
    Next ai
End Function

' The ai passed in is Saved AS to the FOLDER defined as fn.
' User is prompted for a folder if fn passed is "".
' User is also prompted to add 'CRACKED_' to the front of the ai.name
' Otherwise the ai.name is used as the saved file name.
'
Private Sub SaveASAnyAddin_(ai As AddIn, Optional FN As String = "")
Dim ans As Integer, wb As Workbook, prev_state As Boolean
Dim ext As String, m As String
    If ai Is Nothing Then
        MsgBox "SaveASAnytAddin_(): ai is nothing.  Exiting..."
        Exit Sub ' error
    End If
    If ai.Installed = False Then ' If false, the add-in is NOT installed (and therefore, not loaded)
            m = "Error: " & ai.name & "Addin NOT loaded!  Can't Save it...":  Debug.Print m
            MsgBox m
            Exit Sub
    End If
    
    If FN = "" Then ' Get the path and name to save as.
        FN = A_BrowseToFolder
    End If
    If FN = "" Then
        MsgBox "SaveASAnytAddin_(): User Canceled.   Exiting..."
        Exit Sub ' User canceled
    End If
    If Right(FN, 1) <> "\" Then
        FN = FN & "\"
    End If
    
    m = "Add 'CRACKED_' to the front of the file name?"
    ans = MsgBox(m, vbQuestion + vbYesNo, "SaveASAnyAddin_")
    If ans = vbYes Then
        FN = FN & "CRACKED_" & ai.name
    Else
        FN = FN & ai.name
    End If
   
    On Error Resume Next

' SaveAs for ai HERE
    Set wb = Workbooks(ai.name) ' MUST to switch to Workbook!
    ans = my_Addins_Mod.SaveASAddin(wb, FN, True)

    ai.IsAddin = prev_state  ' reset to view add-in worksheets again
    MsgBox "NOTE! You should exit and restart Excel now to be safe."
End Sub
Private Sub RemoveUninstalledAddinsFromList()
Dim i As Long, ai As AddIn, m As String, ans As Long

    On Error Resume Next ' Ignore the error that occurs when Excel can't find the file
    i = 0
    For Each ai In AddIns
        ' Attempt to set the Installed property to False
        ' This action, on a missing file, will trigger Excel's
        ' internal prompt to remove it from the list
        If ai.Installed = True Then
            ans = MsgBox(" Set " & ai.name & " = False?", vbYesNo + vbQuestion, ai.name)
            If ans = vbYes Then
             i = i + 1
                m = m & i & ". Setting " & ai.name & " = False" & vbCrLf
                Application.AddIns(ai.name).Installed = False
            End If
        End If
    Next
    m = m & vbCrLf & "Addins = " & AddIns.count
    Debug.Print m
    
    On Error GoTo 0 ' Resume normal error handling
    
End Sub
Private Sub ShowInstalledAddins() ' for debugging purposes
Dim i As Long, ai As AddIn, m As String
    i = 0
    For Each ai In AddIns
      If ai.Installed Then
         i = i + 1
         m = m & i & ": " & ai.name & vbTab & vbTab & ai.Installed & vbCrLf
      End If
    Next
    m = m & vbCrLf & " Total Addins = " & AddIns.count & "  Installed = " & i
    Debug.Print m
    MsgBox m
End Sub
Private Sub ShowAddins() ' for debugging purposes
Dim i As Long, ai As AddIn, m As String
    i = 1
    For Each ai In AddIns
        m = m & i & ": " & ai.name & vbTab & vbTab & ai.Installed & vbCrLf
        i = i + 1
    Next
    m = m & vbCrLf & "Addins = " & AddIns.count
    Debug.Print m
    MsgBox m
End Sub


Private Function SaveAddin(wb As Workbook, Optional prompt As Boolean = True) As Long ' Fail Safe method...
Dim ans As Integer
Dim ext As String, m As String, prev_state As Boolean, FN As String
    On Error Resume Next
    ext = LCase(my_Addins_Mod.A_File_Ext(wb))
    If wb.ReadOnly Then
    ' This ADDIN Should NOT Be READ ONLY
        m = "This File is READ ONLY!!'" & wb.name & "'" & vbCrLf & vbCrLf
        m = m & "Should NEVER get here...." & vbCrLf & vbCrLf
        m = m & "wb = " & FN & vbCrLf & vbCrLf
        m = m & "Save file with '.SAVED' extension?" & vbCrLf
        m = m & "   OK - Save file with '.SAVED' extension." & vbCrLf
        m = m & "   Cancel - Abort file Save." & vbCrLf
        ans = MsgBox(m, vbOKCancel + vbQuestion, ThisWorkbook.name)
        If ans = vbOK Then
            SaveAddin = SaveASAddin(wb, wb.fullName & ".SAVED", False)
            If SaveAddin = RETURN_SUCCESS Then INSTALL_ADDIN_UPDATE = True
            Exit Function
        End If
        m = "ABORTING SAVE..." & vbCrLf
        ans = MsgBox(m, vbExclamation + vbOKOnly, ThisWorkbook.name)
        SaveAddin = RETURN_CANCELED
        Exit Function
    End If 'read only addin

' NOT a READ ONLY ADDIN Here
    If ext = ".saved" Then
        FN = wb.fullName
        m = "File has a '.SAVED' extension!!  '" & wb.name & "'" & vbCrLf & vbCrLf
        m = m & "SAVE ANYWAY?..." & vbCrLf
        m = m & "wb = " & FN & vbCrLf & vbCrLf
        ans = MsgBox(m, vbOKCancel + vbQuestion, ThisWorkbook.name)
        If ans <> vbOK Then SaveAddin = RETURN_CANCELED: Exit Function
    End If
    ans = vbOK ' default
    
    ' CONFIRM THE NAME OF THIS WORKBOOK IS THIS ADDIN NAME
    ' This can happen when add-in is copied or renamed or Saveas in the Add-in's folder.
    If ThisWorkbook.name <> ADDINAME Then
        m = " Saving to..." & vbCrLf & vbCrLf & wb.name & vbCrLf & vbCrLf
        m = m & "If name NOT RIGHT, then CTRL + BREAK!"
        Debug.Print m
        ans = MsgBox(m, vbOKOnly, ThisWorkbook.fullName)
    End If

    If ans <> vbOK Then _
      SaveAddin = RETURN_CANCELED: Exit Function

' THE FAIL CHECK SAVE of an ADDIN WORKBOOK
    prev_state = wb.IsAddin ' save previous state
    wb.IsAddin = True  ' set to enable add-in file save
    Err.Clear
    wb.Save
    If Err <> 0 Then
        m = "ERROR: " & Err & "...  FILE NOT SAVED!!" & vbCrLf
        m = m & wb.fullName & vbCrLf
        m = m & Err.Description & vbCrLf
        m = m & A_Get_Excel_Version()
        ans = MsgBox(m, vbCritical, ThisWorkbook.fullName)
        Debug.Print (m)
        SaveAddin = RETURN_FAILED
    Else
        ' check Last Modifed Date to make sure it updated < X seconds ago
        If (Now - LastModDate(wb.fullName)) > TimeValue("0:00:03") Then
            m = "ERROR: " & Err & "... FILE NOT UPDATED!! " & vbCrLf
            m = m & wb.fullName & vbCrLf
            ans = MsgBox(m, vbCritical, ThisWorkbook.fullName)
            Debug.Print (m)
            SaveAddin = RETURN_FAILED
        Else
            ans = MsgBox("Addin Saved!  " & "  " & wb.fullName, vbInformation, ThisWorkbook.name)
            wb.Saved = True ' Prevents Excel "Save this Workbook?" prompt
            SaveAddin = RETURN_SUCCESS
       End If
    End If
    wb.IsAddin = prev_state  ' reset to view add-in worksheets again
End Function
Private Function SaveASAddin(wb As Workbook, SaveAsName As String, Optional prompt As Boolean = True) As Long   ' Fail Safe method...
Dim ans As Integer
Dim ext As String, m As String, prev_state As Boolean
Dim FN As String, fnModDate As Date
    On Error Resume Next
    ext = LCase(A_File_Ext(wb))
    prev_state = wb.IsAddin ' save previous state
    ans = vbOK ' default
    If SaveAsName = "" Then SaveASAddin = RETURN_FAILED: Exit Function
    FN = SaveAsName
    
    If prompt Then
        m = " Workbook is a " & ext & " file and we are about to save it." & vbCrLf & vbCrLf
        m = m & " Saving to..." & vbCrLf & vbCrLf & FN
        ans = MsgBox(m, vbOKCancel + vbQuestion, ThisWorkbook.fullName)
    End If
    If ans = vbOK Then
        wb.IsAddin = True  ' set to enable add-in file save
' THE SAVE_AS of ADDIN WORKBOOK
        Err.Clear
        Application.DisplayAlerts = False
        ' xlOpenXMLAddIn (55)  is for ".xlam" files
        wb.SaveAs fileName:=FN, FileFormat:=xlOpenXMLAddIn, ReadOnlyRecommended:=False, CreateBackup:=False
        Application.DisplayAlerts = True
        If Err = 0 Then
            ' check Last Modifed Date to make sure it updated < X seconds ago
            fnModDate = LastModDate(FN)
            If (Now - fnModDate) < TimeValue("0:00:03") Then
                ans = MsgBox("Addin Saved!  " & "  " & FN, vbInformation, ThisWorkbook.name)
                wb.Saved = True ' Prevents Excel "Save this Workbook?" prompt
                SaveASAddin = RETURN_SUCCESS
            Else
                m = "SAVE AS ERROR: " & Err & "... FILE NOT SAVED!! " & vbCrLf
                m = m & FN & vbCrLf
                ans = MsgBox(m, vbCritical, ThisWorkbook.name)
                Debug.Print (m)
                SaveASAddin = RETURN_FAILED
            End If
        Else
            m = "SAVE AS ERROR: " & Err & "...  FILE NOT SAVED!!" & vbCrLf
            m = m & wb.name & vbCrLf & vbCrLf
            m = m & Err.Description & vbCrLf & vbCrLf
            m = m & wb.fullName
            ans = MsgBox(m, vbCritical, ThisWorkbook.name)
            Debug.Print (m)
            SaveASAddin = RETURN_FAILED
        End If
        wb.IsAddin = prev_state  ' reset to view add-in worksheets again
    Else
        SaveASAddin = RETURN_CANCELED: Exit Function
    End If 'ans = vbOK
End Function
Private Sub test_MsgFileDateFileLen()
    MsgBox (MsgFileDateFileLen(ThisWorkbook.fullName))
End Sub
Private Function MsgFileDateFileLen(FN As String) As String ' for debugging purposes
    MsgFileDateFileLen = "   File Date = " & FileDateTime(FN) & "   File Len = " & FileLen(FN)
End Function
Private Sub test_IsAddinLoaded()
Dim s As String, ans As Long
'Const path As String = "\\NAS\BBMDATA\TEMPLATE1\AddIn Backups\"
Const path As String = "C:\Users\Dwight\"
'Const fn As String = "Test.txt"
Const FN As String = "ntuser.dat.LOG1"
'Const fn As String = "Book2.xls"
    'MsgBox Application.UserLibraryPath
    MsgBox (FN & "   IsAddinLoaded = " & my_Addins_Mod.IsAddinLoaded(path, FN) & _
        "   Error " & Err)
End Sub
Private Sub test2_IsAddinLoaded()
Dim s As String, ans As Long
Dim FN As String
Dim path As String
    path = Application.UserLibraryPath
    FN = ThisWorkbook.name
    MsgBox (FN & "   IsAddinLoaded = " & my_Addins_Mod.IsAddinLoaded(path, FN) & _
        "   Error " & Err)
End Sub
Private Sub test3_IsAddinLoaded()
    '  should always = True
    MsgBox my_Addins_Mod.IsAddinLoaded(ThisWorkbook.path, ThisWorkbook.name)
End Sub
Private Sub test4_IsAddinLoaded()
Const path = "C:\myBRC\IIF Transaction Creator\"
Const brcAddin = "Transaction Creator for QB.xla"
    MsgBox my_Addins_Mod.IsAddinLoaded(path, brcAddin)
End Sub
' Used by my BRC Exec and IIF Mod
Function IsAddinLoaded(path As String, adName As String) As Boolean  '
Dim addinWB As Workbook, ff As Long, ErrNo As Long
Dim m As String, fullName As String
    IsAddinLoaded = False
    On Error Resume Next
    If Right(path, 1) <> "\" Then
        fullName = path & "\" & adName
    Else
        fullName = path & adName
    End If
    
    ' See is Addin is in workbooks...
    Set addinWB = Workbooks(adName)
    If Err = 0 Then
        IsAddinLoaded = True ' Addin or WB Is Loaded
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
        Case 0:    IsAddinLoaded = False ' Not loaded
        Case 53:   ' File not found
        Case 70:   ' File opened by another application
                m = fullName & " ADDIN IS LOADED BY ANOTHER APP!! "
        Case 76:   ' Path not found
        Case Else: IsAddinLoaded = True ' Other error
    End Select
    If ErrNo <> 0 Then Debug.Print "IsAddinLoaded(): " & m & Error, ErrNo  ' print error
    If m <> "" Then MsgBox m
End Function

Private Sub test_Ensure_myHelperAddinIsLoaded()
    MsgBox Ensure_myHelperAddinIsLoaded("myQBAddin.xlam")
End Sub
'
' This makes sure and referenced addin gets loaded first!
'
' Call this function from workbook_Open..
'
Function Ensure_myHelperAddinIsLoaded(helperAddinName As String) As Boolean
Dim ai As AddIn
Dim helperPath As String
    ' Check if already installed
    For Each ai In Application.AddIns
        If LCase(ai.name) = LCase(helperAddinName) Then
            If ai.Installed = False Then ai.Installed = True
            Ensure_myHelperAddinIsLoaded = True
            Exit Function
        End If
    Next ai

    ' If not found, Open it directly from a known path
    helperPath = Application.UserLibraryPath & "/" & helperAddinName   ' <<< your path

    If Dir(helperPath) <> "" Then
        Application.Workbooks.Open helperPath
        Ensure_myHelperAddinIsLoaded = True
        Exit Function
    End If

    ' Failed
    Ensure_myHelperAddinIsLoaded = False

End Function


'
' BKUp_File()
'   sfullName   - The full path of source file to be backed up
'   aPath       - The folder to backup file
'   ext         - The string added to end of saved file name
'   prompt      - True for User Promts
'   action      - String included for Prompt and Error Messages (BackUp, Saved, etc)
' Called from This workbook before close
Private Function BkUp_File(sfullName As String, aPath As String, ext As String, Optional prompt As Boolean = True, Optional action As String = "Backup ") As Long
    Dim fso As Object
    Dim ans As Integer, m As String
    Dim backupFullName As String, sname
    Dim Thisdate As Date, BkupDate As Date
    Dim BkupLen As Long
    On Error Resume Next
    BkUp_File = RETURN_FAILED ' default
    If sfullName = "" Then Exit Function
    sname = A_FileNameOnly(sfullName)
    
    ' Place for Addins  on  Network
    backupFullName = aPath & sname & "_" & SaveDateTime() & ext
    
    Thisdate = FileDateTime(sfullName)
    BkupDate = LastModDate(backupFullName)
    If BkupDate = 0 Then
        BkupLen = 0
    Else
        BkupLen = FileLen(backupFullName)
    End If
    If prompt And BkupLen > 0 Then
        m = action & "  File already exists in folder..." & vbCrLf
        m = m & "Folder = '" & aPath & "'" & vbCrLf
        m = m & "File = '" & A_FileNameOnly(backupFullName) & "'" & vbCrLf
        m = m & vbNewLine & "    Yes - To Overwrite. " '
        m = m & vbNewLine & "    No  - NOT Overwrite. "
        m = m & vbNewLine & "    Cancel - To Cancel File Save. "
        ans = MsgBox(m, vbYesNoCancel + vbQuestion, sfullName)
        If ans = vbYes Then ans = vbOK ' Ok to back up
        Select Case ans
            Case vbNo
                backupFullName = aPath & sname & "_" & SaveDateTime(True) & ext
            Case vbCancel
                BkUp_File = RETURN_CANCELED: Exit Function
        End Select
    End If
      
    'https://stackoverflow.com/questions/16943003/vba-to-copy-a-file-from-one-directory-to-another
    Err.Clear
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    'Debug.Print "sfullName = " & sfullName
    'Debug.Print "bkup_Name = " & backupFullName
    Call fso.CopyFile(sfullName, backupFullName, True) ' Save back up copy
    If Err = 0 Then
        If A_myFileExists(backupFullName) Then
' SUCCESS:
        m = sname & " " & action & " File Saved to...  " & backupFullName ' Success
            BkUp_File = RETURN_SUCCESS
            Debug.Print m
            ans = MsgBox(m, vbInformation, sname)
        Else
        m = sname & " ERROR: " & Err & "...FILE BACKUP FAILED!!  " & backupFullName
            Debug.Print m
            ans = MsgBox(m, vbCritical, backupFullName)
        End If
    Else
        m = "ERROR: " & Err & "...FILE BACKUP FAILED!!  " & backupFullName
        Debug.Print m
        m = m & vbCrLf & Err.Description & vbCrLf
        ans = MsgBox(m, vbExclamation, "ThisWB = " & sname)
    End If
    Set fso = Nothing ' release memory
End Function
' Called from this workbook on open
Function CheckForWbUpdates(wb As Variant) As String
Dim newestWbFullName As String, m As String
    If Not A_myFileExists(wb.fullName) Then Exit Function
    
    newestWbFullName = newestFileInPath(NETWORK_ADDIN_FOLDER, wb.name & "*.*")
    
    If newestWbFullName = "" Then Exit Function ' No file exists
    
    If FileDateTime(newestWbFullName) _
        > FileDateTime(wb.fullName) Then
        CheckForWbUpdates = newestWbFullName ' Newer file exists in share folder
    End If
End Function
' Called from this workbook on close
Function newestFileInPath(path As String, FileSpec As String) As String
' Uses Microsoft FileDateTime function and uses DIR; DY 12-2-24
Dim fileName As String
Dim MostRecentFile As String
Dim MostRecentDate As Date
    fileName = Dir(path & FileSpec)
    If fileName = "" Then Exit Function ' no match

    MostRecentFile = fileName
    MostRecentDate = FileDateTime(path & fileName)
    Do While fileName <> ""
        If FileDateTime(path & fileName) > MostRecentDate Then
             MostRecentFile = fileName
             MostRecentDate = FileDateTime(path & fileName)
        End If
        fileName = Dir ' cool Dir statement loops thru file names w/ FileSpec
    Loop
    newestFileInPath = path & MostRecentFile
End Function
Private Function LastModDate(fulname As String) As Date
' https://stackoverflow.com/questions/16657170/last-modification-date-of-open-workbook
Dim fso As Object
Dim file As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.GetFile(fulname)
    If Not file Is Nothing Then
        LastModDate = file.DateLastModified
    End If
    Set fso = Nothing ' release memory
End Function
Private Sub test_SaveDateTime()
    MsgBox SaveDateTime(True)
End Sub
Private Function SaveDateTime(Optional totheSecond As Boolean = False) As String
' https://stackoverflow.com/questions/20635629/how-to-add-date-and-time-to-file-name-using-vba-in-excel
If totheSecond Then
    SaveDateTime = Format(Date, "ddmmmyyyy") & "_" & Format(Time, "h.mm.ssAM/PM") '
Else
    SaveDateTime = Format(Date, "ddmmmyyyy") & "_" & Format(Time, "hAM/PM") ' to the hour
End If
End Function

' ************************************************************
'
' THE FOLLOWING CODE:
'     AddinThisWB_ will only work on install the very first time
'                     ..NEED MORE TESTING
'
'     Delete_Addin Works but is VERY cumbersome
'
' THESE TWO Procs Are NOT very practical.
'
' ************************************************************

Private Sub test_AddinThisWB_()
Dim m As String, ans As Long
Dim thisai As AddIn
    m = Application.UserLibraryPath
    If Right(m, 1) <> "\" Then m = m & "\"
    Set thisai = AddinThisWB_(m, GetMyADDINNAME)
    
End Sub
Sub AddinThisWB()
Dim m As String, ans As Long
Dim thisai As AddIn
Dim drivelttr As String, drivetype As Long
        
        'NOT Installed in AddIns collection... Prompt user.
        m = "You just opened add-in " & ThisWorkbook.name & ". Do you want to install it as an Excel Add-In?"
        m = m & vbNewLine
        m = m & vbNewLine & "No  - Don't Install it, just Open It READ ONLY..."
        m = m & vbNewLine & "Yes - Install the add-in. "
        m = m & vbNewLine & "Cancel - Close the add-in."
        ans = MsgBox(m, vbQuestion + vbYesNoCancel, ThisWorkbook.fullName)
        Select Case ans
            Case vbYes
                drivelttr = my_Addins_Mod.A_GetDriveLetter(ThisWorkbook.fullName)
                ' Comment the next line out if Network files are OK
                'drivetype = my_Addins_Mod.GetDriveType(drivelttr)
                Select Case drivetype
                    Case -1:
                        m = "Invalid Drive Type or Drive Letter. Drive Type = " & drivetype & "  Drive = " & drivelttr & ":" & vbCrLf & vbCrLf
                        m = m & "  Try saving the file on the C: Drive and then open it."
                        m = m & vbCrLf & "Aborting..."
                        ans = MsgBox(m, vbExclamation + vbOKOnly, ThisWorkbook.fullName)
                        ThisWorkbook.Close
                        Exit Sub
                    Case 3:  ' It's a Network Drive
                        m = "This File MUST be on a local Drive.  Drive = '" & drivelttr & ":'" & vbCrLf & vbCrLf
                        m = m & "  Try saving the file on the C: Drive and then open it."
                        m = m & vbCrLf & "Aborting..."
                        ans = MsgBox(m, vbExclamation + vbOKOnly, ThisWorkbook.fullName)
                        ThisWorkbook.Close
                        Exit Sub
                End Select
                
                ' If ThisWorkbook.Name same as ADDINAME, it won't work, don't install
                If ThisWorkbook.name = GetMyADDINNAME Then
                    m = " Can't Install..." & vbCrLf
                    m = m & " Addin Name " & GetMyADDINNAME & "  matches This Workbook Name."
                    m = m & vbNewLine & " Aborting Install..."
                    ans = MsgBox(m, vbCritical, ThisWorkbook.name)
                    Exit Sub
                End If
                
                ' COPY to the AddIns folder and Install in AddIns List.
                On Error Resume Next
                ' 1-15-26 Added this saftey check for Addin Path to always have '\' at end
                m = Application.UserLibraryPath
                If Right(m, 1) <> "\" Then m = m & "\"
                Set thisai = AddinThisWB_(m, GetMyADDINNAME) ' Use Name defined in Addins Mod
                If thisai Is Nothing Then
                    m = ThisWorkbook.name & ": ThisAi is nothing. Closing Workbook.  Install Failed... "
                    ans = MsgBox(m, vbOKOnly + vbInformation)
                    ThisWorkbook.Close ' Install Failed...
                Else
                    'ans = MsgBox("Now that it's installed... Close Excel and open your workbook.", vbOKOnly + vbInformation)
                    ThisWorkbook.Close ' Success
                End If
            Case vbNo
                ThisWorkbook.IsAddin = False  ' False allows user to view the Addin Workbook
            Case vbCancel
                ThisWorkbook.Close
        End Select
End Sub

Private Function AddinThisWB_(addinPath As String, desiredTitle As String) As AddIn
    Dim m As String, ans As Long
    Dim fso As Object
    Dim thisAddInDate As Date
    Dim thisFileLen As Long
    Dim existingFile As String
    Dim existingAddinDate As Date
    Dim existingFileLen As Long
    Dim ret As Long, oldAddIn As AddIn
    Dim inList As Boolean
    Dim abortInstall As Boolean
    Dim dummyWB As Workbook
    Dim desiredFullName As String
    On Error Resume Next
     If addinPath = "" Or desiredTitle = "" Then _
        MsgBox "AddinThisWB_(): Bad Params. Aborting...": Exit Function
    
    ' Saftey Check for addinPath
    If Right(addinPath, 1) <> "\" Then
        m = "AddinThisWB_(): Error. addinPath missing '\' at end." & vbCrLf
        m = m & "addinpath = '" & addinPath & "'   desiredTitle = '" & desiredTitle & "'"
        Debug.Print
        MsgBox m & vbCrLf & vbCrLf & "Aborting...", vbCritical, "AddinThisWB_():"
        Exit Function
    End If
    desiredFullName = addinPath & desiredTitle
    
    ' See if file with desiredTitle exists in Add-In Folder
    If A_myFileExists(desiredFullName) Then
        existingFile = desiredFullName
        existingAddinDate = FileDateTime(desiredFullName)
        existingFileLen = FileLen(desiredFullName)
        
        ' Backup existing version first?
        m = "Addin file exists.  Make a back up copy the existing version first?" & vbCrLf & vbCrLf
        ans = MsgBox(m, vbYesNo, desiredTitle)
        If ans = vbYes Then
            ret = BkUp_File(oldAddIn.fullName, NETWORK_ADDIN_BACKUPS, "." & Environ("computername") & "." & Environ("username"), False)
        End If
    End If
    
    thisAddInDate = FileDateTime(ThisWorkbook.fullName)
    thisFileLen = FileLen(ThisWorkbook.fullName)
    
    ' Check if Add-in is in Add-ins list
    Set oldAddIn = Nothing
    For Each oldAddIn In Application.AddIns
        If LCase(desiredTitle) = LCase(oldAddIn.name) Then _
            Exit For
    Next
              
    If Not oldAddIn Is Nothing Then
        m = "Add-in 'Name' Exists in Add-ins List!" & vbCrLf
        m = m & " You must DELETE Existing Add-in from Add-ins folder before installing this Add-in." & vbCrLf
        m = m & " After Deleted, Open Excel addins list and check the box to remove it from the Addins." & vbCrLf
        MsgBox m, vbCritical + vbExclamation, "AddinThisWB_(): " & desiredTitle
        ' ABORTING HERE
        Exit Function
        
        ' THE FOLLOWING CODE NEEDS WORK.
        
        m = m & " Do you want to DELETE the Existing Add-in?" & vbCrLf
        m = m & vbNewLine
        m = m & vbNewLine & "  Yes - DELETE the Existing Add-in."
        m = m & vbNewLine & "  No  - Don't DELETE it... Abort."
        m = m & vbNewLine & "  Cancel - Don't DELETE... Abort."
        ans = MsgBox(m, vbQuestion + vbYesNoCancel, ThisWorkbook.name)
        If ans = vbYes Then ' Delete the Addin
            ret = Delete_Addin(oldAddIn, True)
            If ret = RETURN_SUCCESS Then
            ' DO WE WANT TO EXIT HERE REGARDLESS? ASK USER TO CHECK ADDIN LIST TO REMOVE FROM ADDINS LIST
                m = "Check the Addin's List. Close ALL Excel workbooks and Reopen to continue Addin install."
            Else
                m = "Delete existing Addin failed.  Can't ADDIN in the Workbook to Addins!. Aborting..."
            End If
            MsgBox m, vbInformation, "AddinThisWB_(): " & desiredTitle
        End If
        Exit Function ' CALLING FUNCTION MUST KNOW TO ABORT FURTHER WORK AND WAIT FOR USER
    End If
    If existingFile <> "" Then 'file already exists
        For Each oldAddIn In Application.AddIns
            If LCase(existingFile) = LCase(oldAddIn.fullName) Then
                inList = True ' Add it's in the Addins list
                Exit For
            End If
        Next
        Select Case thisAddInDate
        Case Is > existingAddinDate:
            m = "This file is NEWER than the installed Add-in... Thats good!" & vbCrLf & vbCrLf & _
                "Press Ok to update the Add-in. '" & desiredTitle & "' ?" & vbNewLine
        Case Is = existingAddinDate:
            m = "This file looks like the same as Existing Add-in..." & vbCrLf & vbCrLf & _
                "This File Date = " & vbTab & vbTab & thisAddInDate & vbTab & "File Length = " & vbTab & Format(thisFileLen, "#,##0") & vbCrLf & _
                "Existing Addin Date = " & vbTab & existingAddinDate & vbTab & "File Length = " & vbTab & Format(existingFileLen, "#,##0") & vbCrLf & _
                vbCrLf & "Aborting Install..."
                abortInstall = True
        Case Is < existingAddinDate:
            m = "This file is OLDER than Existing Add-in!..." & vbCrLf & vbCrLf & _
                "This File Date = " & vbTab & vbTab & thisAddInDate & vbTab & "File Length = " & vbTab & Format(thisFileLen, "#,##0") & vbCrLf & _
                "Existing Addin Date = " & vbTab & existingAddinDate & vbTab & "File Length = " & vbTab & Format(existingFileLen, "#,##0") & vbCrLf & _
                vbCrLf & "Aborting Install..."
                abortInstall = True
        End Select
        If abortInstall Then
            ans = MsgBox(m, vbExclamation + vbOKOnly, ThisWorkbook.name)
            Exit Function
        Else
            ans = MsgBox(m, vbOKCancel, ThisWorkbook.name)
            If ans <> vbOK Then Exit Function
        End If
    End If 'existingFile
    
    ' Check for other workbooks open HERE
    m = A_OtherWorkbooksOpen(ThisWorkbook.name)  ' Returns NULL if No other WB's open
    If A_OtherWorkbooksOpen(ThisWorkbook.name) <> "" Then
        m = m & vbCrLf & "Please Close the other workbooks first, and then try again. Must Abort."
        Debug.Print (m)
        m = m & vbCrLf & "Aborting Install..."
        ans = MsgBox(m, vbExclamation + vbOKOnly, ThisWorkbook.name)
        Exit Function
    End If

    'Open a dummy workbook to avoid an error
    'The reason is, add method is only available if a workbook is open
    If Application.ActiveWorkbook Is Nothing Then
        m = "Opening a dummy workbook to avoid an error"
        Debug.Print m
        MsgBox m
        dummyWB = Application.Workbooks.Add
        DummyWorkbookName = dummyWB.name
        MsgBox "Opened " & dummyWB.name
    End If
    
    ' Check to see if old version of Add-In is loaded in list
    If inList Then
        oldAddIn.Installed = False ' Make sure it's Uninstalled
    End If
    ' IF Addin is STILL loaded/opened, then we have to abort
    If my_Addins_Mod.IsAddinLoaded(addinPath, desiredTitle) Then '
        m = "ERROR: " & Err & " Somehow this Addin is still opened!..." & _
            " Perhaps by another application. Must Abort" & vbCrLf & vbCrLf
        Debug.Print (m)
        m = m & "Aborting Install..."
        ans = MsgBox(m, vbCritical, desiredFullName)
        Exit Function
    End If
    
    'create a file system object to copy the file
    'into addins default folder
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'copy new file to default user addin library
    m = "Copying " & ThisWorkbook.fullName & " to " & desiredFullName
    Debug.Print
    MsgBox m ' debugging
    fso.CopyFile ThisWorkbook.fullName, desiredFullName, True ' overwrite is True by default
     
    If oldAddIn Is Nothing Then
        Set AddinThisWB_ = Application.AddIns.Add(fileName:=desiredFullName)
    Else
        Set AddinThisWB_ = oldAddIn
    End If
    
    If AddinThisWB_ Is Nothing Then
        m = "AddinThisWB_(): FAILED." & vbCrLf
        m = m & "Error " & Err & ". Error installing the Add-in!" & vbCrLf
        m = m & desiredFullName
        Debug.Print (m)
        m = m & vbCrLf & vbCrLf & "Aborting Install..."
        ans = MsgBox(m, vbCritical, desiredTitle)
        Exit Function
    End If
    
    ' Set Install to true  to confirm it loaded
    AddinThisWB_.Installed = True
    
    If Not my_Addins_Mod.IsAddinLoaded(addinPath, desiredTitle) Then
        m = "Addin Install FAILED Loading!...  Loaded = " & my_Addins_Mod.IsAddinLoaded(addinPath, desiredTitle)
        Debug.Print (m)
        m = m & vbCrLf & vbCrLf & "Aborting Install..."
        ans = MsgBox(m, vbCritical, desiredTitle)
        Set AddinThisWB_ = Nothing
        Exit Function
    End If
    
    ' New Addin file size and date must match original
    Err.Clear
    thisAddInDate = FileDateTime(ThisWorkbook.fullName)
    thisFileLen = FileLen(ThisWorkbook.fullName)
    If thisAddInDate = FileDateTime(AddinThisWB_.fullName) And _
       thisFileLen = FileLen(AddinThisWB_.fullName) Then
        m = "Add-In " & AddinThisWB_.name & " is Sucessfully Installed!" & vbCrLf & vbCrLf
        Debug.Print (m)
        ans = MsgBox(m & "Now Close the Workbook...", vbInformation, ThisWorkbook.name)
    Else
        m = "Error " & Err & ". Possible error installing the Add-in. Check the Add-Ins list!" & vbCrLf & vbCrLf
        m = m & " Check if this file exists is Add-Ins Folder." & vbCrLf
        m = m & "File: " & desiredTitle & MsgFileDateFileLen(ThisWorkbook.fullName) & vbCrLf
        m = m & "Addi: " & AddinThisWB_.name & MsgFileDateFileLen(AddinThisWB_.fullName) & vbCrLf
        Debug.Print (m)
        ans = MsgBox(m, vbExclamation, ThisWorkbook.fullName)
        Set AddinThisWB_ = Nothing
    End If
End Function ' AddinThisWB_

Private Sub test_Delete_Addin() '
Dim ans As Long, ai As AddIn
Dim m As String, result As Long
Dim fullName As String
Const test1 = "C:\myBRC\Transaction Copier for QB\QB Transaction Copier.xla"
Const test2 = "C:\Users\Dwight.YACKLEY-ALIENWA\AppData\Roaming\Microsoft\AddIns\myQBAddin.xlam"
    fullName = test2
    result = RETURN_FAILED
    For Each ai In AddIns
        If LCase(ai.name) = LCase(A_FileNameOnly(fullName)) Then
             result = Delete_Addin(ai, True) ' True for prompts
             Exit For
        End If
    Next
    
    If ai Is Nothing Then
        m = "Can NOT find Addin!!" & vbCrLf & vbCrLf
        m = m & "Addin DOES NOT EXIST in Addin List!" & vbCrLf
        m = m & fullName
        Dump_AIs
        ans = MsgBox(m, vbCritical, "test_Delete_Addin()")
        Exit Sub
    End If
    If result = RETURN_SUCCESS Then
        m = "test_Delete_Addin() Reporting SUCCESS!"
        MsgBox m
    End If
    
End Sub
Private Function Delete_Addin(ai As AddIn, Optional prompt As Boolean = True) As Long
    Dim ans As Integer, ret As Long
    Dim m As String
    
    Delete_Addin = RETURN_FAILED ' default return (failed)
    ' See if Addin File exists before deleting it
    If Not A_myFileExists(ai.name) Then ' Delete failed
        m = "Addin  '" & ai.name & "'... File is already gone!" & vbCrLf & vbCrLf
        m = m & "To finish removing addin:" & vbCrLf & "1. Check the addin's box in the addins list." & vbCrLf
        m = m & "2. This will prompt Excel to report 'Cannot find add-in and ask user to 'Delete' from list?' Click Yes." & vbCrLf
        m = m & "3. Also, close and restart Excel to remove addin Name from Excel Toolbar."
        ans = MsgBox(m, vbCritical, ThisWorkbook.name)
        Delete_Addin = RETURN_SUCCESS ' return success
        Exit Function
    End If
   
    If LCase(ai.fullName) = LCase(ThisWorkbook.fullName) Then
        m = "Addin name is the same THIS Workbook. Can't Delete this Addin. Must Abort." & vbCrLf
        m = m & vbNewLine & "Aborting Delete..."
        ans = MsgBox(m, vbCritical, ai.name)
        Exit Function
    End If
    
    ' DO I NEED THIS CHECK...?? Remove later of necessary
    If LCase(ai.fullName) = LCase(ActiveWorkbook.fullName) Then
        m = "Addin name is the same ACTIVE Workbook. Can't Delete this Addin. Must Abort." & vbCrLf
        m = m & vbNewLine & "Aborting Delete..."
        ans = MsgBox(m, vbCritical, ai.name)
        Exit Function
    End If
     
    ' Unload the Addin from system
    ai.Installed = False
    
    ' Make sure it really unloaded.
    If IsAddinLoaded(ai.path, ai.name) Then
        m = "Error: " & Err & vbCrLf & "Can't Delete this Addin... File is still loaded in Workbooks()! Must Abort." & vbCrLf
        m = m & "File = " & ai.fullName & vbCrLf
        m = m & vbNewLine & "Aborting Delete..."
        ans = MsgBox(m, vbCritical, ai.name)
        Exit Function
    End If
    
    'See if we want to Backup file first before deleting
    ans = vbYes ' default
    If prompt = True Then
        m = "Backup the Addin before Deleting? '" & ai.name & "'" & vbCrLf
        m = m & vbNewLine & "    Yes - To Backup File. " ' & vbCrLf
        m = m & vbNewLine & "    No  - To Skip Backup."
        ans = MsgBox(m, vbQuestion & vbYesNo, ai.name)
    End If
    If ans = vbYes Then
         ret = BkUp_File(ai.fullName, NETWORK_ADDIN_BACKUPS, ".DELETED." & Environ("computername") & "." & Environ("username"), True)
    End If
    
    ' Confirm choice before
    ans = vbOK ' Default
    If prompt = True Then
        m = "About to UN-Install and DELETE Add-in '" & ai.name & "'" & vbCrLf
        m = m & vbNewLine & "   OK  - To UN-Install and DELETE."
        m = m & vbNewLine & "   Cancel - To Abort..."
        ans = MsgBox(m, vbOKCancel + vbInformation, ai.name)
    End If
    If ans <> vbOK Then Exit Function
    
    '
    ' DELETING the  existing addin
    '
    Debug.Print "Deleting " & ai.fullName
    Kill ai.fullName
    
    If A_myFileExists(ai.fullName) Then ' Delete failed
        Delete_Addin = -1
        m = "Delete Addin  '" & ai.name & "' FAILED!" & vbCrLf & vbCrLf & _
            " Could NOT remove file." & vbCrLf & ai.fullName
        ans = MsgBox(m, vbCritical, ThisWorkbook.name)
        Exit Function
    End If
    
    m = "Deleted Addin!" & vbCrLf & ai.fullName & vbCrLf & vbCrLf
    
    If prompt Then
        m = m & "To finish removing addin:" & vbCrLf & "1. Check the addin's box in the addins list." & vbCrLf
        m = m & "2. This will prompt Excel to report 'Cannot find add-in and ask user to 'Delete' from list?' Click Yes." & vbCrLf
        m = m & "3. Also, close and restart Excel to remove addin Name from Excel Toolbar."
        ans = MsgBox(m, vbInformation, ThisWorkbook.name)
    End If
    Delete_Addin = RETURN_SUCCESS ' success
End Function


' *****************************************************
'
'  Following are PRIVATE macros and functions copied
'  but also exist in myQBAddin as referenceable
'  functions and marcos defined in ZIP_FILE_FOLDER_EMAIL mod
'
' *****************************************************

Private Function A_myFileExists(fp As String) As Boolean ' https://www.rondebruin.nl/win/s9/win003.htm
' Late Binding... Set reference to Microsoft Scripting Runtime" because we use use Late binding here
Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("scripting.filesystemobject")
    If fso Is Nothing Then Debug.Print "A_myFileExists() FSO is NOTHING!": Exit Function
' HAVING TROUBLE HERE! Sometimes returns FALSE for existing files!
    A_myFileExists = fso.FileExists(fp)
    Set fso = Nothing
End Function

Private Sub test_A_file_ext()
    MsgBox A_File_Ext(ThisWorkbook)
End Sub
' Returns .extention of WB passed in, plus the last '.' in a string
Function A_File_Ext(wb As Workbook) As String
    A_File_Ext = LCase(Right(wb.fullName, Len(wb.fullName) - InStrRev(wb.fullName, ".") + 1))
End Function
Private Sub test_A_strip_ext()
    MsgBox A_Strip_Ext(ThisWorkbook.name)
End Sub
' Strips any and ALL ext 'x' strings at end of file name
 Function A_Strip_Ext(fname As String, Optional X As String = ".xlam") As String
    A_Strip_Ext = fname
    Do While LCase(Right(A_Strip_Ext, Len(X))) = LCase(X)
        A_Strip_Ext = LCase(Left(A_Strip_Ext, Len(A_Strip_Ext) - Len(X)))
    Loop
End Function
 Function A_FileNameOnly(pname) As String
' Returns the filename from a path/filename string
Dim temp As Variant
    If pname = "" Then Exit Function
    temp = Split(pname, Application.PathSeparator)
    A_FileNameOnly = temp(UBound(temp))
End Function
Private Sub test_A_OtherWorkbooksOpen()
    MsgBox (A_OtherWorkbooksOpen(""))
    MsgBox (A_OtherWorkbooksOpen(ThisWorkbook.name))
End Sub
Private Function A_OtherWorkbooksOpen(myWorkbookName As String) As String
Dim i As Long, wbct As Long
    wbct = 1
    For i = 1 To Workbooks.count
        If LCase(Workbooks(i).name) = LCase(myWorkbookName) Then GoTo nextOne
        If LCase(Workbooks(i).name) Like LCase("PERSONAL.x*") Then GoTo nextOne
        A_OtherWorkbooksOpen = A_OtherWorkbooksOpen & wbct & ": " & Workbooks(i).name & vbCrLf
        wbct = wbct + 1
nextOne:
    Next
    If A_OtherWorkbooksOpen <> "" Then A_OtherWorkbooksOpen = _
        "Other Workboooks Are Open:" & vbCrLf & vbCrLf & A_OtherWorkbooksOpen
End Function

Private Function A_MyDocsPath() As String
    If LCase(Environ("Username")) = "dwight" Then
        A_MyDocsPath = "\\NAS\BBMDATA\dwight\QUICKBOOKS Work area\" ' use for my work environ
    Else
        A_MyDocsPath = Environ("USERPROFILE") & "\My Documents\"
    End If
End Function
 Function A_PathNameOnly(pname) As String
' Returns the Path from a path/ string
Dim temp As Variant
Dim fname As String
Dim l As Integer
    l = Len(pname)
    temp = Split(pname, Application.PathSeparator)
     fname = temp(UBound(temp))
    l = l - Len(fname) ' leave off filename
    A_PathNameOnly = Left(pname, l)
    ' MsgBox "A_PathNameOnly(): result " & A_PathNameOnly
End Function

' Returns a one letter string with drive letter
Private Sub test_A_GetDriveLetter()
    Debug.Print ("Drive = " & A_GetDriveLetter("N:\abc")) ' N
    Debug.Print ("Drive = " & A_GetDriveLetter(A_PathNameOnly(ThisWorkbook.fullName))) ' C
    Debug.Print ("Drive = " & A_GetDriveLetter("\\NAS\BBMDATA")) ' Fails
    Debug.Print ("Drive = " & A_GetDriveLetter("\\NAS\BBMDATA\DWIGHT")) ' Fails
End Sub
' Returns a one letter string with drive letter
 Function A_GetDriveLetter(drvpath) As String
' Returns NULL if drive letter Not Exist
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/driveletter-property
Dim fs, d
    On Error Resume Next
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(fs.GetDriveName(drvpath))
    A_GetDriveLetter = d.driveLetter
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
Private Sub test_A_GetDriveType()
Dim t
    On Error Resume Next
    Debug.Print "Drive Type = " & A_GetDriveType("C")                 ' 2
    Debug.Print "Drive Type = " & A_GetDriveType("N:")                ' 3
    Debug.Print "Drive Type = " & A_GetDriveType("\\NAS\BBMDATA\")    ' 3
    Debug.Print "Drive Type = " & A_GetDriveType("\\NAS\")    ' FAILS -1
    Debug.Print "Drive Type = " & A_GetDriveType("N:\dwight") ' FAILS -1
End Sub
' Not a very useful function...
Private Function A_GetDriveType(drvpath) As Long
Dim fs, d, s, t
' Returns 0: "Unknown", 1: "Removable", 2: "Fixed", 3: "Network", 4: "CD-ROM", 5: "RAM Disk"
    On Error Resume Next
    A_GetDriveType = -1 ' default for errors set HERE!
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(drvpath)
    If d.drivetype >= 0 Then A_GetDriveType = d.drivetype
End Function
' Not a very useful function...
Private Sub test_A_ShowDriveType()
    Debug.Print A_ShowDriveType("c:\")            'Result = Drive letter C Fixed
    Debug.Print A_ShowDriveType("\\nas\\bbmdata") ' Invalid Drive
    Debug.Print A_ShowDriveType("c:\users")       ' Invalid Drive
    Debug.Print A_ShowDriveType(A_PathNameOnly(ThisWorkbook.fullName)) ' Invalid Drive
    Debug.Print A_ShowDriveType(ThisWorkbook.fullName)    ' Invalid Drive
End Sub
' Not a very useful function...
Private Function A_ShowDriveType(drvpath)
Dim fs, d, t
    On Error Resume Next
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(drvpath)
    A_ShowDriveType = "Invalid Drive letter in '" & drvpath & "'" ' Default message
    Select Case d.drivetype
        Case 0: t = "Unknown"
        Case 1: t = "Removable"
        Case 2: t = "Fixed"
        Case 3: t = "Network"
        Case 4: t = "CD-ROM"
        Case 5: t = "RAM Disk"
    End Select
    A_ShowDriveType = "Drive letter " & d.driveLetter & " " & t
End Function

' REMOVED FROM PUBLIC DY 1-11-26
Sub A_Show_Excel_Version()
    MsgBox A_Get_Excel_Version(), vbInformation
End Sub
Private Function A_Get_Excel_Version() As String
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
    A_Get_Excel_Version = "This Excel version is '" & version & "   which is... " & excel
End Function

Private Function A_BrowseToFolder(Optional initialFolder As String = "") As String
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
            A_BrowseToFolder = .SelectedItems(1)
            Next vrtSelectedItem
        End If
    End With
 'Set the object variable to Nothing.
 Set fd = Nothing
End Function

