VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Clears Fill in Active worksheet if TAB is NON-COLORED
Sub Clear_MyCells() '              Ctrl + E
   On Error Resume Next  ' Needed for if Range Is Nothing to work properly and continue thru the sub
   ' if the activesheet tab is NOT coloured, then call Clear_cells finstion with activesheet as the range
   If ActiveSheet.Tab.ColorIndex = xlColorIndexNone Then Call Clear_Cells(ActiveSheet.Cells)
End Sub

' These Macros are used by the ribbons_Ctx mod
' is myQBAddin and can are called called from the
' myQBAddin ribbon 'Run Test1', 'Run Test 2', 'Run Test 3'
Sub test1()
Dim ans As Long
    ans = MsgBox(ThisWorkbook.name & ": About to call ExportZip_My_VB()", vbYesNoCancel)
    If ans <> vbYes Then Exit Sub
    'Call MyQBAddin.ExportZip_My_VB
End Sub
Sub test2()
    MsgBox ThisWorkbook.name & ": No Macro set for Test 2().", vbInformation
End Sub
Sub Test3()
    MsgBox ThisWorkbook.name & ": No Macro set for Test 3().", vbInformation
End Sub

Private Sub Workbook_AddinInstall()
    MsgBox ThisWorkbook.name & " is now installed!", vbInformation, ThisWorkbook.name
End Sub
Private Sub Workbook_AddinUninstall()


    MsgBox ThisWorkbook.name & " is now UN-installed!", vbInformation, ThisWorkbook.name
End Sub
Private Sub Workbook_Open()
    Dim ai As AddIn, thisai As AddIn
    Dim m As String
    Dim ans As Long, newestFile As String
    Dim version As Variant, ret As Long
    Dim isInteractive As Boolean
    
    On Error Resume Next
    
    '**** FIRST! CHECK IF EXCEL IS NOT OPENED BY WINDOWS FILE EXPLORER ****
    
    isInteractive = Application.Interactive ' is Excel is running in interactive mode?
    ' If not interactive, it is highly likely a Preview Pane or background automation
    If isInteractive = False Then
        ' Close the add-in workbook without saving changes
        ThisWorkbook.Saved = True
        ThisWorkbook.Close SaveChanges:=False
    End If


    '**** CHECK version of Excel ****
    
    version = Val(Application.version)
    If version < 14 Then ' If Older than Excel 2010(version 14)
        Call my_Addins_Mod.A_Show_Excel_Version
        ThisWorkbook.Close SaveChanges:=False ' close workbook
    End If
    
    '**** CHECK version of Excel ****
    
    ' If not .xlam,  don't install , don't open
    If LCase(my_Addins_Mod.A_File_Ext(ThisWorkbook)) <> ".xlam" Then
        m = ThisWorkbook.name & " workbook is not a '.xlam' file. Aborting Open..."
        Debug.Print m
        MsgBox m, vbCritical, ThisWorkbook.name
        ThisWorkbook.Close SaveChanges:=False ' close workbook
        Exit Sub ' Done. Don't Open
    End If
    
    ' This workbook requires myQBAddin.xlam addin to be loaded first.
'This NEEDS WORK. Keeps loopin Adding In MyQbAddin
    'If Not my_Addins_Mod.Ensure_myHelperAddinIsLoaded("myQBAddin.xlam") Then
    '    M = "myQBAddin.xlam not found." & vbCrLf & _
    '            ThisWorkbook.name & " references myQBAddin.xlam." & _
    '            vbCrLf & vbCrLf & "Aborting Open..."
    '    Debug.Print M
    '    MsgBox M, vbCritical, ThisWorkbook.name
    '    ThisWorkbook.Close SaveChanges:=False ' close workbook
    '    Exit Sub ' Done. Don't Open
    'End If
    
    
    ' Is this addin in the AddIns collection?
    For Each ai In AddIns
        If LCase(ai.name) = LCase(ThisWorkbook.name) Then
            Set thisai = ai
            Exit For
        End If
    Next ai
    
    ' If thisai is nothing, User may be trying to Install the this as a addin
    If thisai Is Nothing Then
        ' Install Works only Add-in NOT in the list, NOT in the folder
        ' my_Addins_Mod.AddinThisWB STILL NEEDS WORK!
        my_Addins_Mod.AddinThisWB
    Else
    ' It's Installed.
 
        'Now check for newer File updates or in the Shared folder
        newestFile = my_Addins_Mod.CheckForWbUpdates(thisai) ' Returns newest File Name if newer
        If newestFile <> "" Then
            Debug.Print (" Newer File in Shared Folder = " & m)
            m = "FYI...        " & GetMyNETWORK_ADDIN_FOLDER & vbCrLf
            m = m & "A NEWER version of Excel AddIn " & thisai.name & " is available:" & vbCrLf
            m = m & "FILE = " & GetMyNETWORK_ADDIN_FOLDER & my_Addins_Mod.A_FileNameOnly(newestFile) & vbCrLf
            m = m & vbCrLf & "You will be asked to Install it when you Close Excel." & vbCrLf
                    
            ans = MsgBox(m, vbInformation, thisai.name)
            my_Addins_Mod.SetMyINSTALL_ADDIN_UPDATE True
        End If
    End If 'Installing ThisAi
    
    If ThisWorkbook.ReadOnly Then
    ' If launched by a background process, it's probably a File Explorer "preview" of a .xls* file
        ' Application.Visible False often indicates Excel is running in a background process.
        If Not Application.Visible Then
        ' This ADDIN could Be READ ONLY if used by File Explorer as preview of a .xls* file
            m = "WARNING!!: This Addin LAUNCHED BY A BACKGROUND PROCESS... PROBABLY FILE EXPLORER.   '" & ThisWorkbook.name & "'" & vbCrLf & vbCrLf
            m = m & ThisWorkbook.fullName & vbCrLf & vbCrLf
            m = m & "CLOSING the Addin..." & vbCrLf
            ans = MsgBox(m, vbOKOnly + vbInformation, ThisWorkbook.name)
            ' If running in background, close the workbook without saving changes
            Me.Close SaveChanges:=False
            
        Else ' it was launched by a visible application
        ' This ADDIN Should NOT Be READ ONLY
            m = "WARNING!!: This Addin OPENING AS READ ONLY!!     '" & ThisWorkbook.name & "'" & vbCrLf & vbCrLf
            m = m & ThisWorkbook.fullName & vbCrLf & vbCrLf
            m = m & "OK to Continue." & vbCrLf
            ans = MsgBox(m, vbOKOnly + vbQuestion + vbCritical, ThisWorkbook.name)
        End If
        
     End If 'read only addin
End Sub ' Workbook_Open

Private Sub Workbook_BeforeClose(Cancel As Boolean)
Dim newestFile As String, destFile As String, m As String, ans As Long
Dim prev_state As Boolean
    On Error Resume Next
    prev_state = Application.EnableEvents
    Application.EnableEvents = True



'Next line for Debugging Purposes
' my_addins_mod.setINSTALL_ADDIN_UPDATE True
    ' if INSTALL_ADDIN_UPDATE flag is set
    If my_Addins_Mod.GetMyINSTALL_ADDIN_UPDATE Then
        ' the next line should match the check on open?
        newestFile = my_Addins_Mod.CheckForWbUpdates(ThisWorkbook) ' Returns newest File Name if newer
        If newestFile = "" Then
            m = "INSTALL_ADDIN_UPDATE flag was set WITH NO Newer File in Shared Folder!"
            m = m & vbCrLf & "Skipping INSTALL!..."
            ans = MsgBox(m, vbExclamation + vbOKOnly, ThisWorkbook.name)
            GoTo EndBeforeClose
        End If
        newestFile = newestFileInPath(my_Addins_Mod.GetMyNETWORK_ADDIN_FOLDER, ThisWorkbook.name & "*.*")
        destFile = ThisWorkbook.path & "\" & my_Addins_Mod.GetMyADDINNAME
       
        m = "INSTALLING UPDATE... " & vbCrLf
        m = m & vbCrLf & newestFile & vbCrLf
        m = m & "to "
        m = m & vbCrLf & destFile & vbCrLf
        m = m & vbCrLf & "OK?"
        ans = MsgBox(m, vbQuestion + vbOK, ThisWorkbook.name)
        If ans = vbOK Then
            ' Must Use Thisworkbook.path?
            ' m = ThisWorkbook.path & "\" & "Install_Addin /P /O " & newestFile & " " & destFile
            ' 4-11-24 Install_Add.bat to use its default args for newestFile & destFile
           
           m = ThisWorkbook.path & "\" & "Install_Addin /P /O "
            ' "/S /K" Continues without waiting
            Debug.Print ThisWorkbook.name & ": Shell Command '" & m & "'"
            Call Shell("cmd.exe /S /K" & m, vbNormalFocus)
        End If
    End If
EndBeforeClose:
    Application.EnableEvents = prev_state
End Sub
Private Sub Make_ReadOnly()
Dim m As String, ans As Long
        m = "Do you want to make this file Read Only now?" & vbCrLf & vbCrLf
        ans = MsgBox(m, vbQuestion + vbYesNoCancel, ThisWorkbook.name)
        If ans = vbYes Then
        ' ReadOnly.bat is a bat file to change a file attribute to Read Only
            m = ThisWorkbook.path & "\" & "ReadOnly " & ThisWorkbook.fullName
            ' "/S /K" Continues without waiting
            Debug.Print ThisWorkbook.name & ": Shell Command '" & m & "'"
            Call Shell("cmd.exe /S /K" & m, vbNormalFocus)
        End If
End Sub
