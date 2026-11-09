Attribute VB_Name = "zEXPORT_IMPORT_VBA"
Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If


'
' Export myQBAddin  VBA
'
Sub ExportVBA_myQBAddin()
    'Activate "myQBAddin" workbook
    ThisWorkbook.Activate
    ExportVBAfromActiveWorkbook
End Sub
'
' Active Workbook VBA export to my folder
'
Private Sub ExportVBAfromActiveWorkbook()
Dim wbName As String, tmp_path As String
Dim ans As Long, m As String
    ' This works on ACTIVE Workbook!!!
    wbName = ActiveWorkbook.name
    tmp_path = "\\nas\bbmdata\Dwight\Junk\VBA\" & wbName & "_VBA\"
    m = "About to Export code from " & wbName & " to :" & vbCrLf & vbCrLf & tmp_path
    m = m & tmp_path & vbCrLf & vbCrLf
        
    ans = MsgBox(m, vbOKCancel)
    If ans <> vbOK Then Exit Sub
    
    ' Makes sure folder is there and empty
    CleanTmpFolder tmp_path
    
    'Optional directory = "", Optional prompt = True, Optional modsOnly = "*.*")
    ExportVisualBasicCode tmp_path, False, "*"
    
    
End Sub

Private Sub ExportZippedVBAfromActiveWorkbook()
Dim wbName As String, tmp_path As String
Dim ans As Long, m As String
Dim Zipfilename As String, zip_path As String

    ' This works on ACTIVE Workbook!!!
    wbName = ActiveWorkbook.name
    Zipfilename = wbName
    tmp_path = CreateUserTempFolder(wbName) & "\"
    zip_path = "\\nas\bbmdata\Dwight\Junk\VBA_ZIPPED\"
    Zipfilename = Zipfilename '
    
    m = "About to Export code from " & wbName & " to :" & vbCrLf & vbCrLf
    m = m & tmp_path & vbCrLf & vbCrLf
    m = m & "Components will be zipped into a file called: " & Zipfilename & vbCrLf & vbCrLf
    m = m & "The zipped file is located in:" & vbCrLf
    m = m & zip_path
     
       
    ans = MsgBox(m, vbOKCancel)
    If ans <> vbOK Then Exit Sub
    
    ' Makes sure folder is there and empty
    CleanTmpFolder tmp_path
    
    
    'Optionals: directory = "", Optional prompt = True, Optional modsOnly = "*.bas")
    tmp_path = ExportVisualBasicCode(tmp_path, False, "*.bas")
    
    'variables: folderName As Variant, optionals:  defPath "", ZipName  "", prompt = True)
    m = Zip_All_Files_in_Folder(tmp_path, zip_path, Zipfilename, False)
    
    If m = "" Then
        MsgBox "ExportZippedVBAfromActiveWorkbook failed. See Immediate Window for more details."
        Exit Sub
    End If
    MsgBox "Created " & m
    
    ' Open widows explorer with folder for User
    OpenSpecificFolder zip_path

End Sub
Sub OpenSpecificFolder(folderPath As String)
    
    ' Launch Windows File Explorer to the target folder
    Shell "explorer.exe """ & folderPath & """", vbNormalFocus
End Sub
Function CreateUserTempFolder(FolderName As String) As String
    Dim fso As Object
    Dim tempPath As String
    Dim newPath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Get user temp folder path
    tempPath = fso.GetSpecialFolder(2)
    
    ' Build new folder path
    newPath = tempPath & "\" & FolderName
    
    ' Create folder if it does not exist
    If Not fso.FolderExists(newPath) Then
        fso.CreateFolder newPath
    End If
    
    ' Return the full path
    CreateUserTempFolder = newPath
End Function


' Export IIF CREATOR Workbook VBA
' MUST SAVE IIF CREATOR.XLA AS .XLSM FIRST! because ExportVisualBasicCode ONLY works with ACTIVE workbook
Private Sub test_ExportIIF_CREATORCode()
Dim wbName As String, path As String
Dim ans As Long
    ' This works on ACTIVE Workbook!!!
    wbName = ActiveWorkbook.name
    path = "\\NAS\bbmdata\Dwight\VBA Stuff\PROTECTION CRACKING\" & wbName & "_VBA\"
    ans = MsgBox("About to Export code from " & wbName & " to " & _
        vbCrLf & vbCrLf & path, vbOKCancel)
    If ans <> vbOK Then Exit Sub
    CleanTmpFolder path
    ExportVisualBasicCode path, False, "*"
End Sub
' Running this I found a LOT of Junk in my AI list
Private Sub Dump_AIs()
Dim ai As AddIn
    For Each ai In AddIns
        Debug.Print ai.name
    Next ai
End Sub
Private Sub Dump_Workbooks()
Dim wb As Workbook
    For Each wb In Workbooks
        Debug.Print wb.name
    Next wb
End Sub


' Active Workbook VBA Export
' https://gist.github.com/steve-jansen/7589478  But originnaly was from RdB
' Excel macro to export all VBA source code in this project to text files for proper source control versioning
' Requires enabling the Excel setting in Options/Trust Center/Trust Center Settings/Macro Settings/Trust access to the VBA project object model
Function ExportVisualBasicCode(Optional directory As String = "", Optional prompt As Boolean = True, Optional modsOnly As String = "*.*")
    
    Const Padding = 24
    
    Dim ans As Long
    Dim VBComponent As Object
    Dim count As Integer
    Dim path As String
    Dim fso As New FileSystemObject
    Dim compName As String
    
    If prompt Then directory = BrowseToFolder(directory)
    If directory = "" Then Exit Function
    If Right(directory, 1) <> "\" Then directory = directory & "\"
  
    If Not fso.FolderExists(directory) Then
        Call fso.CreateFolder(directory)
    End If
    Set fso = Nothing
    
    For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
        compName = ComponentFileName(VBComponent)
        If compName Like modsOnly Then
            On Error Resume Next
            Err.Clear
            path = directory & compName
            
            Call VBComponent.Export(path)
            If Err.Number <> 0 Then
                Call MsgBox("Failed to export " & compName & " to " & path, vbCritical)
            Else
                count = count + 1
                'Debug.Print "Exported " & Left$(compName & ":" & Space(Padding), Padding) & path
            End If
            On Error GoTo 0
        End If
    Next
    If prompt Then _
    ans = MsgBox("Successfully exported " & CStr(count) & " VBA files to " & directory, vbInformation, ActiveWorkbook.name)
    ExportVisualBasicCode = directory
End Function
Private Function ComponentFileName(VBComponent As Object) As String
Dim extension As String
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
        Select Case VBComponent.Type
            Case ClassModule, Document
                extension = ".cls"
            Case Form
                extension = ".frm"
            Case Module
                extension = ".bas"
            Case Else
                extension = ".UNKNOWN " & VBComponent.Type & " .txt"
        End Select
        ComponentFileName = VBComponent.name & extension
End Function
Private Sub test_ApplicationStatusBar()
    Application.StatusBar = "Successfully exported " & 5 & " VBA files to " & "C:\"
    'Application.displayStatusBar = True
    Application.OnTime Now + TimeSerial(0, 0, 10), "ClearStatusBar"
End Sub
Private Sub ClearStatusBar()
    Application.StatusBar = False
End Sub
' https://docs.microsoft.com/en-us/office/vba/api/excel.application.statusbar
Private Sub testApplicationStatusBar2()
Dim oldstatusbar
    oldstatusbar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
    Application.StatusBar = "Please be patient..."
   ' Workbooks.Open FileName:="LARGE.XLS"
    Application.OnTime Now + TimeSerial(0, 0, 10), "ClearStatusBar"
    Application.DisplayStatusBar = False
    Application.DisplayStatusBar = oldstatusbar
End Sub


Private Sub test_CheckForComponents()
    MsgBox "CheckForComponents = " & CheckForComponents("\\NAS\bbmdata\Dwight\JUNK\VBA\O_VBA")
End Sub
Function CheckForComponents(directory As String) As Boolean

Dim VBComponent As Object
Dim objFSO As Scripting.FileSystemObject
Dim objFile As Scripting.file, extension As String
Dim compName As String, ans As Long
    
    If directory = "" Then Exit Function
    If Right(directory, 1) <> "\" Then directory = directory & "\"
    
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(directory).Files.count = 0 Then
       MsgBox "There are no files to import."
       CheckForComponents = True
       Exit Function
    End If
    For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
        compName = ComponentFileName(VBComponent)
        If FileOrFolderExists(directory & compName) Then
            ans = MsgBox("Warning! Modules already exist!...  '" & compName & "'", vbCritical, "WARNING!")
            CheckForComponents = True
            Exit Function
        End If
    Next
    CheckForComponents = False
End Function
Sub Delete_VBAModules(onlyMatching As String)
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim compName As String
    
    Set vbProj = ActiveWorkbook.VBProject
        For Each vbComp In vbProj.VBComponents
            If vbComp.Type = vbext_ct_Document Then
                'Thisworkbook or worksheet module
                'We do nothing
            Else
                compName = ComponentFileName(vbComp)
                If compName Like onlyMatching Then
                    If MsgBox("Remove Component " & vbComp.name & "?", vbYesNo) = vbYes Then
                        vbProj.VBComponents.Remove vbComp
                        Debug.Print "Removed Component " & vbComp.name
                    End If
                End If
            End If
        Next vbComp
End Sub
Private Sub Delete_VBAModulesDirectory(directory As String)
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim compName As String
    
    If directory = "" Then Exit Sub
    If Right(directory, 1) <> "\" Then directory = directory & "\"
    
    Set vbProj = ActiveWorkbook.VBProject
        For Each vbComp In vbProj.VBComponents
            If vbComp.Type = vbext_ct_Document Then
                'Thisworkbook or worksheet module
                'We do nothing
            Else
                compName = ComponentFileName(vbComp)
                If FileOrFolderExists(directory & compName) Then
                    Debug.Print "Removing Component " & vbComp.name
                    vbProj.VBComponents.Remove vbComp
                End If
            End If
        Next vbComp
End Sub
Private Sub DeleteALL_VBAModulesAndUserForms()
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
        Set vbProj = ActiveWorkbook.VBProject
        For Each vbComp In vbProj.VBComponents
            If vbComp.Type = vbext_ct_Document Then
                'Thisworkbook or worksheet module
                'We do nothing
            Else
                vbProj.VBComponents.Remove vbComp
            End If
        Next vbComp
End Sub

