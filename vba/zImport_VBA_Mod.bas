Attribute VB_Name = "zImport_VBA_Mod"
'Attribute VB_Name = "zImport_VBA_Mod"
Option Explicit

'===================================================================================
' MODULE: VBAImportUtils
' Purpose: Imports exported VBA component files (.bas / .cls / .frm+.frx) from a
'          folder into the ActiveWorkbook's VBA project.
'
'          SMART HANDLING OF .cls FILES:
'          - If a .cls file's name matches an EXISTING Document-type component
'            (ThisWorkbook, Sheet1, Sheet2, etc. - Type = vbext_ct_Document),
'            its code is injected into that component's CodeModule (safe -
'            preserves the component's binding to the actual sheet/workbook).
'          - Otherwise, it's treated as an ordinary class module: any existing
'            component with the same name is removed, then the file is
'            imported normally.
'
' Requirements:
'   - "Trust access to the VBA project object model" must be enabled:
'     File > Options > Trust Center > Trust Center Settings > Macro Settings.
'
' Error handling:
'   - Every failure point in this module logs details to the Immediate window
'     via Debug.Print AND shows a MsgBox describing exactly what failed, which
'     file/component was involved, and the Err.Number/Err.Description.
'===================================================================================

' vbext_ComponentType constants (used via late binding, so declared manually)
Private Const CT_DOCUMENT As Long = 100    ' vbext_ct_Document (ThisWorkbook, Sheets)
Private Const CT_MSFORM As Long = 3        ' vbext_ct_MSForm   (.frm)
Private Const CT_STDMODULE As Long = 1     ' vbext_ct_StdModule (.bas)
Private Const CT_CLASSMODULE As Long = 2   ' vbext_ct_ClassModule (.cls)
' const CT_DOCUMENT is used by Procedure : ImportVBAComponentsFromFolder
'                        Procedure Purpose : Imports all .bas, .cls, and .frm (with companion .frx) files
'                                           Routes  .cls files intelligently: document-module name collisions go
'                                            through CodeModule injection;
'                                           Everything else uses standard Remove + Import.

'
' Active Workbook VBA Import from a .Zip File
'
Private Sub ImportZippedVBAtoActiveWorkbook()
Dim wbName As String, zipFilePath As String, tmpFolder As String
Dim ans As Long, m As String, result As Boolean
    ' This uses ACTIVE Workbook!!!
    wbName = ActiveWorkbook.name
   
    'zipFilePath = SelectFilePath("C:\Users\Dwight.YACKLEY-ALIENWA\Downloads\*.zip")
    zipFilePath = SelectFilePath(ActiveWorkbook.path)

    tmpFolder = "\\nas\bbmdata\Dwight\Junk\VBA\" & "TEMP" & "_VBA\"
'    M = "About to Import code to Active Workbook" & wbName & vbCrLf & vbCrLf

    result = ImportModulesFromZip(zipFilePath, tmpFolder)
    If result Then
        m = "Successfully imported all zipped files."
        MsgBox m
        Exit Sub
    Else
        m = "Failed!  Check Debug.print messages in Immediate Window for more details."
        MsgBox m, vbCritical + vbInformation, "ImportZippedVBAtoActiveWorkbook()..."
    End If
End Sub

'==============================================================================
' Added:    ImportVBAFiletoActiveWorkbook / ImportSingleBasModule / ImportSingleVBAComponent
' Purpose:  Same workflow as ImportZippedVBAtoActiveWorkbook (prompt for a file,
'           warn the user, import into the ActiveWorkbook's VBA project) but for
'           a single, already-unzipped .bas file - no unzip/temp-folder step.
'==============================================================================
'
' Active Workbook VBA Import from a single .bas File
'
Sub ImportVBAFiletoActiveWorkbook()
Dim wbName As String, basFilePath As String, result As Boolean
Dim m As String
    ' This uses ACTIVE Workbook!!!
    wbName = ActiveWorkbook.name

    basFilePath = SelectFilePath("C:\Users\Dwight.YACKLEY-ALIENWA\Downloads\*.bas")

    result = ImportSingleBasModule(basFilePath)
    
    If result Then
        m = "Successfully imported file."
        MsgBox m
        Exit Sub
    Else
        m = "Failed!  Check Debug.print messages in Immediate Window for more details."
        MsgBox m, vbCritical + vbInformation, "ImportVBAFiletoActiveWorkbook()..."
    End If
End Sub

'-------------------------------------------------------------------------------
' Function  : ImportSingleBasModule
' Purpose   : Validates basFilePath, shows the same WARNING/confirmation prompt
'             used by ImportModulesFromZip, then imports the single .bas file
'             into the ActiveWorkbook's VBA project.
' Inputs:   basFilePath - String - full path to a single .bas file
' Returns:  Boolean - True if the file was imported successfully
'-------------------------------------------------------------------------------
Public Function ImportSingleBasModule(ByVal basFilePath As String) As Boolean
Dim m As String, ans As Long
Dim fso As Object

    ImportSingleBasModule = False

    If Len(basFilePath) = 0 Then
        Exit Function   ' user canceled file selection
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(basFilePath) Then
        m = "Not a valid file. " & vbCrLf & vbCrLf
        m = m & basFilePath
        MsgBox m, vbCritical, " ImportSingleBasModule():"
        Exit Function
    End If

    If LCase$(fso.GetExtensionName(basFilePath)) <> "bas" Then
        m = "Not a .bas file. " & vbCrLf & vbCrLf
        m = m & basFilePath
        MsgBox m, vbCritical, " ImportSingleBasModule():"
        Exit Function
    End If

    m = "WARNING: You are about to Import a VBA component into the workbook named " & ActiveWorkbook.name & vbCrLf & vbCrLf
    If ActiveWorkbook.name = ThisWorkbook.name Then
        m = m & "WARNING: ACTIVE Workbook is the same as THIS Workbook. " & ThisWorkbook.name & vbCrLf & vbCrLf
    End If
    m = m & "The following vba component will be added or replaced: " & vbCrLf & vbCrLf
    m = m & vbTab & fso.GetFileName(basFilePath) & vbCrLf
    m = m & vbCrLf & vbCrLf & "Are you sure you want to continue?"
    ans = MsgBox(m, vbYesNo + vbQuestion + vbCritical, "WARNING...")
    If ans <> vbYes Then Exit Function

    ' Imports the single .bas Standard Module component
    If Not ImportSingleVBAComponent(basFilePath) Then
        m = "Import VBA Component from .bas file failed. " & vbCrLf & vbCrLf
        m = m & "File '" & basFilePath & "'"
        MsgBox m, vbCritical, " ImportSingleBasModule():"
        Exit Function
    End If

    ImportSingleBasModule = True

End Function

'-------------------------------------------------------------------------------
' Function  : ImportSingleVBAComponent
' Purpose   : Imports one .bas file into the ActiveWorkbook's VBA project,
'             mirroring the "bas" Case branch of ImportVBAComponentsFromFolder
'             (Remove any existing same-named component, Import, then rename
'             if the file's internal Attribute VB_Name is stale).
' Parameters:
'   filePath - (String) Full path to the .bas file to import.
' Returns   : (Boolean) True if imported (and renamed, if needed) successfully.
'-------------------------------------------------------------------------------
Function ImportSingleVBAComponent(ByVal filePath As String) As Boolean

    Dim vbProj As Object
    Dim fso As Object
    Dim importedComp As Object
    Dim componentName As String
    Dim m As String

    ImportSingleVBAComponent = False
    On Error GoTo ErrHandler

    Set fso = CreateObject("Scripting.FileSystemObject")
    componentName = fso.GetBaseName(filePath)

    ' Confirm access to the VBE / active VBProject before proceeding.
    ' Most common failure point: "Trust access to the VBA project object model"
    ' disabled in Trust Center -> Err 70 / "Programmatic access is not trusted".
    On Error Resume Next
    Err.Clear
    Set vbProj = Application.VBE.ActiveVBProject
    If Err.Number <> 0 Or vbProj Is Nothing Then
        m = "Could not access Application.VBE.ActiveVBProject." & vbCrLf & vbCrLf
        m = m & "Err: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf
        m = m & "This usually means 'Trust access to the VBA project object model' " & _
                "is disabled (File > Options > Trust Center > Trust Center Settings > " & _
                "Macro Settings)."
        Debug.Print "ImportSingleVBAComponent(): " & m
        MsgBox m, vbCritical, "ImportSingleVBAComponent():"
        On Error GoTo ErrHandler
        Exit Function
    End If
    On Error GoTo ErrHandler

    ' Standard modules: safe to Remove + Import
    RemoveComponentIfExists vbProj, componentName, includeDocuments:=False

    On Error Resume Next
    Err.Clear
    Set importedComp = vbProj.VBComponents.Import(filePath)
    If Err.Number <> 0 Then
        m = "Import failed for '" & fso.GetFileName(filePath) & "' (bas)." & vbCrLf & vbCrLf
        m = m & "Path: " & filePath & vbCrLf
        m = m & "Err: " & Err.Number & " - " & Err.Description
        Debug.Print "ImportSingleVBAComponent(): FAILURE importing '" & fso.GetFileName(filePath) & "' - " & m
        MsgBox m, vbCritical, "ImportSingleVBAComponent():"
        Err.Clear
        On Error GoTo ErrHandler
        Exit Function
    End If
    On Error GoTo ErrHandler

    Debug.Print "ImportSingleVBAComponent(): SUCCESS imported '" & fso.GetFileName(filePath) & "' as component '" & importedComp.name & "'"

    ' Import() names the component from the file's internal "Attribute VB_Name"
    ' value, not the file name - if that's stale, rename it to match the file.
    If Not RenameComponentIfNeeded(importedComp, componentName) Then Exit Function

    ImportSingleVBAComponent = True
    Exit Function

ErrHandler:
    m = "Unhandled error in ImportSingleVBAComponent." & vbCrLf & vbCrLf
    m = m & "File: " & filePath & vbCrLf
    m = m & "Err: " & Err.Number & " - " & Err.Description
    Debug.Print "ImportSingleVBAComponent(): " & m
    MsgBox m, vbCritical, "ImportSingleVBAComponent():"
    ImportSingleVBAComponent = False
End Function


Private Sub test_ImportModulesFromZip()
    MsgBox (ImportModulesFromZip("C:\Users\Dwight.YACKLEY-ALIENWA\Downloads\files(2).zip", "\\NAS\bbmdata\Dwight\JUNK\VBA\NS_RPT"))
End Sub
Public Function ImportModulesFromZip(ByVal Zipfile As String, ByVal tmpFolderPath) As Boolean
Dim m As String, ans As Long
    ImportModulesFromZip = False ' default reply
    
    ' Ensure the path ends with a backslash
    If Right(tmpFolderPath, 1) <> "\" Then
        tmpFolderPath = tmpFolderPath & "\"
    End If
    
    ' Unzip File to tmp folder
    If Not ExtractZipToFolder(Zipfile, tmpFolderPath) Then
        m = "Not Valid ZIP file or User canceled Unzip. " & vbCrLf & vbCrLf
        m = m & Zipfile
        MsgBox m, vbCritical, " ImportModulesFromZip():"
        Exit Function
    End If
    
    m = "WARNING: You are about to Import VBA components into the workbook named " & ActiveWorkbook.name & vbCrLf & vbCrLf
    If ActiveWorkbook.name = ThisWorkbook.name Then
        m = m & "WARNING: ACTIVE Workbook is the same as THIS Workbook. " & ThisWorkbook.name & vbCrLf & vbCrLf
    End If
    m = m & "The following vba components will be added or replaced: " & vbCrLf & vbCrLf
    m = m & ListFileNamesInFolder(tmpFolderPath)
    m = m & vbCrLf & vbCrLf & "Are you sure you want to continue?"
    ans = MsgBox(m, vbYesNo + vbQuestion + vbCritical, "WARNING...")
    If ans <> vbYes Then Exit Function
    
    ' Imports All VBA components... Modules, Forms  and .cls "Document" files intelligently
    If Not ImportVBAComponentsFromFolder(tmpFolderPath) Then
        m = "Import VBA Components from Zip file failed. " & vbCrLf & vbCrLf
        m = m & "Zip file '" & Zipfile & "'" & vbCrLf & vbCrLf
        m = m & "Temp folder '" & tmpFolderPath & "'"
        MsgBox m, vbCritical, " ImportVBAComponentsFromFolder():"
        ImportModulesFromZip = False
        Exit Function
    End If
    ImportModulesFromZip = True ' Success
End Function

'==============================================================================
' Function: ListFileNamesInFolder
' Purpose:  Returns a string containing the names of all files in the
'           specified folder. Each file name is preceded by a tab character
'           and followed by a vbCrLf.
' Inputs:   folderPath - String - full path to the target folder
' Returns:  String - tab-prefixed, line-separated list of file names
'==============================================================================
Function ListFileNamesInFolder(ByVal folderPath As String) As String
    Dim result As String
    Dim fileName As String
    
    ' Ensure the path ends with a backslash
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If
    
    fileName = Dir(folderPath)
    
    Do While fileName <> ""
        result = result & vbTab & fileName & vbCrLf
        fileName = Dir()
    Loop
    
    ListFileNamesInFolder = result
End Function

Private Sub test_ImportVBAComponentsFromFolder()
    MsgBox (ImportVBAComponentsFromFolder(""))
End Sub

'-------------------------------------------------------------------------------
' Procedure : ImportVBAComponentsFromFolder
' Purpose   : Imports all .bas, .cls, and .frm (with companion .frx) files
'             from folderPath into the ActiveWorkbook's VBA project. Routes
'             .cls files intelligently: document-module name collisions go
'             through CodeModule injection; everything else uses standard
'             Remove + Import.
'
' Parameters:
'   folderPath - (String) Full path to the folder containing exported VBA
'                component files (.bas / .cls / .frm / .frx).
'
' Returns   : (Boolean) True if all discovered components were processed
'             successfully; False if any error occurred.
'-------------------------------------------------------------------------------
Function ImportVBAComponentsFromFolder(ByVal folderPath As String) As Boolean

    Dim fso As Object
    Dim vbProj As Object
    Dim fileItem As Object
    Dim folderItem As Object
    Dim importedComp As Object
    Dim componentName As String
    Dim ext As String
    Dim success As Boolean
    Dim m As String
    Dim isDoc As Boolean

    ImportVBAComponentsFromFolder = False
    success = True

    On Error GoTo ErrHandler

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Len(folderPath) = 0 Then
        m = "ImportVBAComponentsFromFolder was called with an empty folderPath."
        Debug.Print "ImportVBAComponentsFromFolder(): " & m
        MsgBox m, vbCritical, "ImportVBAComponentsFromFolder():"
        Exit Function
    End If

    If Not fso.FolderExists(folderPath) Then
        m = "Folder does not exist: " & folderPath
        Debug.Print "ImportVBAComponentsFromFolder(): " & m
        MsgBox m, vbCritical, "ImportVBAComponentsFromFolder():"
        Exit Function
    End If

    If Right$(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    ' Confirm access to the VBE / active VBProject before proceeding.
    ' Most common failure point: "Trust access to the VBA project object model"
    ' disabled in Trust Center -> Err 70 / "Programmatic access is not trusted".
    On Error Resume Next
    Err.Clear
    Set vbProj = Application.VBE.ActiveVBProject
    If Err.Number <> 0 Or vbProj Is Nothing Then
        m = "Could not access Application.VBE.ActiveVBProject." & vbCrLf & vbCrLf
        m = m & "Err: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf
        m = m & "This usually means 'Trust access to the VBA project object model' " & _
                "is disabled (File > Options > Trust Center > Trust Center Settings > " & _
                "Macro Settings)."
        Debug.Print "ImportVBAComponentsFromFolder(): " & m
        MsgBox m, vbCritical, "ImportVBAComponentsFromFolder():"
        On Error GoTo ErrHandler
        Exit Function
    End If
    On Error GoTo ErrHandler

    Set folderItem = fso.GetFolder(folderPath)

    For Each fileItem In folderItem.Files

        ext = LCase$(fso.GetExtensionName(fileItem.name))
        componentName = fso.GetBaseName(fileItem.name)

        Select Case ext
        ' I added the next Case to get NS_Trade_Historical imports to work from Claude
            Case "frx"
                Debug.Print "ImportVBAComponentsFromFolder(): skipping '" & fileItem.name & "' (.frm is auto-handled via companion .frm)"
            
            Case "bas", "frm"
                ' Standard modules and forms: safe to Remove + Import
                RemoveComponentIfExists vbProj, componentName, includeDocuments:=False

                On Error Resume Next
                Err.Clear
                Set importedComp = vbProj.VBComponents.Import(fileItem.path)
                If Err.Number <> 0 Then
                    m = "Import failed for '" & fileItem.name & "' (" & ext & ")." & vbCrLf & vbCrLf
                    m = m & "Path: " & fileItem.path & vbCrLf
                    m = m & "Err: " & Err.Number & " - " & Err.Description
                    Debug.Print "ImportVBAComponentsFromFolder(): FAILURE importing '" & fileItem.name & "' - " & m
                    MsgBox m, vbCritical, "ImportVBAComponentsFromFolder():"
                    Err.Clear
                    success = False
                    On Error GoTo ErrHandler
                Else
                    On Error GoTo ErrHandler
                    Debug.Print "ImportVBAComponentsFromFolder(): SUCCESS imported '" & fileItem.name & "' as component '" & importedComp.name & "'"
                    ' Import() names the component from the file's internal "Attribute VB_Name"
                    ' value, not the file name - if that's stale, rename it to match the file.
                    If Not RenameComponentIfNeeded(importedComp, componentName) Then success = False
                End If

            Case "cls"
                ' Smart routing: document module vs. ordinary class module
                isDoc = ComponentExistsAsDocument(vbProj, componentName)

                If isDoc Then
                    If Not InjectCodeModuleFromFile(vbProj, componentName, fileItem.path) Then
                        m = "InjectCodeModuleFromFile failed for '" & componentName & "' (" & fileItem.name & "). " & _
                            "See the message above (or Immediate window) for the specific cause."
                        Debug.Print "ImportVBAComponentsFromFolder(): FAILURE - " & m
                        MsgBox m, vbCritical, "ImportVBAComponentsFromFolder():"
                        success = False
                    Else
                        Debug.Print "ImportVBAComponentsFromFolder(): SUCCESS injected code into document component '" & componentName & "' from '" & fileItem.name & "'"
                    End If
                Else
                    RemoveComponentIfExists vbProj, componentName, includeDocuments:=False

                    On Error Resume Next
                    Err.Clear
                    Set importedComp = vbProj.VBComponents.Import(fileItem.path)
                    If Err.Number <> 0 Then
                        m = "Import failed for '" & fileItem.name & "' (cls)." & vbCrLf & vbCrLf
                        m = m & "Path: " & fileItem.path & vbCrLf
                        m = m & "Err: " & Err.Number & " - " & Err.Description
                        Debug.Print "ImportVBAComponentsFromFolder(): FAILURE importing '" & fileItem.name & "' - " & m
                        MsgBox m, vbCritical, "ImportVBAComponentsFromFolder():"
                        Err.Clear
                        success = False
                        On Error GoTo ErrHandler
                    Else
                        On Error GoTo ErrHandler
                        Debug.Print "ImportVBAComponentsFromFolder(): SUCCESS imported '" & fileItem.name & "' as component '" & importedComp.name & "'"
                        If Not RenameComponentIfNeeded(importedComp, componentName) Then success = False
                    End If
                End If

            Case "frx"
                Debug.Print "ImportVBAComponentsFromFolder(): skipping '" & fileItem.name & "' (.frx is auto-handled via companion .frm)"

            Case Else
                Debug.Print "ImportVBAComponentsFromFolder(): skipping '" & fileItem.name & "' (extension '" & ext & "' not handled)"

        End Select

    Next fileItem

    ImportVBAComponentsFromFolder = success
    If success Then
        ImportVBAComponentsFromFolder = True
        Debug.Print "ImportVBAComponentsFromFolder(): completed successfully for folder " & folderPath
    Else
        ImportVBAComponentsFromFolder = False
        Debug.Print "ImportVBAComponentsFromFolder(): completed WITH FAILURES for folder " & folderPath & " - see messages above."
    End If
    Exit Function

ErrHandler:
    m = "Unhandled error in ImportVBAComponentsFromFolder." & vbCrLf & vbCrLf
    m = m & "Folder: " & folderPath & vbCrLf
    If Not fileItem Is Nothing Then
        m = m & "Current file: " & fileItem.name & vbCrLf
    End If
    m = m & "Err: " & Err.Number & " - " & Err.Description
    Debug.Print "ImportVBAComponentsFromFolder(): " & m
    MsgBox m, vbCritical, "ImportVBAComponentsFromFolder():"
    ImportVBAComponentsFromFolder = False
End Function


'-------------------------------------------------------------------------------
' Procedure : ComponentExistsAsDocument
' Purpose   : Checks whether a component with the given name currently exists
'             in vbProj AND is a Document-type component (ThisWorkbook, or a
'             Sheet/Chart module).
'
' Parameters:
'   vbProj        - (Object) Target VBProject.
'   componentName - (String) Component name to check.
'
' Returns   : (Boolean) True if a same-named Document-type component exists.
'-------------------------------------------------------------------------------
Private Function ComponentExistsAsDocument(ByVal vbProj As Object, ByVal componentName As String) As Boolean

    Dim comp As Object
    Dim m As String

    ComponentExistsAsDocument = False
    On Error GoTo ErrHandler

    On Error Resume Next
    Err.Clear
    Set comp = vbProj.VBComponents(componentName)
    On Error GoTo ErrHandler

    If Not comp Is Nothing Then
        ComponentExistsAsDocument = (comp.Type = CT_DOCUMENT)
    End If

    Exit Function

ErrHandler:
    m = "ComponentExistsAsDocument failed for component '" & componentName & "'." & vbCrLf & vbCrLf
    m = m & "Err: " & Err.Number & " - " & Err.Description
    Debug.Print "ComponentExistsAsDocument(): " & m
    MsgBox m, vbCritical, "ComponentExistsAsDocument():"
    ComponentExistsAsDocument = False
End Function


'-------------------------------------------------------------------------------
' Function  : RenameComponentIfNeeded
' Purpose   : Renames a just-imported VBComponent to match the desired name
'             (the file's base name), if it doesn't already match.
'
'             Import() names the new component from the "Attribute VB_Name"
'             line stored INSIDE the file, not from the file name on disk.
'             If that attribute is stale (e.g. the file "MyReport.bas"
'             internally still says Attribute VB_Name = "Module1", left
'             over from when it was first created), the imported component
'             ends up named "Module1" unless it's explicitly renamed here.
'
' Parameters:
'   comp        - (Object) The just-imported VBComponent.
'   desiredName - (String) The name it should have (the file's base name).
'
' Returns   : (Boolean) True if the component already had the desired name,
'             or the rename succeeded. False if the rename failed (e.g. a
'             component with that name already exists, or the name is
'             invalid).
'-------------------------------------------------------------------------------
Private Function RenameComponentIfNeeded(ByVal comp As Object, ByVal desiredName As String) As Boolean

    Dim oldName As String
    Dim m As String

    RenameComponentIfNeeded = True

    If comp Is Nothing Then Exit Function
    If comp.name = desiredName Then Exit Function

    oldName = comp.name

    On Error Resume Next
    Err.Clear
    comp.name = desiredName
    If Err.Number <> 0 Then
        m = "Rename failed: '" & oldName & "' -> '" & desiredName & "'." & vbCrLf & vbCrLf
        m = m & "Err: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf
        m = m & "(Likely cause: a component named '" & desiredName & "' already exists in the " & _
                "project, or '" & desiredName & "' is not a valid VBA component name.)"
        Debug.Print "RenameComponentIfNeeded(): FAILURE " & m
        MsgBox m, vbCritical, "RenameComponentIfNeeded():"
        Err.Clear
        RenameComponentIfNeeded = False
    Else
        Debug.Print "RenameComponentIfNeeded(): SUCCESS renamed '" & oldName & "' -> '" & desiredName & "'"
    End If
    On Error GoTo 0

End Function


'-------------------------------------------------------------------------------
' Procedure : RemoveComponentIfExists
' Purpose   : Removes a VBA component (module/class/form) from vbProj if a
'             component with the specified name currently exists. Document-
'             type components are never removed (they can't be, and doing so
'             would raise an error) - guarded by includeDocuments.
'
' Parameters:
'   vbProj           - (Object) Target VBProject.
'   componentName    - (String) Name of the component to remove, if found.
'   includeDocuments - (Boolean) Reserved safety flag; Document-type
'                      components are always skipped regardless of this
'                      value, since VBIDE does not allow their removal.
'-------------------------------------------------------------------------------
Private Sub RemoveComponentIfExists(ByVal vbProj As Object, ByVal componentName As String, ByVal includeDocuments As Boolean)

    Dim comp As Object
    Dim m As String

    On Error GoTo ErrHandler

    On Error Resume Next
    Err.Clear
    Set comp = vbProj.VBComponents(componentName)
    On Error GoTo ErrHandler

    If Not comp Is Nothing Then
        If comp.Type <> CT_DOCUMENT Then
            On Error Resume Next
            Err.Clear
            vbProj.VBComponents.Remove comp
            If Err.Number <> 0 Then
                m = "Remove failed for component '" & componentName & "'." & vbCrLf & vbCrLf
                m = m & "Err: " & Err.Number & " - " & Err.Description
                Debug.Print "RemoveComponentIfExists(): " & m
                MsgBox m, vbCritical, "RemoveComponentIfExists():"
                Err.Clear
            End If
            On Error GoTo ErrHandler
        Else
            Debug.Print "RemoveComponentIfExists(): skipping removal of '" & componentName & "' - it is a Document-type component."
        End If
    End If

    Exit Sub

ErrHandler:
    m = "Unexpected error in RemoveComponentIfExists for component '" & componentName & "'." & vbCrLf & vbCrLf
    m = m & "Err: " & Err.Number & " - " & Err.Description
    Debug.Print "RemoveComponentIfExists(): " & m
    MsgBox m, vbCritical, "RemoveComponentIfExists():"
End Sub


'-------------------------------------------------------------------------------
' Procedure : InjectCodeModuleFromFile
' Purpose   : Updates the CODE of an existing component (typically a document
'             module such as ThisWorkbook or a Sheet) by reading a .cls
'             file's code body and writing it into the component's live
'             CodeModule - without deleting/recreating the component itself.
'
' Parameters:
'   vbProj        - (Object) Target VBProject.
'   componentName - (String) Name of the EXISTING component to update.
'   filePath      - (String) Path to the exported .cls file whose code body
'                   will be injected.
'
' Returns   : (Boolean) True if injection succeeded; False if the named
'             component doesn't exist, the file is missing, or an error
'             occurred.
'
' Notes     :
'   - Strips the class header block (VERSION line through the
'     "Attribute VB_Exposed" line) since that metadata belongs to the
'     component definition, not the code body - re-injecting it as code
'     would cause a compile error.
'   - Only touches CodeModule contents - the component's Type, name, and
'     binding to its underlying sheet/workbook are untouched.
'-------------------------------------------------------------------------------
Function InjectCodeModuleFromFile(ByVal vbProj As Object, ByVal componentName As String, ByVal filePath As String) As Boolean

    Dim comp As Object
    Dim codeMod As Object
    Dim fso As Object
    Dim ts As Object
    Dim fullText As String
    Dim codeBody As String
    Dim markerPos As Long
    Dim m As String

    InjectCodeModuleFromFile = False
    On Error GoTo ErrHandler

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        m = "File does not exist: " & filePath
        Debug.Print "InjectCodeModuleFromFile(): " & m
        MsgBox m, vbCritical, "InjectCodeModuleFromFile():"
        Exit Function
    End If

    On Error Resume Next
    Err.Clear
    Set comp = vbProj.VBComponents(componentName)
    On Error GoTo ErrHandler

    If comp Is Nothing Then
        m = "Target component '" & componentName & "' does not exist in the VBA project - cannot inject code into it."
        Debug.Print "InjectCodeModuleFromFile(): " & m
        MsgBox m, vbCritical, "InjectCodeModuleFromFile():"
        Exit Function
    End If

    ' Read entire file
    Set ts = fso.OpenTextFile(filePath, 1)   ' 1 = ForReading
    fullText = ts.ReadAll
    ts.Close

    If Len(fullText) = 0 Then
        m = "File '" & filePath & "' is empty - nothing to inject into '" & componentName & "'."
        Debug.Print "InjectCodeModuleFromFile(): " & m
        MsgBox m, vbCritical, "InjectCodeModuleFromFile():"
        Exit Function
    End If

    ' Strip the class header block: everything up through the line
    ' "Attribute VB_Exposed = ...". What remains is the actual code body.
    markerPos = InStr(fullText, "Attribute VB_Exposed")
    If markerPos > 0 Then
        markerPos = InStr(markerPos, fullText, vbCrLf)
        If markerPos > 0 Then
            codeBody = Mid$(fullText, markerPos + 2)
        Else
            codeBody = fullText
        End If
    Else
        Debug.Print "InjectCodeModuleFromFile(): no 'Attribute VB_Exposed' marker found in '" & filePath & "' - using full file content as-is (this may be expected if the file isn't a standard exported .cls header)."
        codeBody = fullText
    End If

    Set codeMod = comp.CodeModule

    ' Clear existing code, then write the new code body
    On Error Resume Next
    Err.Clear
    If codeMod.CountOfLines > 0 Then
        codeMod.DeleteLines 1, codeMod.CountOfLines
    End If
    If Err.Number <> 0 Then
        m = "Failed to clear existing code from '" & componentName & "' before injection." & vbCrLf & vbCrLf
        m = m & "Err: " & Err.Number & " - " & Err.Description
        Debug.Print "InjectCodeModuleFromFile(): " & m
        MsgBox m, vbCritical, "InjectCodeModuleFromFile():"
        Err.Clear
        On Error GoTo ErrHandler
        Exit Function
    End If
    On Error GoTo ErrHandler

    On Error Resume Next
    Err.Clear
    codeMod.AddFromString codeBody
    If Err.Number <> 0 Then
        m = "AddFromString failed while injecting code into '" & componentName & "'." & vbCrLf & vbCrLf
        m = m & "Source file: " & filePath & vbCrLf
        m = m & "Err: " & Err.Number & " - " & Err.Description
        Debug.Print "InjectCodeModuleFromFile(): " & m
        MsgBox m, vbCritical, "InjectCodeModuleFromFile():"
        Err.Clear
        On Error GoTo ErrHandler
        Exit Function
    End If
    On Error GoTo ErrHandler

    InjectCodeModuleFromFile = True
    Debug.Print "InjectCodeModuleFromFile(): successfully injected code into '" & componentName & "' from " & filePath
    Exit Function

ErrHandler:
    m = "Unhandled error in InjectCodeModuleFromFile for component '" & componentName & "'." & vbCrLf & vbCrLf
    m = m & "Source file: " & filePath & vbCrLf
    m = m & "Err: " & Err.Number & " - " & Err.Description
    Debug.Print "InjectCodeModuleFromFile(): " & m
    MsgBox m, vbCritical, "InjectCodeModuleFromFile():"
    InjectCodeModuleFromFile = False
End Function

'==============================================================================
' Function  : ExtractZipToFolder
' Purpose   : Extracts all contents of a ZIP file into tmpFilePath, replacing
'             any existing contents there. Works with local and UNC paths.
'             Uses PowerShell's Expand-Archive (not Shell.Namespace). Captures
'             PowerShell's actual error text to a log file so failures are
'             diagnosable instead of just returning a bare exit code.
' Inputs:   zipFilePath - String - full path to the source .zip file
'           tmpFilePath - String - full path (local or UNC) to extract into
' Returns:  Boolean - True on success, False on failure
' Notes:    fso.DeleteFolder throws Err 76 "Path not found" on UNC paths that
'           have a trailing backslash, even when FolderExists confirms the
'           folder is there. Delete must be done BEFORE adding the trailing
'           slash (or on a stripped copy of the path).
'           Err 70 ("Permission denied") is the signature of a folder being
'           locked/open elsewhere (e.g. Explorer, Excel, AV scan) - the error
'           handler below reports that case with a specific message.
'==============================================================================
Private Sub test_ExtractZipToFolder()
    MsgBox ExtractZipToFolder("C:\Users\Dwight.YACKLEY-ALIENWA\Downloads\files(2).zip", "\\NAS\bbmdata\Dwight\JUNK\VBA\NS_RPT\")
End Sub
Function ExtractZipToFolder(ByVal zipFilePath As String, ByVal tmpFilePath As String) As Boolean
    Dim m As String, ans As Long
    Dim fso As Object
    Dim wsh As Object
    Dim psCommand As String
    Dim exitCode As Long
    Dim logPath As String
    Dim logText As String
    Dim deletePath As String

    ExtractZipToFolder = False
    On Error GoTo ErrHandler

    If Not IsValidZipFile(zipFilePath) Then
        m = "Not Valid ZIP file. " & vbCrLf & vbCrLf & zipFilePath
        MsgBox m, vbCritical, "myQBAddin: ExtractZipToFolder():"
        Exit Function
    End If

    ' Makes sure tmpFilePath folder is there and empty
    CleanTmpFolder tmpFilePath
       
    ' 1. Unzip via PowerShell Expand-Archive (handles UNC paths natively).
    ' Log file for PowerShell's actual error text (written to a LOCAL temp
    ' folder, not the UNC destination, so logging itself can't fail on UNC)
    logPath = Environ$("TEMP") & "\ExtractZipToFolder_" & Format(Now, "yyyymmdd_hhnnss") & ".log"

    Set wsh = CreateObject("WScript.Shell")
    psCommand = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command " & _
                """try { Expand-Archive -LiteralPath '" & zipFilePath & _
                "' -DestinationPath '" & tmpFilePath & "' -Force -ErrorAction Stop } " & _
                "catch { $_.Exception.Message | Out-File -FilePath '" & logPath & "' -Encoding UTF8; exit 1 }"""

    exitCode = wsh.Run(psCommand, 0, True)   ' 0 = hidden window, True = wait for completion

    If exitCode <> 0 Then
        logText = "(no log file written)"
        If fso.FileExists(logPath) Then
            logText = fso.OpenTextFile(logPath, 1).ReadAll
        End If
        m = "PowerShell Expand-Archive failed with exit code " & exitCode & "." & vbCrLf & vbCrLf
        m = m & "Zip: " & zipFilePath & vbCrLf & "Dest: " & tmpFilePath & vbCrLf & vbCrLf
        m = m & "PowerShell error: " & logText
        MsgBox m, vbCritical, "myQBAddin: ExtractZipToFolder():"
        Exit Function
    End If

    ' 2. Confirm something actually landed in the destination
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(tmpFilePath) Then
        If Not (fso.GetFolder(tmpFilePath).Files.count = 0 And fso.GetFolder(tmpFilePath).SubFolders.count = 0) Then
            ExtractZipToFolder = True
        Else
            MsgBox "Extraction ran but destination folder is empty.", vbCritical, "myQBAddin: ExtractZipToFolder():"
        End If
    Else
        MsgBox "Extraction ran but destination folder does not exist.", vbCritical, "myQBAddin: ExtractZipToFolder():"
    End If

    Exit Function

ErrHandler:
    m = "Error in ExtractZipToFolder. " & vbCrLf & vbCrLf
    m = m & "Zip: " & zipFilePath & vbCrLf & "Dest: " & tmpFilePath & vbCrLf
    m = m & "Err: " & Err.Number & " - " & Err.Description

    If Err.Number = 70 Then
        m = m & vbCrLf & vbCrLf & "This error (Permission Denied) usually means the folder or a " & _
            "file inside it is currently open elsewhere - e.g. in Windows Explorer, " & _
            "Excel, or being scanned by antivirus. Close any open windows on that " & _
            "folder and try again."
    End If
    Debug.Print m
    MsgBox m, vbCritical, "myQBAddin: ExtractZipToFolder():"
    ExtractZipToFolder = False
End Function

'==============================================================================
' Procedure : CreateFolderPath
' Purpose   : Recursively creates a full folder path, including any missing
'             parent folders. Works with both local and UNC paths.
' Inputs:   fso        - Object - already-instantiated Scripting.FileSystemObject
'           folderPath - String - full folder path to create
'==============================================================================
Private Sub CreateFolderPath(ByVal fso As Object, ByVal folderPath As String)
    Dim parentPath As String

    If fso.FolderExists(folderPath) Then Exit Sub

    ' Strip trailing slash if present
    If Right$(folderPath, 1) = "\" Then
        folderPath = Left$(folderPath, Len(folderPath) - 1)
    End If

    ' Guard: don't try to recurse above a UNC server\share root (e.g. "\\NAS\bbmdata")
    ' or above a drive root (e.g. "C:\") - GetParentFolderName returns "" at the top.
    parentPath = fso.GetParentFolderName(folderPath)

    If Len(parentPath) > 0 And Not fso.FolderExists(parentPath) Then
        CreateFolderPath fso, parentPath
    End If

    On Error Resume Next
    fso.CreateFolder folderPath
    If Err.Number <> 0 And Not fso.FolderExists(folderPath) Then
        Dim errNum As Long, errDesc As String
        errNum = Err.Number: errDesc = Err.Description
        On Error GoTo 0
        Err.Raise errNum, "CreateFolderPath", "Failed creating '" & folderPath & "': " & errDesc
    End If
    On Error GoTo 0
End Sub

'==============================================================================
' Function  : IsValidZipFile
' Purpose   : Determines whether filePath points to a valid ZIP archive.
'             Validates existence (via FSO, not Dir$), extension, and the
'             binary ZIP signature.
' Inputs:   filePath - String - full path to the file to test
' Returns:  Boolean - True if a valid ZIP signature is found, else False
'==============================================================================
Function IsValidZipFile(ByVal filePath As String) As Boolean
    Dim fso As Object
    Dim fileNum As Integer
    Dim sig(0 To 3) As Byte
    Dim result As Boolean

    result = False
    On Error GoTo CleanFail

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 1. Must exist (FSO check instead of Dir$)
    If Not fso.FileExists(filePath) Then
        IsValidZipFile = False
        Exit Function
    End If

    ' 2. Must have a .zip extension
    If LCase$(Right$(filePath, 4)) <> ".zip" Then
        IsValidZipFile = False
        Exit Function
    End If

    ' 3. Must be at least 4 bytes long (FSO Size instead of FileLen)
    If fso.GetFile(filePath).Size < 4 Then
        IsValidZipFile = False
        Exit Function
    End If

    ' 4. Read first 4 bytes and check against known ZIP signatures
    fileNum = FreeFile
    Open filePath For Binary Access Read As #fileNum
    Get #fileNum, 1, sig
    Close #fileNum

    If sig(0) = &H50 And sig(1) = &H4B And sig(2) = &H3 And sig(3) = &H4 Then
        result = True                                      ' "PK\3\4" standard
    ElseIf sig(0) = &H50 And sig(1) = &H4B And sig(2) = &H5 And sig(3) = &H6 Then
        result = True                                      ' "PK\5\6" empty archive
    ElseIf sig(0) = &H50 And sig(1) = &H4B And sig(2) = &H7 And sig(3) = &H8 Then
        result = True                                      ' "PK\7\8" spanned archive
    End If

    IsValidZipFile = result
    Exit Function

CleanFail:
    If fileNum <> 0 Then
        On Error Resume Next
        Close #fileNum
    End If
    IsValidZipFile = False
End Function





