Attribute VB_Name = "my_Exec_Mod"
Option Explicit
' This Mod contains my Exec Macros and Functions

Private Sub test_myOpenCompanyFiles()
Dim s As String
    s = myOpenCompanyFiles()
    MsgBox "'" & s & "' len = " & Len(s)
End Sub

Function myOpenCompanyFiles() As String
Dim scriptstr As String, stdout, m As String   ' M as variant!
Dim stdOutArray() As String, i As Long

' The following line is PC specific.
    scriptstr = "C:\Users\Dwight.YACKLEY-ALIENWA\OneDrive - BBM Incorporated\my PS\function Get-OpenQBWCountAndNames.ps1"
    stdout = myRunPowerShellScriptAndGetOutput(scriptstr)
    
    If stdout = "" Then
        m = "NO output received back from Pwsh exec!"
        myOpenCompanyFiles = ""
        GoTo exit_func
    End If
    
    ' Use the Split function to parse the string into an array
    ' The comma (",") is specified as the delimiter
    stdOutArray = Split(stdout, ",")
   
    ' First element is how many files are open
    If stdOutArray(0) = CStr(0) Then
        m = "No company qb file(s) open."
        myOpenCompanyFiles = ""
        GoTo exit_func
    End If
    m = "Element 0: " & CStr(stdOutArray(0)) & vbCrLf
    
    ' next elements are the open file names
    For i = 1 To UBound(stdOutArray)
        m = m & "Element " & i & ": " & stdOutArray(i) & vbCrLf
        ' in .ps1 file ADDED an extra ',' at end of stdout TO FIX MISSING '"' ON LAST Element in array stdOutArray(i)
        myOpenCompanyFiles = myOpenCompanyFiles & UCase(stdOutArray(i)) & "," ' make a comma separated list to return
    Next i
    myOpenCompanyFiles = Left(myOpenCompanyFiles, Len(myOpenCompanyFiles) - 0) ' Strip off last ','
    
exit_func:
    ' Release the memory occupied by the dynamic array
    Erase stdOutArray
    Debug.Print m
    Debug.Print myOpenCompanyFiles
End Function

Private Sub test_IsMyCompanyFileOpen()
    MsgBox IsMyCompanyFileOpen("BBMDY")
End Sub

Function IsMyCompanyFileOpen(company_file As String) As Boolean
Dim scriptstr As String, stdout, m As String   ' M as variant!
Dim stdOutArray() As String, i As Long

' The following line is PC specific.
    scriptstr = "C:\Users\Dwight.YACKLEY-ALIENWA\OneDrive - BBM Incorporated\my PS\function Get-OpenQBWCountAndNames.ps1"
    stdout = myRunPowerShellScriptAndGetOutput(scriptstr)
    
    If stdout = "" Then
        m = "NO output received back from Pwsh exec!"
        IsMyCompanyFileOpen = False
        GoTo exit_func
    End If
    
    ' Use the Split function to parse the string into an array
    ' The comma (",") is specified as the delimiter
    stdOutArray = Split(stdout, ",")
   
    ' First element is how many files are open
    If stdOutArray(0) = CStr(0) Then
        m = "No company qb file(s) open."
        IsMyCompanyFileOpen = False
        GoTo exit_func
    End If
    m = "Element 0: " & CStr(stdOutArray(0)) & vbCrLf
    
    ' next elements are the open file names
    For i = 1 To UBound(stdOutArray)
        m = m & "Element " & i & ": " & stdOutArray(i) & vbCrLf
        ' in .ps1 file ADDED an extra ',' at end of stdout TO FIX MISSING '"' ON LAST Element in array stdOutArray(i)
        If UCase(company_file) = UCase(stdOutArray(i)) Then IsMyCompanyFileOpen = True
    Next i
    
exit_func:
    ' Release the memory occupied by the dynamic array
    Erase stdOutArray
    Debug.Print m
End Function

' Uses script Get-NotReadOnlyQBWFiles.ps1  to see whch QB files are open
Function WhichQBCompaniesAreOpen()
Dim scriptstr As String, stdout
Dim stdOutArray() As String
Dim i As Long
Dim m As String
' The following line is PC specific.
    scriptstr = "C:\Users\Dwight.YACKLEY-ALIENWA\OneDrive - BBM Incorporated\my PS\function Get-NotReadOnlyQBWFiles.ps1"
    stdout = myRunPowerShellScriptAndGetOutput(scriptstr)

    If stdout = "" Then
        m = "NO output received back!"
        MsgBox m, , "WhichQBCompaniesAreOpen()": Exit Function
    End If
    
    ' Use the Split function to parse the string into an array
    ' The comma (",") is specified as the delimiter
    stdOutArray = Split(stdout, ",")

    ' Loop through the array and display each parsed element
    For i = LBound(stdOutArray) To UBound(stdOutArray)
        Debug.Print "Element " & i & ": " & stdOutArray(i)
    Next i
    
    m = stdOutArray(0) & " file(s) open." & vbCrLf
    For i = 1 To UBound(stdOutArray)
        m = m & stdOutArray(i) & vbCrLf
    Next i

    ' Release the memory occupied by the dynamic array
    Erase stdOutArray
    MsgBox m
End Function
' FYI......THIS SCRIPT RUNS VERY SLOW!
Private Sub test_CountExcelProcesses()
Dim cnt As Long, m As String
    cnt = CountExcelProcesses
    If cnt > 1 Then
        m = "Warning: Multiple Excel Processes is Running! cnt = " & cnt & vbCrLf & vbCrLf
        m = m & "Check Immediate Window for more information!"
        MsgBox m, vbCritical + vbInformation, "CountExcelProcesses()"
    Else
        MsgBox "All good. " & cnt, , "CountExcelProcesses()"
    End If
End Sub
' FYI......THIS SCRIPT RUNS VERY SLOW!
' Uses script function Show-ExcelProcesses.ps1 to see whch QB files are open
Function CountExcelProcesses() As Long
Dim scriptstr As String, stdout
Dim stdOutArray() As String
Dim i As Long
Dim m As String
' The following line is PC specific.
    scriptstr = "C:\Users\Dwight.YACKLEY-ALIENWA\OneDrive - BBM Incorporated\my PS\function Show-ExcelProcesses.ps1"
    stdout = myRunPowerShellScriptAndGetOutput(scriptstr)

    If stdout = "" Then
        m = "NO output received back from CountExcelProcesses()!"
        MsgBox m: Exit Function
    End If
    
    ' Use the Split function to parse the string into an array
    ' The comma (",") is specified as the delimiter
    stdOutArray = Split(stdout, ",")

    CountExcelProcesses = stdOutArray(0) ' The count of running EXCEL processes
    
    ' Loop through the array and display each parsed element
    For i = LBound(stdOutArray) To UBound(stdOutArray)
        Debug.Print "Element " & i & ": " & stdOutArray(i)
    Next i

    ' Release the memory occupied by the dynamic array
    Erase stdOutArray
End Function

Private Sub test_myRunPowerShellScriptAndGetOutput()
Dim scriptstr As String, stdout
' The following line is PC specific.
    scriptstr = "C:\Users\Dwight.YACKLEY-ALIENWA\OneDrive - BBM Incorporated\my PS\function Get-NotReadOnlyQBWFiles.ps1"
    MsgBox myRunPowerShellScriptAndGetOutput(scriptstr)
End Sub

'
'   My powershell run function
'
Function myRunPowerShellScriptAndGetOutput(strScriptPath As String)
    Dim objShell As Object
    Dim objExec As Object

    ' Create a WScript.Shell object
    Set objShell = CreateObject("WScript.Shell")

    ' Execute the PowerShell script and capture the output.
    '
    ' NoLogo:    Suppresses the copyright banner that appears when a new PowerShell session starts.
    ' NoProfile: Prevents PowerShell from loading the user profile, which can sometimes contain additional messages
    ' ExecutionPolicy Bypass: Ensures the automation does not fail due to security settings.
    ' "-File" argument ensures the script is executed as a file
    Set objExec = objShell.Exec("powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File """ & strScriptPath & """")
  
    ' Read all the standard output from the PowerShell script
    myRunPowerShellScriptAndGetOutput = objExec.stdout.ReadAll
    
    ' Clean up objects
    Set objExec = Nothing
    Set objShell = Nothing

    ' Returns standard output
    'MsgBox myRunPowerShellScriptAndGetOutput
End Function
Sub test_StdOut(objExec As Object)
Dim line As String, output As String
        line = objExec.stdout.ReadLine
        Debug.Print line
    
End Sub

' Testing  function Get-NotReadOnlyQBWFiles.ps1 file
Private Sub test_NotReadOnlyQBWFiles()
Dim scriptstr As String, stdout
Dim stdOutArray() As String
Dim i As Long
Dim m As String
' The following line is PC specific.
    scriptstr = "C:\Users\Dwight.YACKLEY-ALIENWA\OneDrive - BBM Incorporated\my PS\function Get-NotReadOnlyQBWFiles.ps1"
    stdout = myRunPowerShellScriptAndGetOutput(scriptstr)

    If stdout = "" Then
        m = "NO output received back!"
        MsgBox m: Exit Sub
    End If
    
    ' Use the Split function to parse the string into an array
    ' The comma (",") is specified as the delimiter
    stdOutArray = Split(stdout, ",")

    ' Loop through the array and display each parsed element
    For i = LBound(stdOutArray) To UBound(stdOutArray)
        Debug.Print "Element " & i & ": " & stdOutArray(i)
    Next i
    
    m = stdOutArray(0) & " file(s) open." & vbCrLf
    For i = 1 To UBound(stdOutArray)
        m = m & stdOutArray(i) & vbCrLf
    Next i

    ' Release the memory occupied by the dynamic array
    Erase stdOutArray
    MsgBox m
End Sub

' Testing  function CountExcel.ps1 file

Private Sub test_Show_ExcelProcesses()
Dim scriptstr As String, stdout
Dim stdOutArray() As String
Dim i As Long
Dim m As String
' The following line is PC specific.
    scriptstr = "C:\Users\Dwight.YACKLEY-ALIENWA\OneDrive - BBM Incorporated\my PS\function Show-ExcelProcesses.ps1"
    stdout = myRunPowerShellScriptAndGetOutput(scriptstr)

    If stdout = "" Then
        m = "NO output received back!"
        MsgBox m: Exit Sub
    End If
    
    ' Use the Split function to parse the string into an array
    ' The comma (",") is specified as the delimiter
    stdOutArray = Split(stdout, ",")

    ' Loop through the array and display each parsed element
    For i = LBound(stdOutArray) To UBound(stdOutArray)
        Debug.Print "Element " & i & ": " & stdOutArray(i)
    Next i
    
    m = stdOutArray(0) & " EXCEL Process(es) open." & vbCrLf
    For i = 1 To UBound(stdOutArray)
        m = m & stdOutArray(i) & vbCrLf
    Next i

    ' Release the memory occupied by the dynamic array
    Erase stdOutArray
    MsgBox m
End Sub


' Not used - currently
Private Function IsStringArrayEmpty(ByRef arr As Variant) As Boolean
    On Error GoTo ErrHandler
    If UBound(arr) < LBound(arr) Then
        IsStringArrayEmpty = True: Debug.Print "Array is dimensioned but empty."
    Else
        IsStringArrayEmpty = False: Debug.Print "Array contains elements."
    End If
    Exit Function

ErrHandler:
    If Err.Number = 9 Then ' Subscript out of range error, indicating unallocated array
        IsStringArrayEmpty = True: Debug.Print "Subscript out of range error, indicating unallocated array."
    Else
        ' Handle other potential errors if necessary
        Err.Clear
        IsStringArrayEmpty = True: Debug.Print "Assumed empty if other errors occur."
    End If
End Function

' Testing different powershell scripts
Private Sub test1_myRunPowerShellScriptAndGetOutput()
    Dim objShell As Object
    Dim objExec As Object
    Dim strScriptPath As String
    Dim strOutput As String

    ' Define the path to your PowerShell script
    ' Replace "C:\Path\To\YourScript.ps1" with the actual path to your .ps1 file
' The following line is PC specific.
    strScriptPath = "C:\Users\Dwight.YACKLEY-ALIENWA\OneDrive - BBM Incorporated\my PS\function Get-NotReadOnlyQBWFiles.ps1"

    ' Create a WScript.Shell object
    Set objShell = CreateObject("WScript.Shell")

    ' Execute the PowerShell script and capture the output
    ' The "-File" argument ensures the script is executed as a file
    'Set objExec = objShell.Exec("powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File strScriptPath")
    Set objExec = objShell.Exec("powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File """ & strScriptPath & """")
    'Set objExec = objShell.Exec("powershell.exe -NoLogo -NoProfile -Command ""Get-Service | Select-Object Name, Status""")

    
    ' Read all the standard output from the PowerShell script
    strOutput = objExec.stdout.ReadAll

    ' Display the captured output (you can modify this to suit your needs)
    MsgBox "PowerShell Script Output:" & vbCrLf & strOutput

    ' Clean up objects
    Set objExec = Nothing
    Set objShell = Nothing

End Sub

' Testing different ways to invoke Powershell
Private Sub test_()
Dim objShell As Object
Dim objExec As Object
Dim strOutput

    Set objShell = CreateObject("WScript.Shell")
    ' Must keep  -NoLogo and -NoProfile in powershell command
    Set objExec = objShell.Exec("powershell.exe -NoLogo -NoProfile -Command ""Get-Service | Select-Object Name, Status""")
    
    ' Wait for the command to finish
'    Do Until objExec.Status = 1
'        WScript.Sleep 100
'    Loop
    
    ' Capture and display the output
    strOutput = objExec.stdout.ReadAll
    ' Display the captured output (you can modify this to suit your needs)
    MsgBox "PowerShell Script Output:" & vbCrLf & strOutput

    ' Clean up objects
    Set objExec = Nothing
    Set objShell = Nothing

End Sub


