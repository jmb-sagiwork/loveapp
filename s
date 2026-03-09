Private Sub RunXnetx()
    Dim objShellPythonXnet As Object
    Dim pyExe As String
    Dim pyFile As String
    Dim cmd As String
    Dim exitCode As Long

    Set objShellPythonXnet = CreateObject("WScript.Shell")

    pyExe = "C:\Users\C8S7HD\AppData\Local\Programs\Python\Python313\python.exe"
    pyFile = ThisWorkbook.Path & "\Xnet6.py"

    If Dir(pyExe) = "" Then
        MsgBox "python.exe not found:" & vbCrLf & pyExe, vbCritical
        Exit Sub
    End If

    If Dir(pyFile) = "" Then
        MsgBox "Python script not found:" & vbCrLf & pyFile, vbCritical
        Exit Sub
    End If

    cmd = """" & pyExe & """ """ & pyFile & """"

    Debug.Print cmd
    exitCode = objShellPythonXnet.Run(cmd, 1, True)

    Debug.Print "Exit Code: " & exitCode
    MsgBox "Exit Code: " & exitCode

    Set objShellPythonXnet = Nothing
End Sub
