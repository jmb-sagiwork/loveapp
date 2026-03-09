Private Sub RunXnet()
    Dim objShellPythonXnet As Object
    Dim pyExe As String
    Dim pyFile As String
    Dim cmd As String
    Dim exitCode As Long

    Set objShellPythonXnet = CreateObject("WScript.Shell")

    pyExe = "C:\Users\YourName\AppData\Local\Programs\Python\Python313\python.exe"
    pyFile = ThisWorkbook.Path & "\Xnet6.py"

    cmd = Chr(34) & pyExe & Chr(34) & " " & Chr(34) & pyFile & Chr(34)

    Debug.Print cmd
    exitCode = objShellPythonXnet.Run(cmd, 1, True)

    Debug.Print "Exit Code: " & exitCode

    Set objShellPythonXnet = Nothing
End Sub
