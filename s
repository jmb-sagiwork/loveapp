Private Sub RunXnetx()
    Dim sh As Object
    Dim pyExe As String
    Dim pyFile As String
    Dim cmd As String

    Set sh = CreateObject("WScript.Shell")

    pyExe = "C:\Users\C8S7HD\AppData\Local\Programs\Python\Python314\python.exe"
    pyFile = ThisWorkbook.Path & "\Xnet6.py"

    cmd = "cmd /k ""cd /d """ & ThisWorkbook.Path & """ && """ & pyExe & """ """ & pyFile & """"""

    Debug.Print cmd
    sh.Run cmd, 1, True
End Sub
