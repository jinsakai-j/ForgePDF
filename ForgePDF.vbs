Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")

' Folder skrip = folder ForgePDF (bergerak bersama folder app, tidak hardcode)
strDir = fso.GetParentFolderName(WScript.ScriptFullName)
sh.CurrentDirectory = strDir

' Cari pythonw.exe: cek PATH dulu, baru lokasi umum
sql = ""
For Each e In Split(sh.ExpandEnvironmentStrings("%PATH%"), ";")
    If e <> "" Then
        If Right(e, 1) <> "\" Then e = e & "\"
        If fso.FileExists(e & "pythonw.exe") Then
            sPyW = e & "pythonw.exe"
            Exit For
        End If
    End If
Next

If sPyW = "" Then
    cands = Array( _
        "%LocalAppData%\Programs\Python\Python314\pythonw.exe", _
        "%LocalAppData%\Programs\Python\Python313\pythonw.exe", _
        "%LocalAppData%\Programs\Python\Python312\pythonw.exe", _
        "%ProgramFiles%\Python313\pythonw.exe", _
        "%ProgramFiles%\Python312\pythonw.exe", _
        "%ProgramFiles%\Python311\pythonw.exe", _
        "%ProgramFiles(x86)%\Python311\pythonw.exe", _
        "C:\Python311\pythonw.exe", _
        "C:\Python312\pythonw.exe")
    For Each c In cands
        Dim fp
        fp = sh.ExpandEnvironmentStrings(c)
        If fso.FileExists(fp) Then
            sPyW = fp
            Exit For
        End If
    Next
End If

If sPyW = "" Then
    MsgBox "pythonw.exe tidak ditemukan di PATH maupun lokasi umum." & vbCrLf & _
           "Install Python (centang 'Add to PATH') lalu coba lagi.", vbCritical, "ForgePDF"
    WScript.Quit 1
End If

sh.Run """" & sPyW & """ main.py", 0, False