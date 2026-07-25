Set WshShell = CreateObject("WScript.Shell")echo WshShell.CurrentDirectory = "C:\Users\uSer\ForgePDF"
WshShell.Run "pythonw.exe main.py", 0, False
