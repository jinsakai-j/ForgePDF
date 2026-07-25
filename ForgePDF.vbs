Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "C:\Users\uSer\ForgePDF"
WshShell.Run """C:\Users\uSer\AppData\Local\Programs\Python\Python314\pythonw.exe"" main.py", 0, False
