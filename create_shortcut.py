import os
import win32com.client

desktop_path = r"C:\Users\uSer\OneDrive\Desktop"

# Remove old shortcuts
for f in os.listdir(desktop_path):
    if "PDF" in f or "Forge" in f:
        try:
            os.remove(os.path.join(desktop_path, f))
        except Exception:
            pass

# Create clean new shortcut
shell = win32com.client.Dispatch("WScript.Shell")
shortcut = shell.CreateShortcut(os.path.join(desktop_path, "ForgePDF.lnk"))
shortcut.TargetPath = r"C:\Windows\System32\wscript.exe"
shortcut.Arguments = r'"C:\Users\uSer\ForgePDF\ForgePDF.vbs"'
shortcut.WorkingDirectory = r"C:\Users\uSer\ForgePDF"
shortcut.Save()

print("Clean ForgePDF shortcut created successfully on Desktop!")
