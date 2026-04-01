Set shell = CreateObject("WScript.Shell")
scriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
shell.Run "powershell -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File """ & scriptDir & "\launch.ps1""", 0, False
