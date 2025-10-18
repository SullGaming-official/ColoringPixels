Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
vbsFolder = fso.GetParentFolderName(WScript.ScriptFullName)
htaPath = fso.BuildPath(vbsFolder & "\Launcher", "Launcher.hta")
shell.Run "mshta.exe """ & fso.GetAbsolutePathName(htaPath) & """"