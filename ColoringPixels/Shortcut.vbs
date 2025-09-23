Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' 🧱 Match your launcher path logic
vbsFolder = fso.GetParentFolderName(WScript.ScriptFullName)
htaPath = fso.BuildPath(vbsFolder & "\Launcher", "Launcher.hta")

' 🖼️ Create desktop shortcut if missing
desktopPath = shell.SpecialFolders("Desktop")
shortcutPath = desktopPath & "\ColoringPixels Launcher.lnk"

If Not fso.FileExists(shortcutPath) Then
  Set shortcut = shell.CreateShortcut(shortcutPath)
  shortcut.TargetPath = "mshta.exe"
  shortcut.Arguments = """" & fso.GetAbsolutePathName(htaPath) & """"
  shortcut.WindowStyle = 1
  shortcut.IconLocation = htaPath & ",0"
  shortcut.Description = "Launch ColoringPixels with SullGaming rituals"
  shortcut.Save
End If
WScript.Quit