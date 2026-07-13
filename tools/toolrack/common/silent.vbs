' silent.vbs -- start launch.ps1 with no visible console (gui-type tools).
' args: 0=tool dir, 1=variant index, 2=target path
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
p = fso.GetParentFolderName(WScript.ScriptFullName)
cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & p & "\launch.ps1""" & _
      " -Gui -Tool """ & WScript.Arguments(0) & """" & _
      " -Variant " & WScript.Arguments(1) & _
      " -Target """ & WScript.Arguments(2) & """"
sh.Run cmd, 0, False
