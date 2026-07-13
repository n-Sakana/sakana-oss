' common/host-start.vbs -- start the local PowerShell Host bootstrap windowless.
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
base = fso.GetParentFolderName(WScript.ScriptFullName)
ps = sh.ExpandEnvironmentStrings("%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe")
cmd = """" & ps & """ -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & _
      base & "\host-bootstrap.ps1"""
sh.Run cmd, 0, False
