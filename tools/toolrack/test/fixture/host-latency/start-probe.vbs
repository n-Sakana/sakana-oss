' start-probe.vbs -- windowless launcher for the workplace host probe.
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
base = fso.GetParentFolderName(WScript.ScriptFullName)
ps = sh.ExpandEnvironmentStrings("%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe")
cmd = """" & ps & """ -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & _
      base & "\probe-global-host.ps1"" -Run -StateRoot """ & base & """"
sh.Run cmd, 0, False
