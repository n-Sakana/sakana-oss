Attribute VB_Name = "probe_ok"
Option Explicit
Option Private Module
'@name Harness Probe OK
'@ribbon Harness Probe OK
'@group Internal
'@maxjobs 1
'@internal true

Public Sub Run(ByVal ctx As InfraContext)
    ctx.LogMessage "probe ok"
    ctx.Status 0.5, "probe status"
End Sub

