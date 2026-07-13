Attribute VB_Name = "probe_capped"
Option Explicit
Option Private Module
'@name Capped Worker Probe
'@ribbon Capped Worker Probe
'@group Internal
'@maxjobs 2
'@internal true

Public Sub Run(ByVal ctx As InfraContext)
    Call ctx.RunJob("probe_capped.Worker", "{}")
End Sub

Public Sub Worker(ByVal job As InfraJob)
    probe_worker.Worker job
End Sub
