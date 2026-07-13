Attribute VB_Name = "probe_worker_error"
Option Explicit
Option Private Module
'@name Error Worker Probe
'@ribbon Error Worker Probe
'@group Internal
'@maxjobs 1
'@internal true

Public Sub Run(ByVal ctx As InfraContext)
    Call ctx.RunJob("probe_worker_error.Worker", "{}")
End Sub

Public Sub Worker(ByVal job As InfraJob)
    Err.Raise 5, "probe_worker_error.Worker", "probe worker failure"
End Sub

