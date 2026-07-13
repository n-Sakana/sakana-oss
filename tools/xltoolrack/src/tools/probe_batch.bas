Attribute VB_Name = "probe_batch"
Option Explicit
Option Private Module
'@name Batch Rollback Probe
'@ribbon Batch Rollback Probe
'@group Internal
'@maxjobs 3
'@internal true

Public Sub Run(ByVal ctx As InfraContext)
    Call ctx.RunJob("probe_batch.Worker", "{}")
    Call ctx.RunJob("probe_batch.Worker", "{}")
    Call ctx.RunJob("probe_batch.Worker", "{}")
End Sub

Public Sub Worker(ByVal job As InfraJob)
    probe_worker.Worker job
End Sub

