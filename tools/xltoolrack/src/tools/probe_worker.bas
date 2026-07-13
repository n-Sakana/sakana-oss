Attribute VB_Name = "probe_worker"
Option Explicit
Option Private Module
'@name Worker Probe
'@ribbon Worker Probe
'@group Internal
'@maxjobs 3
'@internal true

Private m_tick As Long

Public Sub Run(ByVal ctx As InfraContext)
    Call ctx.RunJob("probe_worker.Worker", "{}")
End Sub

Public Sub Worker(ByVal job As InfraJob)
    Dim values(1 To 1, 1 To 2) As Variant
    ' A rejected push means the FE is busy: skip this tick, keep running.
    values(1, 1) = m_tick + 1
    values(1, 2) = CDbl(Now)
    If job.Push("state", "A1:B1", values) Then m_tick = m_tick + 1
End Sub

