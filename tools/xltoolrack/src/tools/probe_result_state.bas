Attribute VB_Name = "probe_result_state"
Option Explicit
Option Private Module
'@name Result State Probe
'@ribbon Result State Probe
'@group Internal
'@maxjobs 1
'@internal true

Private m_tick As Long

Public Sub Run(ByVal ctx As InfraContext)
    Call ctx.RunJob("probe_result_state.Worker", "{}")
End Sub

Public Sub Worker(ByVal job As InfraJob)
    m_tick = m_tick + 1
    If Not job.Push("state", "A1", m_tick) Then Err.Raise vbObjectError + 2820, "probe_result_state.Worker", "push failed"
End Sub

Public Sub OnResult(ByVal ctx As InfraContext, ByVal jobId As String, ByVal version As Long)
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.Interactive = False
    ctx.Target.Worksheets(1).Range("A1").Value2 = version
End Sub
