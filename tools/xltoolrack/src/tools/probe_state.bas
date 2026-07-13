Attribute VB_Name = "probe_state"
Option Explicit
Option Private Module
'@name Harness Probe State
'@ribbon Harness Probe State
'@group Internal
'@maxjobs 1
'@internal true

Public Sub Run(ByVal ctx As InfraContext)
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
End Sub

