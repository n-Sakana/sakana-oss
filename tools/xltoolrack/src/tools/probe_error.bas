Attribute VB_Name = "probe_error"
Option Explicit
Option Private Module
'@name Harness Probe Error
'@ribbon Harness Probe Error
'@group Internal
'@maxjobs 1
'@internal true

Public Sub Run(ByVal ctx As InfraContext)
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Err.Raise 5, "probe_error.Run", "probe failure"
End Sub
