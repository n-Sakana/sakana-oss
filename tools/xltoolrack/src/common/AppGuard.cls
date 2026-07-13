VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppGuard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_prevScreen As Boolean
Private m_prevAlerts As Boolean
Private m_prevCalc As Long
Private m_haveScreen As Boolean
Private m_haveAlerts As Boolean
Private m_haveCalc As Boolean
Private m_active As Boolean

Private Sub Class_Initialize()
    On Error GoTo InitFailed

    m_prevScreen = Application.ScreenUpdating
    m_haveScreen = True
    m_prevAlerts = Application.DisplayAlerts
    m_haveAlerts = True
    m_prevCalc = Application.Calculation
    m_haveCalc = True

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    m_active = True
    Exit Sub

InitFailed:
    RestoreSaved
    Err.Raise Err.Number, "AppGuard.Class_Initialize", Err.Description
End Sub

Public Sub Restore()
    If Not m_active Then Exit Sub
    RestoreSaved
    m_active = False
End Sub

Private Sub RestoreSaved()
    On Error Resume Next
    If m_haveCalc Then Application.Calculation = m_prevCalc
    If m_haveAlerts Then Application.DisplayAlerts = m_prevAlerts
    If m_haveScreen Then Application.ScreenUpdating = m_prevScreen
    On Error GoTo 0
End Sub

Private Sub Class_Terminate()
    Restore
End Sub

