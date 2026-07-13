VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppStateSnapshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Cursor is deliberately excluded. During an OnTime callback Excel can expose
' a transient busy value through its cursor property; saving and restoring it
' value pins the transient state and destroys Excel's contextual cell cursors.

Private m_prevScreen As Boolean
Private m_prevAlerts As Boolean
Private m_prevCalc As Long
Private m_prevEvents As Boolean
Private m_prevInteractive As Boolean
Private m_active As Boolean

Private Sub Class_Initialize()
    On Error GoTo InitFailed
    m_prevScreen = Application.ScreenUpdating
    m_prevAlerts = Application.DisplayAlerts
    m_prevCalc = Application.Calculation
    m_prevEvents = Application.EnableEvents
    m_prevInteractive = Application.Interactive
    m_active = True
    Exit Sub
InitFailed:
    Err.Raise Err.Number, "AppStateSnapshot.Class_Initialize", Err.Description
End Sub

Public Sub Restore()
    If Not m_active Then Exit Sub
    RestoreSaved
    m_active = False
End Sub

' Restore the captured baseline without consuming the snapshot. The FE pump
' uses one snapshot for a whole batch of result callbacks.
Public Sub RestoreForReuse()
    If Not m_active Then Exit Sub
    RestoreSaved
End Sub

Private Sub RestoreSaved()
    On Error Resume Next
    If Application.Interactive <> m_prevInteractive Then Application.Interactive = m_prevInteractive
    If Application.EnableEvents <> m_prevEvents Then Application.EnableEvents = m_prevEvents
    If Application.Calculation <> m_prevCalc Then Application.Calculation = m_prevCalc
    If Application.DisplayAlerts <> m_prevAlerts Then Application.DisplayAlerts = m_prevAlerts
    If Application.ScreenUpdating <> m_prevScreen Then Application.ScreenUpdating = m_prevScreen
    On Error GoTo 0
End Sub

Private Sub Class_Terminate()
    Restore
End Sub
