VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_previous As Variant
Private m_active As Boolean

Public Sub Begin(ByVal toolId As String)
    If m_active Then Exit Sub
    m_previous = Application.StatusBar
    m_active = True
    Application.StatusBar = "xltoolrack: " & toolId
End Sub

Public Sub Show(ByVal fraction As Double, ByVal message As String)
    If Not m_active Then Begin vbNullString
    Application.StatusBar = message
End Sub

Public Sub Clear()
    If Not m_active Then Exit Sub
    On Error Resume Next
    Application.StatusBar = m_previous
    m_active = False
    On Error GoTo 0
End Sub

Private Sub Class_Terminate()
    Clear
End Sub

