VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents App As Excel.Application
Private m_host As Object

Public Sub Init(ByVal applicationValue As Excel.Application, ByVal host As Object)
    Set App = applicationValue
    Set m_host = host
End Sub

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo CleanExit
    If m_host Is Nothing Then Exit Sub
    m_host.HandleSheetChange Sh, Target
CleanExit:
    If Err.Number <> 0 Then Debug.Print "!! AppEventHandler.SheetChange | Err " & CStr(Err.Number)
End Sub

Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    On Error Resume Next
    If Not m_host Is Nothing Then m_host.HandleWorkbookBeforeClose Wb
    On Error GoTo 0
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    On Error Resume Next
    If Not m_host Is Nothing Then m_host.NoteWorkbookActivated Wb
    On Error GoTo 0
End Sub

