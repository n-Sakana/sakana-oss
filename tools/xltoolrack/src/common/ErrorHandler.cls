VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Report(ByVal toolId As String, ByVal number As Long, ByVal description As String, _
                  Optional ByVal asynchronous As Boolean = False)
    Dim message As String
    message = toolId & " | Err " & CStr(number) & ": " & description
    Debug.Print "!! " & message

    If Environ$("XLTOOLRACK_TEST") = "1" Then Exit Sub
    If asynchronous Then
        On Error Resume Next
        Application.StatusBar = "xltoolrack error: " & message
        On Error GoTo 0
    ElseIf Application.Visible Then
        MsgBox message, vbExclamation, "xltoolrack"
    End If
End Sub

