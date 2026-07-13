VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_path As String
Private m_started As Boolean

Public Sub Begin(ByVal toolId As String, Optional ByVal jobId As String = vbNullString)
    Dim root As String
    Dim suffix As String

    root = LogRoot()
    EnsureFolder root
    suffix = Format$(Now, "yyyymmdd_hhnnss") & "_" & Format$(CLng(Timer * 1000!) Mod 1000, "000")
    If Len(jobId) > 0 Then suffix = suffix & "_" & SafeName(jobId)
    m_path = root & Application.PathSeparator & SafeName(toolId) & "_" & suffix & ".log"
    m_started = True
    AppendLine "BEGIN tool=" & toolId & IIf(Len(jobId) > 0, " job=" & jobId, vbNullString)
End Sub

Public Sub Attach(ByVal existingPath As String)
    m_path = existingPath
    m_started = (Len(m_path) > 0)
End Sub

Public Sub WriteLine(ByVal message As String, Optional ByVal level As String = "INFO")
    If Not m_started Then Exit Sub
    AppendLine UCase$(level) & " " & message
End Sub

Public Sub Done()
    If Not m_started Then Exit Sub
    AppendLine "DONE"
    m_started = False
End Sub

Public Sub Fail(ByVal message As String)
    If Len(m_path) = 0 Then Exit Sub
    AppendLine "FAIL " & message
    m_started = False
End Sub

Public Property Get Path() As String
    Path = m_path
End Property

Private Sub AppendLine(ByVal text As String)
    Dim fileNo As Integer
    On Error GoTo CleanExit
    fileNo = FreeFile
    Open m_path For Append As #fileNo
    Print #fileNo, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " " & text
CleanExit:
    On Error Resume Next
    If fileNo > 0 Then Close #fileNo
    On Error GoTo 0
End Sub

Private Function LogRoot() As String
    Dim configured As String
    configured = Environ$("XLTOOLRACK_LOG_DIR")
    If Len(configured) > 0 Then
        LogRoot = configured
    Else
        LogRoot = Environ$("LOCALAPPDATA") & Application.PathSeparator & "xltoolrack" & Application.PathSeparator & "log"
    End If
End Function

Private Sub EnsureFolder(ByVal folderPath As String)
    Dim fso As Object
    Dim parentPath As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderPath) Then Exit Sub
    parentPath = fso.GetParentFolderName(folderPath)
    If Len(parentPath) > 0 Then
        If Not fso.FolderExists(parentPath) Then EnsureFolder parentPath
    End If
    fso.CreateFolder folderPath
End Sub

Private Function SafeName(ByVal value As String) As String
    Dim i As Long
    Dim ch As String
    Dim result As String
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or _
           (ch >= "0" And ch <= "9") Or ch = "_" Or ch = "-" Then
            result = result & ch
        Else
            result = result & "_"
        End If
    Next i
    If Len(result) = 0 Then result = "run"
    SafeName = result
End Function
