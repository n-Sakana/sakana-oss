VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JobWorker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_job As InfraJob
Private m_toolId As String
Private m_donePath As String
Private m_exitReason As String
Private m_selfQuit As Boolean
Private m_nextDue As Date

Public Sub Init(ByVal jobId As String, ByVal toolId As String, ByVal argsJson As String, _
                ByVal donePath As String, ByVal token As String, ByVal channelBase As String, _
                ByVal logPath As String, ByVal feGonePath As String, ByVal feLeasePath As String)
    Set m_job = New InfraJob
    m_job.InitRemote jobId, argsJson, token, channelBase, logPath, feGonePath, feLeasePath
    m_toolId = toolId
    m_donePath = donePath
    m_exitReason = vbNullString
    m_selfQuit = False
    Application.DisplayAlerts = False
    Application.EnableEvents = True
End Sub

Public Sub RunLoop()
    Dim failureNumber As Long
    Dim failureDescription As String

    On Error GoTo Failed
    m_nextDue = Now
    Do
        If Not m_job.Alive Then
            m_exitReason = m_job.ExitReason
            If Len(m_exitReason) = 0 Then m_exitReason = "stopped"
            Exit Do
        End If

        ToolRegistry.Registry_Worker m_toolId, m_job

        m_nextDue = DateAdd("s", 1, m_nextDue)
        If m_nextDue > Now Then
            Application.Wait m_nextDue
        ElseIf DateDiff("s", m_nextDue, Now) > 1 Then
            m_nextDue = Now
        End If
    Loop
    RequestSelfQuit m_exitReason
    Exit Sub

Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    m_job.ReportFailure failureNumber, failureDescription
    On Error GoTo 0
    RequestSelfQuit "worker_error_" & CStr(failureNumber)
End Sub

Public Sub BeforeClose()
    Dim fileNo As Integer
    If Not m_selfQuit Then Exit Sub
    If Len(m_donePath) = 0 Then Exit Sub
    On Error Resume Next
    fileNo = FreeFile
    Open m_donePath For Output As #fileNo
    Print #fileNo, "job_id=" & m_job.JobId
    Print #fileNo, "reason=" & m_exitReason
    Print #fileNo, "closed_at=" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Close #fileNo
    On Error GoTo 0
End Sub

Private Sub RequestSelfQuit(ByVal reason As String)
    If m_selfQuit Then Exit Sub
    m_selfQuit = True
    m_exitReason = reason
    On Error Resume Next
    If Not m_job Is Nothing Then m_job.ReleaseFE
    ThisWorkbook.Saved = True
    Application.DisplayAlerts = False
    Application.Quit
    On Error GoTo 0
End Sub

