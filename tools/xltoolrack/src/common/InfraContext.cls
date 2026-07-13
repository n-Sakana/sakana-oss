VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InfraContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_toolId As String
Private m_log As Logger
Private m_status As Status
Private m_host As Object
Private m_target As Workbook
Private m_startedJobs As Collection

Public Sub Init(ByVal toolId As String, ByVal logger As Logger, ByVal statusValue As Status, _
                ByVal host As Object, ByVal targetWorkbook As Workbook)
    m_toolId = toolId
    Set m_log = logger
    Set m_status = statusValue
    Set m_host = host
    Set m_target = targetWorkbook
    Set m_startedJobs = New Collection
End Sub

Public Sub LogMessage(ByVal message As String, Optional ByVal level As String = "INFO")
    If m_log Is Nothing Then Exit Sub
    m_log.WriteLine message, level
End Sub

Public Sub Status(ByVal fraction As Double, ByVal message As String)
    If m_status Is Nothing Then Exit Sub
    m_status.Show fraction, message
End Sub

Public Function RunJob(ByVal workerEntry As String, ByVal argsJson As String) As String
    Dim jobId As String
    If m_host Is Nothing Then Err.Raise vbObjectError + 2400, "InfraContext.RunJob", "JobHost is not available."
    jobId = CStr(m_host.StartJob(m_toolId, workerEntry, argsJson, m_target))
    m_startedJobs.Add jobId
    RunJob = jobId
End Function

Public Sub StopJob(ByVal jobId As String)
    If m_host Is Nothing Then Exit Sub
    m_host.StopJob jobId
End Sub

Public Property Get Target() As Workbook
    Set Target = m_target
End Property

Public Property Get ToolId() As String
    ToolId = m_toolId
End Property

Public Sub Fail(ByVal message As String)
    Err.Raise vbObjectError + 2401, "InfraContext.Fail", message
End Sub

Public Function ReadJob(ByVal jobId As String, ByVal address As String) As Variant
    If m_host Is Nothing Then Err.Raise vbObjectError + 2402, "InfraContext.ReadJob", "JobHost is not available."
    ReadJob = m_host.ReadPayload(jobId, address)
End Function

Public Sub BindInput(ByVal jobId As String, ByVal sheet As Worksheet, ByVal address As String)
    If m_host Is Nothing Then Err.Raise vbObjectError + 2403, "InfraContext.BindInput", "JobHost is not available."
    m_host.RegisterInput jobId, sheet, address
End Sub

Public Sub RefreshInput(ByVal jobId As String)
    If m_host Is Nothing Then Err.Raise vbObjectError + 2404, "InfraContext.RefreshInput", "JobHost is not available."
    m_host.SyncRegisteredInput jobId
End Sub

Public Sub RollbackJobs()
    Dim i As Long
    If m_host Is Nothing Then Exit Sub
    If m_startedJobs Is Nothing Then Exit Sub
    On Error Resume Next
    For i = m_startedJobs.Count To 1 Step -1
        m_host.StopJob CStr(m_startedJobs(i))
    Next i
    On Error GoTo 0
End Sub

Public Function StartedJobCount() As Long
    If m_startedJobs Is Nothing Then Exit Function
    StartedJobCount = m_startedJobs.Count
End Function
