VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InfraJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Worker-side job handle. The worker never calls into the FE process: results
' are published to the session aggregate file, input arrives from the channel
' input file, and FE liveness is derived from the stop flag, the shutdown
' flag, and the write lock on the FE lease file. This keeps a busy or
' cell-editing FE completely undisturbed.

Private m_args As String
Private m_feWorkbook As Object
Private m_jobId As String
Private m_token As String
Private m_channelBase As String
Private m_logPath As String
Private m_feGonePath As String
Private m_feLeasePath As String
Private m_version As Long
Private m_inputVersion As Long
Private m_alive As Boolean
Private m_remote As Boolean
Private m_exitReason As String

Public Sub InitLocal(ByVal argsJson As String, ByVal feWorkbook As Object)
    m_args = argsJson
    Set m_feWorkbook = feWorkbook
    m_alive = True
    m_remote = False
End Sub

Public Sub InitRemote(ByVal jobId As String, ByVal argsJson As String, ByVal token As String, _
                      ByVal channelBase As String, ByVal logPath As String, _
                      ByVal feGonePath As String, ByVal feLeasePath As String)
    m_jobId = jobId
    m_args = argsJson
    m_token = token
    m_channelBase = channelBase
    m_logPath = logPath
    m_feGonePath = feGonePath
    m_feLeasePath = feLeasePath
    m_version = 0
    m_inputVersion = 0
    m_alive = True
    m_remote = True
    m_exitReason = vbNullString
End Sub

Public Property Get Args() As String
    Args = m_args
End Property

Public Property Get FEWorkbook() As Object
    Set FEWorkbook = m_feWorkbook
End Property

Public Property Get JobId() As String
    JobId = m_jobId
End Property

Public Property Get Alive() As Boolean
    If m_remote Then m_alive = CheckChannelState()
    Alive = m_alive
End Property

Public Property Get ExitReason() As String
    ExitReason = m_exitReason
End Property

Public Property Get InputVersion() As Long
    InputVersion = m_inputVersion
End Property

Public Sub SetAliveForTest(ByVal value As Boolean)
    m_alive = value
End Sub

Public Function Push(ByVal sheetName As String, ByVal address As String, ByVal values As Variant) As Boolean
    If Not m_remote Then Err.Raise vbObjectError + 2410, "InfraJob.Push", "Job is not bound to a remote channel."
    If Not PublishAggregate(address, values) Then
        LogMessage "payload publish skipped (file busy)"
        Exit Function
    End If
    m_version = m_version + 1
    Push = True
End Function

Private Function PublishAggregate(ByVal address As String, ByVal values As Variant) As Boolean
    Dim startedAt As Double
    Dim tick As Double
    startedAt = Timer
    Do
        If ChannelFile.WriteAggregateRecord(m_channelBase, m_jobId, m_version + 1, address, values) Then
            PublishAggregate = True
            Exit Function
        End If
        DoEvents
        tick = Timer
        If tick < startedAt Then tick = tick + 86400#
    Loop While tick - startedAt < 0.25
End Function

' Returns the most recent input rectangle the FE published, or Empty when no
' input has arrived yet. The address argument is kept for the tool contract;
' the whole registered range is always returned.
Public Function ReadInput(ByVal address As String) As Variant
    Dim version As Long
    Dim fileAddress As String
    Dim values As Variant
    If Not m_remote Then Err.Raise vbObjectError + 2411, "InfraJob.ReadInput", "Job is not bound to a remote channel."
    If ChannelFile.ReadRect(ChannelFile.InputPath(m_channelBase), version, fileAddress, values) Then
        m_inputVersion = version
        ReadInput = values
    End If
End Function

Public Sub LogMessage(ByVal message As String)
    Dim fileNo As Integer
    If Len(m_logPath) = 0 Then
        Debug.Print "-- InfraJob | " & message
        Exit Sub
    End If
    On Error Resume Next
    fileNo = FreeFile
    Open m_logPath For Append As #fileNo
    Print #fileNo, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " BE " & message
    Close #fileNo
    On Error GoTo 0
End Sub

Public Sub ReportFailure(ByVal number As Long, ByVal description As String)
    On Error Resume Next
    ChannelFile.WriteError ChannelFile.ErrorPath(m_channelBase), number, description
    LogMessage "FAIL Err " & CStr(number) & ": " & description
    On Error GoTo 0
End Sub

Public Sub ReleaseFE()
    Set m_feWorkbook = Nothing
    m_alive = False
End Sub

Private Function CheckChannelState() As Boolean
    If ChannelFile.FlagExists(ChannelFile.StopPath(m_channelBase)) Then
        m_exitReason = "stop_requested"
        Exit Function
    End If
    If ChannelFile.FlagExists(m_feGonePath) Then
        m_exitReason = "fe_unavailable"
        Exit Function
    End If
    If Not FeLeaseAlive() Then
        m_exitReason = "fe_unavailable"
        Exit Function
    End If
    CheckChannelState = True
End Function

' The FE holds the lease file open with a deny-all lock. A locked file means
' the FE process is alive (even mid cell-edit); an unlocked or missing file
' means it is gone - a crash releases the lock the moment the process dies.
Private Function FeLeaseAlive() As Boolean
    Dim fileNo As Integer
    Dim errorNumber As Long
    If Len(m_feLeasePath) = 0 Then
        FeLeaseAlive = True
        Exit Function
    End If
    If Not ChannelFile.FlagExists(m_feLeasePath) Then Exit Function
    On Error Resume Next
    Err.Clear
    fileNo = FreeFile
    Open m_feLeasePath For Binary Access Write Lock Read Write As #fileNo
    errorNumber = Err.Number
    If errorNumber = 0 Then Close #fileNo
    Err.Clear
    On Error GoTo 0
    FeLeaseAlive = (errorNumber <> 0)
End Function
