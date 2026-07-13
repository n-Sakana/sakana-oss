VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JobHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const JOBS_SHEET As String = "_xtr_jobs"
Private Const WORKER_FILE As String = "xltoolrack-worker.xlsm"
Private Const DEFAULT_GLOBAL_LIMIT As Long = 10
Private Const ERROR_POLL_DELAY_SECONDS As Double = 2#
Private Const STALE_PAYLOAD_SECONDS As Double = 5#

Private m_hostBook As Workbook
Private m_jobsSheet As Worksheet
Private m_jobs As Object
Private m_nextRow As Long
Private m_sequence As Long
Private m_globalLimit As Long
Private m_lastTarget As Workbook
Private m_failStage As String
Private m_sweeping As Boolean
Private m_lastSweepTick As Double
Private m_pumping As Boolean
Private m_lastPumpTick As Double

Private Sub Class_Initialize()
    Set m_jobs = CreateObject("Scripting.Dictionary")
    m_jobs.CompareMode = 1
    m_nextRow = 2
    m_globalLimit = DEFAULT_GLOBAL_LIMIT
    m_lastSweepTick = -10
End Sub

Public Sub Init(ByVal hostBook As Workbook)
    Set m_hostBook = hostBook
    Set m_jobsSheet = EnsureSheet(JOBS_SHEET)
    If Len(CStr(m_jobsSheet.Range("A1").Value2)) = 0 Then
        m_jobsSheet.Range("A1:M1").Value = Array("channel", "job_id", "token", "version", "stop", "heartbeat", "state", "error_no", "error_text", "tool_id", "payload_sheet", "input_sheet", "done_path")
    End If
    m_jobsSheet.Visible = xlSheetVeryHidden
    m_hostBook.Saved = True
End Sub

Public Function StartJob(ByVal toolId As String, ByVal workerEntry As String, _
                         ByVal argsJson As String, ByVal target As Workbook) As String
    Dim state As JobState
    Dim beApp As Object
    Dim beBook As Object
    Dim previousSecurity As Long
    Dim securityChanged As Boolean
    Dim workerSource As String
    Dim logger As Logger
    Dim originalNumber As Long
    Dim originalDescription As String
    Dim originalSource As String

    On Error GoTo Failed
    SweepFinished
    If ActiveCount() >= m_globalLimit Then Err.Raise vbObjectError + 2700, "JobHost.StartJob", "global limit reached"
    If CountForTool(toolId) >= ToolRegistry.Registry_MaxJobs(toolId) Then Err.Raise vbObjectError + 2701, "JobHost.StartJob", "tool limit reached"
    If target Is Nothing Then Err.Raise vbObjectError + 2702, "JobHost.StartJob", "target workbook is required"

    Set state = New JobState
    m_sequence = m_sequence + 1
    state.ToolId = toolId
    state.WorkerEntry = workerEntry
    state.ArgsJson = argsJson
    state.JobId = NewJobId(toolId)
    state.ChannelId = "j" & Format$(m_sequence, "000000")
    state.SessionToken = state.JobId & "_" & Format$(CLng(Timer * 1000!) Mod 100000000, "00000000")
    state.RowNumber = m_nextRow
    m_nextRow = m_nextRow + 1
    state.PayloadSheetName = "_xtr_o_" & state.ChannelId
    state.ChannelBase = PathFs.ChannelBasePath(state.ChannelId)
    state.DonePath = Environ$("TEMP") & Application.PathSeparator & "xltoolrack_done_" & PathFs.SessionId() & "_" & state.ChannelId & ".flag"
    state.WorkerCopy = Environ$("TEMP") & Application.PathSeparator & "xltoolrack_worker_" & PathFs.SessionId() & "_" & state.ChannelId & ".xlsm"
    state.LastVersion = 0
    state.LastPayloadTick = Timer
    state.StateName = "starting"
    Set state.Target = target

    Set logger = New Logger
    logger.Begin toolId, state.JobId
    logger.WriteLine "job reserved"
    state.LogPath = logger.Path

    m_jobs.Add state.JobId, vbNullString
    Set m_jobs.Item(state.JobId) = state
    CreateJobChannels state

    workerSource = m_hostBook.Path & Application.PathSeparator & WORKER_FILE
    If Len(Dir$(workerSource)) = 0 Then Err.Raise vbObjectError + 2703, "JobHost.StartJob", "worker workbook not found"
    On Error Resume Next
    If Len(Dir$(state.WorkerCopy)) > 0 Then Kill state.WorkerCopy
    If Len(Dir$(state.DonePath)) > 0 Then Kill state.DonePath
    If Len(Dir$(PathFs.FeShutdownFlagPath())) > 0 Then Kill PathFs.FeShutdownFlagPath()
    On Error GoTo Failed
    DeleteChannelFiles state
    HostMain.EnsureLease
    FileCopy workerSource, state.WorkerCopy

    Set beApp = CreateObject("Excel.Application")
    If m_failStage = "after_create" Then Err.Raise vbObjectError + 2707, "JobHost.StartJob", "injected start failure after_create"
    beApp.Visible = False
    beApp.DisplayAlerts = False
    beApp.EnableEvents = True
    previousSecurity = beApp.AutomationSecurity
    beApp.AutomationSecurity = 1
    securityChanged = True
    Set beBook = beApp.Workbooks.Open(state.WorkerCopy, 0, True)
    beApp.AutomationSecurity = previousSecurity
    securityChanged = False
    Set state.BeApp = beApp
    Set state.BeBook = beBook

    beApp.Run QualifiedMacro(CStr(beBook.Name), "WorkerBridge.WorkerBootstrap"), _
        state.JobId, state.ToolId, state.ArgsJson, state.DonePath, _
        state.SessionToken, state.ChannelBase, state.LogPath, _
        PathFs.FeShutdownFlagPath(), PathFs.FeLeasePath()

    state.StateName = "running"
    WriteJobCell state.RowNumber, "G", "running"
    m_hostBook.Saved = True
    JobPump.Pump_Start
    StartJob = state.JobId
    Exit Function

Failed:
    originalNumber = Err.Number
    originalDescription = Err.Description
    originalSource = Err.Source
    On Error Resume Next
    If securityChanged Then beApp.AutomationSecurity = previousSecurity
    If Not state Is Nothing Then
        WriteJobCell state.RowNumber, "E", 1
        state.StateName = "failed"
    End If
    If Not beApp Is Nothing Then beApp.Quit
    Set beBook = Nothing
    Set beApp = Nothing
    If Not state Is Nothing Then RollbackState state
    On Error GoTo 0
    Err.Raise originalNumber, originalSource, originalDescription
End Function

Public Sub StopJob(ByVal jobId As String)
    Dim state As JobState
    Set state = GetState(jobId)
    If state Is Nothing Then Exit Sub
    If Not state.IsActive Then Exit Sub
    state.StateName = "stopping"
    On Error Resume Next
    ChannelFile.WriteFlag ChannelFile.StopPath(state.ChannelBase), "stop"
    WriteJobCell state.RowNumber, "E", 1
    WriteJobCell state.RowNumber, "G", "stopping"
    Set state.BeBook = Nothing
    Set state.BeApp = Nothing
    m_hostBook.Saved = True
    On Error GoTo 0
End Sub

Public Sub StopAll()
    Dim keys As Variant
    Dim i As Long
    If m_jobs.Count = 0 Then Exit Sub
    keys = m_jobs.Keys
    For i = LBound(keys) To UBound(keys)
        StopJob CStr(keys(i))
    Next i
End Sub

Public Function RunningCount() As Long
    MaybePump
    SweepFinished
    RunningCount = ActiveCount()
End Function

Public Function CountForTool(ByVal toolId As String) As Long
    Dim key As Variant
    Dim state As JobState
    For Each key In m_jobs.Keys
        Set state = m_jobs(key)
        If state.IsActive Then
            If StrComp(state.ToolId, toolId, vbTextCompare) = 0 Then CountForTool = CountForTool + 1
        End If
    Next key
End Function

Public Sub SetGlobalLimitForTest(ByVal value As Long)
    If value < 1 Then Err.Raise vbObjectError + 2704, "JobHost.SetGlobalLimitForTest", "limit must be positive"
    m_globalLimit = value
End Sub

Public Sub SetFailStageForTest(ByVal value As String)
    m_failStage = value
End Sub

Public Sub HandleSheetChange(ByVal changedSheet As Object, ByVal targetRange As Range)
    Dim hit As Range
    Dim key As Variant
    Dim state As JobState

    On Error GoTo CleanExit
    MaybeSweep
    JobPump.Pump_EnsureArmed
    If changedSheet Is m_jobsSheet Then Exit Sub

    For Each key In m_jobs.Keys
        Set state = m_jobs(key)
        If Not state.InputTargetSheet Is Nothing Then
            If changedSheet Is state.InputTargetSheet Then
                Set hit = Application.Intersect(targetRange, state.InputTargetSheet.Range(state.InputTargetAddress))
                If Not hit Is Nothing Then SyncInput state
            End If
        End If
    Next key
CleanExit:
    If Err.Number <> 0 Then Debug.Print "!! JobHost.HandleSheetChange | Err " & CStr(Err.Number)
End Sub

' One pump cycle, driven by the JobPump OnTime tick: consume worker error
' files, apply new payload versions to the payload sheets, dispatch OnResult,
' then sweep. Returns True while any job is still active so the pump knows
' whether to reschedule itself.
Public Function PumpOnce(Optional ByRef profileFreshRecords As Long = 0, _
                         Optional ByRef profileAggregateReadMs As Double = 0#, _
                         Optional ByRef profileAggregateExistsMs As Double = 0#, _
                         Optional ByRef profileAggregateOpenMs As Double = 0#, _
                         Optional ByRef profileAggregateLineInputMs As Double = 0#, _
                         Optional ByRef profileAggregateDecodeMs As Double = 0#, _
                         Optional ByRef profileAggregateCloseMs As Double = 0#, _
                         Optional ByRef profileFlushTools As Long = 0, _
                         Optional ByRef profileDispatchMs As Double = 0#, _
                         Optional ByRef profileErrorPollMs As Double = 0#, _
                         Optional ByRef profileCleanupMs As Double = 0#, _
                         Optional ByVal profileRequested As Boolean = False) As Boolean
    Dim keys As Variant
    Dim i As Long
    Dim state As JobState
    Dim records As Collection
    Dim record As ChannelRecord
    Dim errorNumber As Long
    Dim errorText As String
    Dim dispatchGuard As AppStateSnapshot
    Dim flushStates As Object
    Dim flushKeys As Variant
    Dim flushToolCount As Long
    Dim freshRecordCount As Long
    Dim aggregateStarted As Double
    Dim aggregateReadMs As Double
    Dim aggregateExistsMs As Double
    Dim aggregateOpenMs As Double
    Dim aggregateLineInputMs As Double
    Dim aggregateDecodeMs As Double
    Dim aggregateCloseMs As Double
    Dim dispatchStarted As Double
    Dim dispatchMs As Double
    Dim errorPollStarted As Double
    Dim errorPollMs As Double
    Dim cleanupStarted As Double
    Dim cleanupMs As Double

    Dim previousEvents As Boolean
    If m_pumping Then
        PumpOnce = (ActiveCount() > 0)
        Exit Function
    End If
    m_pumping = True
    previousEvents = Application.EnableEvents
    On Error GoTo CleanExit
    ' Events stay off for the whole cycle so the pump's own sheet writes do not
    ' re-enter HandleSheetChange. One aggregate read supplies every job result.
    Application.EnableEvents = False
    If m_jobs.Count > 0 Then
        keys = m_jobs.Keys
        Set state = m_jobs(keys(LBound(keys)))
        If profileRequested Then aggregateStarted = Timer
        Set records = ChannelFile.ReadAggregate(state.ChannelBase, aggregateExistsMs, _
                                                aggregateOpenMs, aggregateLineInputMs, _
                                                aggregateDecodeMs, aggregateCloseMs, _
                                                profileRequested)
        If profileRequested Then aggregateReadMs = ElapsedMilliseconds(aggregateStarted, Timer)
        If profileRequested Then dispatchStarted = Timer
        For Each record In records
            Set state = GetState(record.JobId)
            If Not state Is Nothing Then
                If state.StateName = "running" Or state.StateName = "starting" Then
                    If record.Version > state.LastVersion Then
                        If profileRequested Then freshRecordCount = freshRecordCount + 1
                        state.LastVersion = record.Version
                        state.LastPayloadTick = Timer
                        state.CachedPayloadAddress = record.Address
                        state.CachedPayloadValues = record.Values
                        state.CachedPayloadDirty = True
                        If ToolRegistry.Registry_HasResult(state.ToolId) Then
                            If dispatchGuard Is Nothing Then Set dispatchGuard = New AppStateSnapshot
                            Infra_DispatchResult Me, state, record.Version, dispatchGuard
                            If (state.StateName = "running" Or state.StateName = "starting") And _
                               ToolRegistry.Registry_HasFlush(state.ToolId) Then
                                If flushStates Is Nothing Then
                                    Set flushStates = CreateObject("Scripting.Dictionary")
                                    flushStates.CompareMode = 1
                                End If
                                If Not flushStates.Exists(state.ToolId) Then
                                    flushStates.Add state.ToolId, state
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next record
        If Not flushStates Is Nothing Then
            flushKeys = flushStates.Keys
            For i = LBound(flushKeys) To UBound(flushKeys)
                Set state = flushStates(flushKeys(i))
                If state.StateName = "running" Or state.StateName = "starting" Then
                    flushToolCount = flushToolCount + 1
                    Infra_DispatchFlush Me, state, dispatchGuard
                End If
            Next i
        End If
        If profileRequested Then dispatchMs = ElapsedMilliseconds(dispatchStarted, Timer)

        ' Error files are exceptional. Healthy workers refresh LastPayloadTick,
        ' so their missing error paths are not touched on the steady-state path.
        If profileRequested Then errorPollStarted = Timer
        For i = LBound(keys) To UBound(keys)
            If m_jobs.Exists(keys(i)) Then
                Set state = m_jobs(keys(i))
                If state.StateName = "running" Or state.StateName = "starting" Then
                    If PayloadAgeSeconds(state) >= ERROR_POLL_DELAY_SECONDS Then
                    If ChannelFile.ReadError(ChannelFile.ErrorPath(state.ChannelBase), errorNumber, errorText) Then
                        ChannelFile.DeleteQuiet ChannelFile.ErrorPath(state.ChannelBase)
                        NoteWorkerFailure state, errorNumber, errorText
                    End If
                    End If
                End If
            End If
        Next i
        If profileRequested Then errorPollMs = ElapsedMilliseconds(errorPollStarted, Timer)
    End If
CleanExit:
    If Err.Number <> 0 Then Debug.Print "!! JobHost.PumpOnce | Err " & CStr(Err.Number)
    On Error Resume Next
    If profileRequested Then cleanupStarted = Timer
    If Not dispatchGuard Is Nothing Then dispatchGuard.Restore
    MaybeSweepFromPump
    PumpOnce = (ActiveCount() > 0)
    Application.EnableEvents = previousEvents
    If profileRequested Then
        cleanupMs = ElapsedMilliseconds(cleanupStarted, Timer)
        profileFreshRecords = freshRecordCount
        profileAggregateReadMs = aggregateReadMs
        profileAggregateExistsMs = aggregateExistsMs
        profileAggregateOpenMs = aggregateOpenMs
        profileAggregateLineInputMs = aggregateLineInputMs
        profileAggregateDecodeMs = aggregateDecodeMs
        profileAggregateCloseMs = aggregateCloseMs
        profileFlushTools = flushToolCount
        profileDispatchMs = dispatchMs
        profileErrorPollMs = errorPollMs
        profileCleanupMs = cleanupMs
    End If
    m_pumping = False
    On Error GoTo 0
End Function

Private Sub MaybeSweepFromPump()
    Dim keys As Variant
    Dim i As Long
    Dim state As JobState
    If m_sweeping Then Exit Sub
    If m_jobs.Count = 0 Then Exit Sub
    m_sweeping = True
    On Error GoTo CleanExit
    keys = m_jobs.Keys
    For i = LBound(keys) To UBound(keys)
        If m_jobs.Exists(keys(i)) Then
            Set state = m_jobs(keys(i))
            If state.StateName = "stopping" Or state.StateName = "failed" Then
                SweepStateIfDone state
            ElseIf state.StateName = "running" Or state.StateName = "starting" Then
                If PayloadIsStale(state) Then SweepStateIfDone state
            End If
        End If
    Next i
CleanExit:
    m_sweeping = False
End Sub

Private Function PayloadIsStale(ByVal state As JobState) As Boolean
    PayloadIsStale = (PayloadAgeSeconds(state) >= STALE_PAYLOAD_SECONDS)
End Function

Private Function PayloadAgeSeconds(ByVal state As JobState) As Double
    Dim tick As Double
    tick = Timer
    If tick < state.LastPayloadTick Then tick = tick + 86400#
    PayloadAgeSeconds = tick - state.LastPayloadTick
End Function

' Opportunistic pump for automation-heavy sessions: a stream of inbound COM
' calls (tests, external drivers) can starve OnTime, so the query entry
' points deliver results themselves at most every 0.4 seconds.
Public Sub MaybePump()
    Dim tick As Double
    tick = Timer
    If tick < m_lastPumpTick Then m_lastPumpTick = tick - 1
    If tick - m_lastPumpTick < 0.4 Then Exit Sub
    m_lastPumpTick = tick
    Call PumpOnce
End Sub

Public Function HasActiveJobs() As Boolean
    HasActiveJobs = (ActiveCount() > 0)
End Function

Private Sub NoteWorkerFailure(ByVal state As JobState, ByVal number As Long, ByVal description As String)
    Dim errors As ErrorHandler
    state.StateName = "failed"
    state.HoldUntil = DateAdd("s", 4, Now)
    WriteJobCell state.RowNumber, "G", "failed"
    WriteJobCell state.RowNumber, "H", number
    WriteJobCell state.RowNumber, "I", description
    WriteJobCell state.RowNumber, "E", 1
    Set errors = New ErrorHandler
    errors.Report state.ToolId, number, description, True
End Sub

Public Sub HandleWorkbookBeforeClose(ByVal closingBook As Workbook)
    Dim key As Variant
    Dim state As JobState
    For Each key In m_jobs.Keys
        Set state = m_jobs(key)
        If state.Target Is closingBook Then StopJob state.JobId
    Next key
End Sub

Public Sub NoteWorkbookActivated(ByVal candidate As Workbook)
    On Error Resume Next
    MaybeSweep
    JobPump.Pump_EnsureArmed
    If Not candidate Is m_hostBook Then
        If Not candidate.IsAddin Then Set m_lastTarget = candidate
    End If
    On Error GoTo 0
End Sub

Public Function ReadPayload(ByVal jobId As String, ByVal address As String) As Variant
    Dim state As JobState
    Set state = GetState(jobId)
    If state Is Nothing Then Err.Raise vbObjectError + 2705, "JobHost.ReadPayload", "unknown job"
    If Len(state.CachedPayloadAddress) > 0 Then
        If StrComp(address, state.CachedPayloadAddress, vbTextCompare) = 0 Then
            ReadPayload = state.CachedPayloadValues
            Exit Function
        End If
    End If
    If state.CachedPayloadDirty Then
        m_hostBook.Worksheets(state.PayloadSheetName).Range(state.CachedPayloadAddress).Value2 = state.CachedPayloadValues
        state.CachedPayloadDirty = False
        m_hostBook.Saved = True
    End If
    ReadPayload = m_hostBook.Worksheets(state.PayloadSheetName).Range(address).Value2
End Function

Public Sub RegisterInput(ByVal jobId As String, ByVal targetSheet As Worksheet, ByVal address As String)
    Dim state As JobState
    Set state = GetState(jobId)
    If state Is Nothing Then Err.Raise vbObjectError + 2706, "JobHost.RegisterInput", "unknown job"
    Set state.InputTargetSheet = targetSheet
    state.InputTargetAddress = address
    SyncInput state
End Sub

Public Sub SyncRegisteredInput(ByVal jobId As String)
    Dim state As JobState
    Set state = GetState(jobId)
    If state Is Nothing Then Err.Raise vbObjectError + 2708, "JobHost.SyncRegisteredInput", "unknown job"
    If state.InputTargetSheet Is Nothing Then Exit Sub
    SyncInput state
End Sub

Public Function VersionFor(ByVal jobId As String) As Long
    Dim state As JobState
    MaybePump
    Set state = GetState(jobId)
    If state Is Nothing Then Exit Function
    VersionFor = state.LastVersion
End Function

Public Function StateFor(ByVal jobId As String) As String
    Dim state As JobState
    MaybePump
    Set state = GetState(jobId)
    If state Is Nothing Then Exit Function
    StateFor = state.StateName
End Function

Public Function DonePathFor(ByVal jobId As String) As String
    Dim state As JobState
    Set state = GetState(jobId)
    If state Is Nothing Then Exit Function
    DonePathFor = state.DonePath
End Function

Public Function LogPathFor(ByVal jobId As String) As String
    Dim state As JobState
    Set state = GetState(jobId)
    If state Is Nothing Then Exit Function
    LogPathFor = state.LogPath
End Function

Public Function WorkerCopyFor(ByVal jobId As String) As String
    Dim state As JobState
    Set state = GetState(jobId)
    If state Is Nothing Then Exit Function
    WorkerCopyFor = state.WorkerCopy
End Function

Public Function ErrorFor(ByVal jobId As String) As String
    Dim state As JobState
    Set state = GetState(jobId)
    If state Is Nothing Then Exit Function
    ErrorFor = CStr(m_jobsSheet.Range("H" & CStr(state.RowNumber)).Value2) & ":" & _
        CStr(m_jobsSheet.Range("I" & CStr(state.RowNumber)).Value2)
End Function

Public Sub FailJob(ByVal jobId As String, ByVal number As Long, ByVal description As String)
    Dim state As JobState
    Set state = GetState(jobId)
    If state Is Nothing Then Exit Sub
    state.StateName = "failed"
    state.HoldUntil = DateAdd("s", 4, Now)
    WriteJobCell state.RowNumber, "G", "failed"
    WriteJobCell state.RowNumber, "H", number
    WriteJobCell state.RowNumber, "I", description
    WriteJobCell state.RowNumber, "E", 1
End Sub

Public Sub SweepFinished()
    Dim keys As Variant
    Dim i As Long
    Dim state As JobState
    If m_sweeping Then Exit Sub
    If m_jobs.Count = 0 Then Exit Sub
    m_sweeping = True
    On Error GoTo CleanExit
    keys = m_jobs.Keys
    For i = LBound(keys) To UBound(keys)
        If m_jobs.Exists(keys(i)) Then
            Set state = m_jobs(keys(i))
            SweepStateIfDone state
        End If
    Next i
CleanExit:
    m_sweeping = False
End Sub

Private Sub SweepStateIfDone(ByVal state As JobState)
    Dim wasUnexpected As Boolean
    If Len(state.DonePath) = 0 Then Exit Sub
    If Len(Dir$(state.DonePath)) = 0 Then Exit Sub
    ' A done flag on a running job means the worker exited on its own. Healthy
    ' running jobs reach this check only after their payload has gone stale;
    ' terminal jobs are checked every pump so stop cleanup remains prompt.
    wasUnexpected = (state.StateName = "running" Or state.StateName = "starting")
    If wasUnexpected Then NoteUnexpectedExit state
    If Now >= state.HoldUntil Then FinalizeState state
End Sub

' Event-safe subset of the sweep: only detect and mark dead jobs. Deleting
' sheets or files inside SheetChange/Activate dispatch can destabilize Excel,
' so FinalizeState stays on the StartJob / RunningCount entry points.
Private Sub MaybeSweep()
    Dim tick As Double
    tick = Timer
    If tick < m_lastSweepTick Then m_lastSweepTick = tick - 2
    If tick - m_lastSweepTick < 2 Then Exit Sub
    m_lastSweepTick = tick
    MarkDeadJobs
End Sub

Private Sub MarkDeadJobs()
    Dim keys As Variant
    Dim i As Long
    Dim state As JobState
    If m_sweeping Then Exit Sub
    If m_jobs.Count = 0 Then Exit Sub
    m_sweeping = True
    On Error GoTo CleanExit
    keys = m_jobs.Keys
    For i = LBound(keys) To UBound(keys)
        Set state = m_jobs(keys(i))
        If state.StateName = "running" Or state.StateName = "starting" Then
            If Len(state.DonePath) > 0 Then
                If Len(Dir$(state.DonePath)) > 0 Then NoteUnexpectedExit state
            End If
        End If
    Next i
CleanExit:
    m_sweeping = False
End Sub

Private Sub NoteUnexpectedExit(ByVal state As JobState)
    Dim logger As Logger
    Dim errors As ErrorHandler
    On Error Resume Next
    state.StateName = "stopped"
    state.HoldUntil = DateAdd("s", 4, Now)
    WriteJobCell state.RowNumber, "G", "stopped"
    Set logger = New Logger
    logger.Attach state.LogPath
    logger.WriteLine "worker exited without a stop request; job finalized", "WARN"
    Set errors = New ErrorHandler
    errors.Report state.ToolId, 0, "worker for job " & state.JobId & " exited unexpectedly", True
    On Error GoTo 0
End Sub

Private Sub CreateJobChannels(ByVal state As JobState)
    Dim values(1 To 1, 1 To 13) As Variant
    Call EnsureSheet(state.PayloadSheetName)
    m_hostBook.Worksheets(state.PayloadSheetName).Cells.ClearContents
    values(1, 1) = state.ChannelId
    values(1, 2) = state.JobId
    values(1, 3) = state.SessionToken
    values(1, 4) = 0
    values(1, 5) = 0
    values(1, 6) = CDbl(Now)
    values(1, 7) = "starting"
    values(1, 8) = 0
    values(1, 9) = vbNullString
    values(1, 10) = state.ToolId
    values(1, 11) = state.PayloadSheetName
    values(1, 12) = state.ChannelBase
    values(1, 13) = state.DonePath
    m_jobsSheet.Range("A" & CStr(state.RowNumber) & ":M" & CStr(state.RowNumber)).Value = values
    m_hostBook.Saved = True
End Sub

Private Sub SyncInput(ByVal state As JobState)
    Dim sourceRange As Range
    Set sourceRange = state.InputTargetSheet.Range(state.InputTargetAddress)
    state.InputVersion = state.InputVersion + 1
    Call ChannelFile.WriteRect(ChannelFile.InputPath(state.ChannelBase), state.InputVersion, "A1", sourceRange.Value2)
End Sub

Private Sub DeleteChannelFiles(ByVal state As JobState)
    ChannelFile.DeleteAggregateRecord state.ChannelBase, state.JobId
    ChannelFile.DeleteQuiet ChannelFile.PayloadPath(state.ChannelBase)
    ChannelFile.DeleteQuiet ChannelFile.PayloadPath(state.ChannelBase) & ".tmp"
    ChannelFile.DeleteQuiet ChannelFile.InputPath(state.ChannelBase)
    ChannelFile.DeleteQuiet ChannelFile.InputPath(state.ChannelBase) & ".tmp"
    ChannelFile.DeleteQuiet ChannelFile.StopPath(state.ChannelBase)
    ChannelFile.DeleteQuiet ChannelFile.ErrorPath(state.ChannelBase)
    ChannelFile.DeleteQuiet ChannelFile.ErrorPath(state.ChannelBase) & ".tmp"
End Sub

Private Function ElapsedMilliseconds(ByVal startedAt As Double, ByVal endedAt As Double) As Double
    If endedAt < startedAt Then endedAt = endedAt + 86400#
    ElapsedMilliseconds = (endedAt - startedAt) * 1000#
End Function

Private Function ActiveCount() As Long
    Dim key As Variant
    Dim state As JobState
    For Each key In m_jobs.Keys
        Set state = m_jobs(key)
        If state.IsActive Then ActiveCount = ActiveCount + 1
    Next key
End Function

Private Function GetState(ByVal jobId As String) As JobState
    If Len(jobId) = 0 Then Exit Function
    If Not m_jobs.Exists(jobId) Then Exit Function
    Set GetState = m_jobs(jobId)
End Function

Private Sub RollbackState(ByVal state As JobState)
    On Error Resume Next
    Set state.BeBook = Nothing
    Set state.BeApp = Nothing
    If m_jobs.Exists(state.JobId) Then m_jobs.Remove state.JobId
    DeleteSheet state.PayloadSheetName
    DeleteChannelFiles state
    Call TryDeleteFile(state.WorkerCopy, True)
    m_jobsSheet.Range("A" & CStr(state.RowNumber) & ":M" & CStr(state.RowNumber)).ClearContents
    m_hostBook.Saved = True
    On Error GoTo 0
End Sub

Private Sub FinalizeState(ByVal state As JobState)
    On Error Resume Next
    Set state.BeBook = Nothing
    Set state.BeApp = Nothing
    state.StateName = "completed"
    If Not TryDeleteFile(state.WorkerCopy, True) Then Exit Sub
    If m_jobs.Exists(state.JobId) Then m_jobs.Remove state.JobId
    DeleteSheet state.PayloadSheetName
    DeleteChannelFiles state
    Call TryDeleteFile(state.DonePath, False)
    m_hostBook.Saved = True
    On Error GoTo 0
End Sub

Private Function TryDeleteFile(ByVal path As String, ByVal waitOnce As Boolean) As Boolean
    If Len(path) = 0 Then
        TryDeleteFile = True
        Exit Function
    End If
    On Error Resume Next
    If Len(Dir$(path)) > 0 Then Kill path
    On Error GoTo 0
    If Len(Dir$(path)) = 0 Then
        TryDeleteFile = True
        Exit Function
    End If
    If waitOnce Then
        Application.Wait Now + TimeSerial(0, 0, 1)
        On Error Resume Next
        Kill path
        On Error GoTo 0
    End If
    TryDeleteFile = (Len(Dir$(path)) = 0)
End Function

Private Function EnsureSheet(ByVal sheetName As String) As Worksheet
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = m_hostBook.Worksheets(sheetName)
    On Error GoTo 0
    If sheet Is Nothing Then
        Set sheet = m_hostBook.Worksheets.Add(After:=m_hostBook.Worksheets(m_hostBook.Worksheets.Count))
        sheet.Name = sheetName
    End If
    sheet.Visible = xlSheetVeryHidden
    Set EnsureSheet = sheet
End Function

Private Sub DeleteSheet(ByVal sheetName As String)
    Dim oldAlerts As Boolean
    oldAlerts = Application.DisplayAlerts
    On Error Resume Next
    Application.DisplayAlerts = False
    m_hostBook.Worksheets(sheetName).Delete
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
End Sub

Private Sub WriteJobCell(ByVal rowNumber As Long, ByVal columnName As String, ByVal value As Variant)
    m_jobsSheet.Range(columnName & CStr(rowNumber)).Value2 = value
End Sub

Private Function NewJobId(ByVal toolId As String) As String
    NewJobId = SafeId(toolId) & "_" & Format$(Now, "yyyymmddhhnnss") & "_" & _
        SafeId(CStr(Application.hWnd)) & "_" & Format$(m_sequence, "0000")
End Function

Private Function SafeId(ByVal value As String) As String
    Dim i As Long
    Dim ch As String
    Dim result As String
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or _
           (ch >= "0" And ch <= "9") Or ch = "_" Then
            result = result & ch
        Else
            result = result & "_"
        End If
    Next i
    If Len(result) = 0 Then result = "job"
    SafeId = result
End Function

Private Function QualifiedMacro(ByVal workbookName As String, ByVal procedureName As String) As String
    QualifiedMacro = "'" & Replace(workbookName, "'", "''") & "'!" & procedureName
End Function
