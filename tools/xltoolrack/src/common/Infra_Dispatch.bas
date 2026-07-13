Attribute VB_Name = "Infra_Dispatch"
Option Explicit

Private m_lastOutcome As String
Private m_lastLogPath As String

Public Sub Infra_RunTool(ByVal control As IRibbonControl)
    Infra_Run CStr(control.id)
End Sub

Public Sub Infra_Run(ByVal toolId As String)
    Dim guard As AppGuard
    Dim logger As Logger
    Dim stat As Status
    Dim ctx As InfraContext
    Dim errors As ErrorHandler
    Dim target As Workbook
    Dim failureNumber As Long
    Dim failureDescription As String

    m_lastOutcome = "STARTING"
    m_lastLogPath = vbNullString
    On Error GoTo Failed

    Set target = ResolveTargetWorkbook()
    Set guard = New AppGuard
    Set logger = New Logger
    Set stat = New Status
    Set errors = New ErrorHandler
    logger.Begin toolId
    m_lastLogPath = logger.Path
    stat.Begin toolId
    Set ctx = New InfraContext
    ctx.Init toolId, logger, stat, HostServices.Jobs(), target

    ToolRegistry.Registry_Run toolId, ctx

    logger.Done
    stat.Clear
    guard.Restore
    m_lastOutcome = "OK"
    Exit Sub

Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    If Not ctx Is Nothing Then ctx.RollbackJobs
    If Not stat Is Nothing Then stat.Clear
    If Not logger Is Nothing Then
        m_lastLogPath = logger.Path
        logger.Fail "Err " & CStr(failureNumber) & ": " & failureDescription
    End If
    If errors Is Nothing Then Set errors = New ErrorHandler
    errors.Report toolId, failureNumber, failureDescription, False
    If Not guard Is Nothing Then guard.Restore
    m_lastOutcome = "ERROR:" & CStr(failureNumber)
    On Error GoTo 0
End Sub

Public Function Infra_LastOutcome() As String
    Infra_LastOutcome = m_lastOutcome
End Function

Public Function Infra_LastLogPath() As String
    Infra_LastLogPath = m_lastLogPath
End Function

Public Sub Infra_DispatchResult(ByVal host As JobHost, ByVal state As JobState, ByVal versionValue As Long, _
                                ByVal batchGuard As AppStateSnapshot)
    Dim logger As Logger
    Dim stat As Status
    Dim ctx As InfraContext
    Dim errors As ErrorHandler
    Dim failureNumber As Long
    Dim failureDescription As String

    On Error GoTo Failed
    Set logger = New Logger
    Set stat = New Status
    logger.Attach state.LogPath
    Set ctx = New InfraContext
    ctx.Init state.ToolId, logger, stat, host, state.Target
    Application.EnableEvents = False
    ToolRegistry.Registry_OnResult state.ToolId, ctx, state.JobId, versionValue
    stat.Clear
    Exit Sub

Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    If Not stat Is Nothing Then stat.Clear
    If Not logger Is Nothing Then logger.Fail "Err " & CStr(failureNumber) & ": " & failureDescription
    If errors Is Nothing Then Set errors = New ErrorHandler
    errors.Report state.ToolId, failureNumber, failureDescription, True
    host.FailJob state.JobId, failureNumber, failureDescription
    If Not batchGuard Is Nothing Then batchGuard.RestoreForReuse
    On Error GoTo 0
End Sub

Public Sub Infra_DispatchFlush(ByVal host As JobHost, ByVal state As JobState, _
                               ByVal batchGuard As AppStateSnapshot)
    Dim logger As Logger
    Dim stat As Status
    Dim ctx As InfraContext
    Dim errors As ErrorHandler
    Dim failureNumber As Long
    Dim failureDescription As String

    On Error GoTo Failed
    Set logger = New Logger
    Set stat = New Status
    logger.Attach state.LogPath
    Set ctx = New InfraContext
    ctx.Init state.ToolId, logger, stat, host, state.Target
    Application.EnableEvents = False
    ToolRegistry.Registry_OnFlush state.ToolId, ctx
    stat.Clear
    Exit Sub

Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    If Not stat Is Nothing Then stat.Clear
    If Not logger Is Nothing Then logger.Fail "Err " & CStr(failureNumber) & ": " & failureDescription
    If errors Is Nothing Then Set errors = New ErrorHandler
    errors.Report state.ToolId, failureNumber, failureDescription, True
    host.FailJob state.JobId, failureNumber, failureDescription
    If Not batchGuard Is Nothing Then batchGuard.RestoreForReuse
    On Error GoTo 0
End Sub

Private Function ResolveTargetWorkbook() As Workbook
    If ActiveWorkbook Is Nothing Then Err.Raise vbObjectError + 2500, "Infra_Run", "Open a target workbook first."
    If ActiveWorkbook Is ThisWorkbook Then Err.Raise vbObjectError + 2501, "Infra_Run", "Activate a target workbook first."
    If ActiveWorkbook.IsAddin Then Err.Raise vbObjectError + 2502, "Infra_Run", "Activate a non-addin workbook first."
    Set ResolveTargetWorkbook = ActiveWorkbook
End Function
