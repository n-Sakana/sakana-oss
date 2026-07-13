Attribute VB_Name = "SelfTests"
Option Explicit
Option Private Module

Public Function AppGuard_SelfTest() As String
    Dim oldScreen As Boolean
    Dim oldAlerts As Boolean
    Dim oldCalc As Long
    Dim ok As Boolean

    On Error GoTo Failed
    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts
    oldCalc = Application.Calculation
    ok = True

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    GuardNormal ok
    GuardNested ok

    On Error Resume Next
    GuardError
    Err.Clear
    On Error GoTo Failed
    If Application.ScreenUpdating <> True Then ok = False
    If Application.DisplayAlerts <> True Then ok = False
    If Application.Calculation <> xlCalculationAutomatic Then ok = False

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    GuardSavedFalse ok

    AppGuard_SelfTest = IIf(ok, "OK", "FAIL")
    GoTo CleanExit

Failed:
    AppGuard_SelfTest = "FAIL:" & CStr(Err.Number)

CleanExit:
    On Error Resume Next
    Application.Calculation = oldCalc
    Application.DisplayAlerts = oldAlerts
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0
End Function

Private Sub GuardNormal(ByRef ok As Boolean)
    Dim guard As AppGuard
    Set guard = New AppGuard
    If Application.ScreenUpdating <> False Then ok = False
    If Application.DisplayAlerts <> False Then ok = False
    If Application.Calculation <> xlCalculationManual Then ok = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Set guard = Nothing
    If Application.ScreenUpdating <> True Then ok = False
    If Application.DisplayAlerts <> True Then ok = False
    If Application.Calculation <> xlCalculationAutomatic Then ok = False
End Sub

Private Sub GuardNested(ByRef ok As Boolean)
    Dim outerGuard As AppGuard
    Dim innerGuard As AppGuard
    Set outerGuard = New AppGuard
    Set innerGuard = New AppGuard
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Set innerGuard = Nothing
    If Application.ScreenUpdating <> False Then ok = False
    If Application.DisplayAlerts <> False Then ok = False
    If Application.Calculation <> xlCalculationManual Then ok = False
    Set outerGuard = Nothing
    If Application.ScreenUpdating <> True Then ok = False
    If Application.DisplayAlerts <> True Then ok = False
    If Application.Calculation <> xlCalculationAutomatic Then ok = False
End Sub

Private Sub GuardError()
    Dim guard As AppGuard
    Set guard = New AppGuard
    Err.Raise 5, "SelfTests.GuardError", "probe"
End Sub

Private Sub GuardSavedFalse(ByRef ok As Boolean)
    Dim guard As AppGuard
    Set guard = New AppGuard
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Set guard = Nothing
    If Application.ScreenUpdating <> False Then ok = False
    If Application.DisplayAlerts <> False Then ok = False
    If Application.Calculation <> xlCalculationManual Then ok = False
End Sub

Public Function LogStatus_SelfTest() As String
    Dim log As Logger
    Dim stat As Status
    Dim shown As String
    Dim restored As String

    On Error GoTo Failed
    Set log = New Logger
    Set stat = New Status
    log.Begin "selftest"
    log.WriteLine "hello from selftest"
    stat.Begin "selftest"
    stat.Show 0.5, "probe-status"
    shown = CStr(Application.StatusBar)
    stat.Clear
    restored = CStr(Application.StatusBar)
    log.Done
    LogStatus_SelfTest = log.Path & "|" & shown & "|" & restored
    Exit Function

Failed:
    On Error Resume Next
    If Not stat Is Nothing Then stat.Clear
    If Not log Is Nothing Then log.Fail "Err " & CStr(Err.Number)
    LogStatus_SelfTest = "FAIL||"
    On Error GoTo 0
End Function

Public Function Context_SelfTest() As String
    Dim logger As Logger
    Dim stat As Status
    Dim host As ContextHostProbe
    Dim ctx As InfraContext
    Dim target As Workbook
    Dim firstJob As String
    Dim ok As Boolean
    Dim stage As String
    Dim failureNumber As Long

    On Error GoTo Failed
    stage = "target"
    Set target = ActiveWorkbook
    stage = "objects"
    Set logger = New Logger
    Set stat = New Status
    Set host = New ContextHostProbe
    Set ctx = New InfraContext
    logger.Begin "context_selftest"
    stat.Begin "context_selftest"
    stage = "init"
    ctx.Init "probe_tool", logger, stat, host, target
    ok = True

    stage = "log_status"
    ctx.LogMessage "delegated log"
    ctx.Status 0.5, "delegated status"
    If CStr(Application.StatusBar) <> "delegated status" Then ok = False
    If Not ctx.Target Is target Then ok = False

    stage = "first_job"
    firstJob = ctx.RunJob("probe.Worker", "{x:1}")
    If firstJob <> "probe_job_1" Then ok = False
    If host.LastTool <> "probe_tool" Then ok = False
    If host.LastWorker <> "probe.Worker" Then ok = False
    If host.LastArgs <> "{x:1}" Then ok = False
    If Not host.LastTarget Is target Then ok = False
    stage = "more_jobs"
    ctx.StopJob firstJob
    Call ctx.RunJob("probe.Worker", "{}")
    Call ctx.RunJob("probe.Worker", "{}")
    ctx.RollbackJobs
    If host.StartedCount <> 3 Then ok = False
    If host.StoppedCount <> 4 Then ok = False

    stage = "infra_job"
    Dim job As InfraJob
    Set job = New InfraJob
    job.InitLocal "args", target
    If job.Args <> "args" Then ok = False
    If Not job.FEWorkbook Is target Then ok = False
    If Not job.Alive Then ok = False

    stat.Clear
    logger.Done
    Context_SelfTest = IIf(ok, "OK", "FAIL")
    Exit Function

Failed:
    failureNumber = Err.Number
    On Error Resume Next
    stat.Clear
    logger.Fail "Err " & CStr(Err.Number)
    Context_SelfTest = "FAIL:" & CStr(failureNumber) & ":" & stage
    On Error GoTo 0
End Function
