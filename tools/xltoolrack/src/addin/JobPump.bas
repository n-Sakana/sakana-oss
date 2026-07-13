Attribute VB_Name = "JobPump"
Option Explicit

' FE-owned OnTime pump. While jobs are active it wakes once per second,
' pulls worker results from the channel files, and dispatches OnResult.
' Excel natively defers OnTime during cell edits and modal dialogs, so the
' pump never collides with the user; it simply catches up afterwards.
' The pump runs only while jobs exist: the last PumpOnce that reports no
' active jobs lets the chain end.

Private Const TICK_SECONDS As Long = 1
Private Const REARM_GRACE_SECONDS As Long = 5
Private Const PROFILE_CAPACITY As Long = 256

Private m_nextTick As Date
Private m_armed As Boolean
Private m_inTick As Boolean
Private m_profileInitialized As Boolean
Private m_profileEnabled As Boolean
Private m_profileNext As Long
Private m_profileCount As Long
Private m_profileStarted(1 To PROFILE_CAPACITY) As Double
Private m_profileElapsedMs(1 To PROFILE_CAPACITY) As Double
Private m_profileFreshRecords(1 To PROFILE_CAPACITY) As Long
Private m_profileAggregateMs(1 To PROFILE_CAPACITY) As Double
Private m_profileAggregateExistsMs(1 To PROFILE_CAPACITY) As Double
Private m_profileAggregateOpenMs(1 To PROFILE_CAPACITY) As Double
Private m_profileAggregateLineInputMs(1 To PROFILE_CAPACITY) As Double
Private m_profileAggregateDecodeMs(1 To PROFILE_CAPACITY) As Double
Private m_profileAggregateCloseMs(1 To PROFILE_CAPACITY) As Double
Private m_profileFlushTools(1 To PROFILE_CAPACITY) As Long
Private m_profilePumpOnceMs(1 To PROFILE_CAPACITY) As Double
Private m_profileDispatchMs(1 To PROFILE_CAPACITY) As Double
Private m_profileErrorPollMs(1 To PROFILE_CAPACITY) As Double
Private m_profileCleanupMs(1 To PROFILE_CAPACITY) As Double

Public Sub Pump_Start()
    Dim failureNumber As Long
    Dim failureDescription As String
    Dim failureSource As String
    On Error GoTo StartFailed
    If m_armed Then Exit Sub
    ScheduleNext
    Exit Sub
StartFailed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    failureSource = Err.Source
    Err.Raise failureNumber, failureSource, failureDescription
End Sub

Public Sub Pump_Stop()
    If m_armed Then
        On Error Resume Next
        Application.OnTime m_nextTick, QualifiedTick(), , False
        On Error GoTo 0
    End If
    m_armed = False
End Sub

Public Sub Pump_Tick()
    Dim jobs As Object
    Dim keepRunning As Boolean
    Dim profileThisTick As Boolean
    Dim tickStarted As Double
    Dim freshRecords As Long
    Dim aggregateReadMs As Double
    Dim aggregateExistsMs As Double
    Dim aggregateOpenMs As Double
    Dim aggregateLineInputMs As Double
    Dim aggregateDecodeMs As Double
    Dim aggregateCloseMs As Double
    Dim flushTools As Long
    Dim pumpOnceStarted As Double
    Dim pumpOnceMs As Double
    Dim dispatchMs As Double
    Dim errorPollMs As Double
    Dim cleanupMs As Double
    If m_inTick Then Exit Sub
    profileThisTick = ProfileEnabled()
    If profileThisTick Then tickStarted = Timer
    m_armed = False
    m_inTick = True
    On Error Resume Next
    ' Yield before and after the short pull so Excel can service input and
    ' contextual cursor messages. The re-entry guards cover both yields.
    DoEvents
    Set jobs = HostServices.Jobs()
    If Not jobs Is Nothing Then
        If profileThisTick Then
            pumpOnceStarted = Timer
            keepRunning = CBool(jobs.PumpOnce(freshRecords, aggregateReadMs, _
                                              aggregateExistsMs, aggregateOpenMs, _
                                              aggregateLineInputMs, aggregateDecodeMs, _
                                              aggregateCloseMs, flushTools, dispatchMs, errorPollMs, _
                                              cleanupMs, True))
            pumpOnceMs = ElapsedMilliseconds(pumpOnceStarted, Timer)
        Else
            keepRunning = CBool(jobs.PumpOnce())
        End If
    End If
    If keepRunning Then ScheduleNext
    DoEvents
    If profileThisTick Then
        RecordProfile tickStarted, ElapsedMilliseconds(tickStarted, Timer), _
                      freshRecords, aggregateReadMs, aggregateExistsMs, _
                      aggregateOpenMs, aggregateLineInputMs, aggregateDecodeMs, _
                      aggregateCloseMs, flushTools, pumpOnceMs, dispatchMs, _
                      errorPollMs, cleanupMs
    End If
    m_inTick = False
    On Error GoTo 0
End Sub

Public Sub Pump_ProfileResetForTest()
    AssertProfileAccess
    m_profileInitialized = True
    m_profileEnabled = True
    m_profileNext = 1
    m_profileCount = 0
End Sub

Public Function Pump_ProfileForTest() As String
    Dim firstAt As Long
    Dim itemAt As Long
    Dim i As Long
    Dim result As String
    AssertProfileAccess
    result = "StartTimer,ElapsedMs,FreshRecords,AggregateReadMs,AggregateExistsMs," & _
             "AggregateOpenMs,AggregateLineInputMs,AggregateDecodeMs,AggregateCloseMs,FlushTools," & _
             "PumpOnceMs,DispatchMs,ErrorPollMs,CleanupMs" & vbCrLf
    If m_profileCount < PROFILE_CAPACITY Then
        firstAt = 1
    Else
        firstAt = m_profileNext
    End If
    For i = 0 To m_profileCount - 1
        itemAt = ((firstAt - 1 + i) Mod PROFILE_CAPACITY) + 1
        result = result & Trim$(Str$(m_profileStarted(itemAt))) & "," & _
                 Trim$(Str$(m_profileElapsedMs(itemAt))) & "," & _
                 CStr(m_profileFreshRecords(itemAt)) & "," & _
                 Trim$(Str$(m_profileAggregateMs(itemAt))) & "," & _
                 Trim$(Str$(m_profileAggregateExistsMs(itemAt))) & "," & _
                 Trim$(Str$(m_profileAggregateOpenMs(itemAt))) & "," & _
                 Trim$(Str$(m_profileAggregateLineInputMs(itemAt))) & "," & _
                 Trim$(Str$(m_profileAggregateDecodeMs(itemAt))) & "," & _
                 Trim$(Str$(m_profileAggregateCloseMs(itemAt))) & "," & _
                 CStr(m_profileFlushTools(itemAt)) & "," & _
                 Trim$(Str$(m_profilePumpOnceMs(itemAt))) & "," & _
                 Trim$(Str$(m_profileDispatchMs(itemAt))) & "," & _
                 Trim$(Str$(m_profileErrorPollMs(itemAt))) & "," & _
                 Trim$(Str$(m_profileCleanupMs(itemAt))) & vbCrLf
    Next i
    Pump_ProfileForTest = result
End Function

' Watchdog, called from app events (sheet change / activate): if the chain
' was lost while jobs are active - a tick error, a canceled schedule - it is
' re-armed here. Duplicate ticks are harmless: PumpOnce dedups by version.
Public Sub Pump_EnsureArmed()
    Dim jobs As Object
    If m_inTick Then Exit Sub
    If m_armed Then
        If Now <= DateAdd("s", REARM_GRACE_SECONDS, m_nextTick) Then Exit Sub
    End If
    On Error Resume Next
    Set jobs = HostServices.Jobs()
    If jobs Is Nothing Then Exit Sub
    If CBool(jobs.HasActiveJobs()) Then
        m_armed = False
        Pump_Start
    End If
    On Error GoTo 0
End Sub

Private Sub ScheduleNext()
    m_nextTick = Now + TimeSerial(0, 0, TICK_SECONDS)
    Application.OnTime m_nextTick, QualifiedTick()
    m_armed = True
End Sub

Private Function QualifiedTick() As String
    QualifiedTick = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!JobPump.Pump_Tick"
End Function

Private Function ProfileEnabled() As Boolean
    If Not m_profileInitialized Then
        m_profileEnabled = (Environ$("XLTOOLRACK_TEST") = "1" Or Environ$("XLTOOLRACK_PROFILE") = "1")
        m_profileNext = 1
        m_profileInitialized = True
    End If
    ProfileEnabled = m_profileEnabled
End Function

Private Sub AssertProfileAccess()
    If Environ$("XLTOOLRACK_TEST") = "1" Then Exit Sub
    If Environ$("XLTOOLRACK_PROFILE") = "1" Then Exit Sub
    Err.Raise vbObjectError + 2710, "JobPump.PumpProfile", "test or profile mode required"
End Sub

Private Sub RecordProfile(ByVal startedAt As Double, ByVal elapsedMs As Double, _
                          ByVal freshRecords As Long, ByVal aggregateMs As Double, _
                          ByVal aggregateExistsMs As Double, ByVal aggregateOpenMs As Double, _
                          ByVal aggregateLineInputMs As Double, ByVal aggregateDecodeMs As Double, _
                          ByVal aggregateCloseMs As Double, _
                          ByVal flushTools As Long, ByVal pumpOnceMs As Double, _
                          ByVal dispatchMs As Double, ByVal errorPollMs As Double, _
                          ByVal cleanupMs As Double)
    On Error Resume Next
    m_profileStarted(m_profileNext) = startedAt
    m_profileElapsedMs(m_profileNext) = elapsedMs
    m_profileFreshRecords(m_profileNext) = freshRecords
    m_profileAggregateMs(m_profileNext) = aggregateMs
    m_profileAggregateExistsMs(m_profileNext) = aggregateExistsMs
    m_profileAggregateOpenMs(m_profileNext) = aggregateOpenMs
    m_profileAggregateLineInputMs(m_profileNext) = aggregateLineInputMs
    m_profileAggregateDecodeMs(m_profileNext) = aggregateDecodeMs
    m_profileAggregateCloseMs(m_profileNext) = aggregateCloseMs
    m_profileFlushTools(m_profileNext) = flushTools
    m_profilePumpOnceMs(m_profileNext) = pumpOnceMs
    m_profileDispatchMs(m_profileNext) = dispatchMs
    m_profileErrorPollMs(m_profileNext) = errorPollMs
    m_profileCleanupMs(m_profileNext) = cleanupMs
    m_profileNext = m_profileNext + 1
    If m_profileNext > PROFILE_CAPACITY Then m_profileNext = 1
    If m_profileCount < PROFILE_CAPACITY Then m_profileCount = m_profileCount + 1
    On Error GoTo 0
End Sub

Private Function ElapsedMilliseconds(ByVal startedAt As Double, ByVal endedAt As Double) As Double
    If endedAt < startedAt Then endedAt = endedAt + 86400#
    ElapsedMilliseconds = (endedAt - startedAt) * 1000#
End Function
