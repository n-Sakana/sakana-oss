Attribute VB_Name = "JobTestApi"
Option Explicit
Option Private Module

Public Function JobTest_StartProbeWorkers(ByVal count As Long) As String
    Dim jobs As Object
    Dim target As Workbook
    Dim i As Long
    Dim jobId As String
    Dim result As String
    HostMain.InitAddin
    Set jobs = HostServices.Jobs()
    Set target = ActiveWorkbook
    For i = 1 To count
        jobId = CStr(jobs.StartJob("probe_worker", "probe_worker.Worker", "{}", target))
        If Len(result) > 0 Then result = result & ","
        result = result & jobId
    Next i
    JobTest_StartProbeWorkers = result
End Function

Public Function JobTest_TryStart(ByVal toolId As String) As String
    Dim jobs As Object
    Dim target As Workbook
    Dim jobId As String
    On Error GoTo Failed
    HostMain.InitAddin
    Set jobs = HostServices.Jobs()
    Set target = ActiveWorkbook
    jobId = CStr(jobs.StartJob(toolId, toolId & ".Worker", "{}", target))
    JobTest_TryStart = "OK:" & jobId
    Exit Function
Failed:
    JobTest_TryStart = "ERROR:" & Err.Description
End Function

Public Function JobTest_VersionSnapshot(ByVal csvJobIds As String) As String
    Dim jobs As Object
    Dim ids As Variant
    Dim i As Long
    Dim result As String
    Set jobs = HostServices.Jobs()
    ids = Split(csvJobIds, ",")
    For i = LBound(ids) To UBound(ids)
        If Len(result) > 0 Then result = result & ","
        result = result & CStr(jobs.VersionFor(CStr(ids(i))))
    Next i
    JobTest_VersionSnapshot = result
End Function

Public Function JobTest_RunningCount() As Long
    Dim jobs As Object
    Set jobs = HostServices.Jobs()
    If jobs Is Nothing Then Exit Function
    JobTest_RunningCount = CLng(jobs.RunningCount())
End Function

Public Sub JobTest_SetGlobalLimit(ByVal value As Long)
    Dim jobs As Object
    HostMain.InitAddin
    Set jobs = HostServices.Jobs()
    jobs.SetGlobalLimitForTest value
End Sub

Public Sub JobTest_StopAll()
    Dim jobs As Object
    Set jobs = HostServices.Jobs()
    If jobs Is Nothing Then Exit Sub
    jobs.StopAll
End Sub

Public Function JobTest_DonePaths(ByVal csvJobIds As String) As String
    Dim jobs As Object
    Dim ids As Variant
    Dim i As Long
    Dim result As String
    Set jobs = HostServices.Jobs()
    ids = Split(csvJobIds, ",")
    For i = LBound(ids) To UBound(ids)
        If Len(result) > 0 Then result = result & "|"
        result = result & CStr(jobs.DonePathFor(CStr(ids(i))))
    Next i
    JobTest_DonePaths = result
End Function

Public Function JobTest_StartErrorWorker() As String
    Dim jobs As Object
    HostMain.InitAddin
    Set jobs = HostServices.Jobs()
    JobTest_StartErrorWorker = CStr(jobs.StartJob("probe_worker_error", "probe_worker_error.Worker", "{}", ActiveWorkbook))
End Function

Public Function JobTest_StateFor(ByVal jobId As String) As String
    Dim jobs As Object
    Set jobs = HostServices.Jobs()
    JobTest_StateFor = CStr(jobs.StateFor(jobId))
End Function

Public Function JobTest_ErrorFor(ByVal jobId As String) As String
    Dim jobs As Object
    Set jobs = HostServices.Jobs()
    JobTest_ErrorFor = CStr(jobs.ErrorFor(jobId))
End Function

Public Function JobTest_DonePathFor(ByVal jobId As String) As String
    Dim jobs As Object
    Set jobs = HostServices.Jobs()
    JobTest_DonePathFor = CStr(jobs.DonePathFor(jobId))
End Function

Public Function JobTest_LogPathFor(ByVal jobId As String) As String
    Dim jobs As Object
    Set jobs = HostServices.Jobs()
    JobTest_LogPathFor = CStr(jobs.LogPathFor(jobId))
End Function

Public Function JobTest_WorkerCopyFor(ByVal jobId As String) As String
    Dim jobs As Object
    Set jobs = HostServices.Jobs()
    JobTest_WorkerCopyFor = CStr(jobs.WorkerCopyFor(jobId))
End Function

Public Sub JobTest_SetFailStage(ByVal value As String)
    Dim jobs As Object
    HostMain.InitAddin
    Set jobs = HostServices.Jobs()
    jobs.SetFailStageForTest value
End Sub

Public Sub JobTest_ResetPumpProfile()
    HostMain.InitAddin
    JobPump.Pump_ProfileResetForTest
End Sub

Public Function JobTest_PumpProfile() As String
    JobTest_PumpProfile = JobPump.Pump_ProfileForTest()
End Function

Public Sub JobTest_DetachForFeClose()
    HostMain.DetachForTest
End Sub
