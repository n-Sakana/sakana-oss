Attribute VB_Name = "WorkerBridge"
Option Explicit

Private m_worker As JobWorker
Private m_bootstrapAt As Date

Public Sub WorkerBootstrap(ByVal jobId As String, ByVal toolId As String, _
                           ByVal argsJson As String, ByVal donePath As String, ByVal token As String, _
                           ByVal channelBase As String, ByVal logPath As String, _
                           ByVal feGonePath As String, ByVal feLeasePath As String)
    On Error GoTo Failed
    Set m_worker = New JobWorker
    m_worker.Init jobId, toolId, argsJson, donePath, token, channelBase, logPath, feGonePath, feLeasePath
    m_bootstrapAt = Now + TimeSerial(0, 0, 1)
    Application.OnTime m_bootstrapAt, QualifiedMacro("WorkerBridge.WorkerDispatch")
    Exit Sub
Failed:
    WriteBootstrapFailure donePath, jobId, Err.Number
    Application.DisplayAlerts = False
    Application.Quit
End Sub

Public Sub WorkerDispatch()
    On Error GoTo Failed
    If m_worker Is Nothing Then Err.Raise vbObjectError + 2600, "WorkerBridge.WorkerDispatch", "Worker is not initialized."
    m_worker.RunLoop
    Exit Sub
Failed:
    Application.DisplayAlerts = False
    Application.Quit
End Sub

Public Sub WorkerBeforeClose()
    On Error Resume Next
    If Not m_worker Is Nothing Then m_worker.BeforeClose
    ThisWorkbook.Saved = True
    On Error GoTo 0
End Sub

Private Function QualifiedMacro(ByVal procedureName As String) As String
    QualifiedMacro = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & procedureName
End Function

Private Sub WriteBootstrapFailure(ByVal donePath As String, ByVal jobId As String, ByVal number As Long)
    Dim fileNo As Integer
    If Len(donePath) = 0 Then Exit Sub
    On Error Resume Next
    fileNo = FreeFile
    Open donePath For Output As #fileNo
    Print #fileNo, "job_id=" & jobId
    Print #fileNo, "reason=bootstrap_error_" & CStr(number)
    Close #fileNo
    On Error GoTo 0
End Sub
