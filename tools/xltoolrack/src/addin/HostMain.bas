Attribute VB_Name = "HostMain"
Option Explicit

Private m_events As AppEventHandler
Private m_shuttingDown As Boolean
Private m_leaseFile As Integer

Public Sub InitAddin()
    Dim jobs As JobHost
    Dim currentJobs As Object
    On Error Resume Next
    If Len(Dir$(PathFs.FeShutdownFlagPath())) > 0 Then Kill PathFs.FeShutdownFlagPath()
    SweepStaleSessionFiles
    On Error GoTo InitFailed
    EnsureLease
    Set currentJobs = HostServices.Jobs()
    If currentJobs Is Nothing Then
        Set jobs = New JobHost
        jobs.Init ThisWorkbook
        HostServices.SetJobs jobs
    End If
    If m_events Is Nothing Then
        Set m_events = New AppEventHandler
        m_events.Init Application, HostServices.Jobs()
    End If
    Exit Sub
InitFailed:
    Debug.Print "!! HostMain.InitAddin | Err " & CStr(Err.Number) & ": " & Err.Description
End Sub

Public Sub Shutdown()
    Dim jobs As Object
    If m_shuttingDown Then Exit Sub
    m_shuttingDown = True
    On Error Resume Next
    WriteShutdownFlag
    Set jobs = HostServices.Jobs()
    If Not jobs Is Nothing Then jobs.StopAll
    JobPump.Pump_Stop
    If Not m_events Is Nothing Then Set m_events.App = Nothing
    Set m_events = Nothing
    HostServices.Clear
    ReleaseLease
    ThisWorkbook.Saved = True
    m_shuttingDown = False
    On Error GoTo 0
End Sub

' The lease stays open (share-deny) until Shutdown; workers probe its lock
' to tell "FE busy or editing" apart from "FE gone" without any COM call.
Public Sub EnsureLease()
    Dim fileNo As Integer
    If m_leaseFile <> 0 Then Exit Sub
    On Error Resume Next
    If Len(Dir$(PathFs.FeLeasePath())) > 0 Then Kill PathFs.FeLeasePath()
    Err.Clear
    fileNo = FreeFile
    Open PathFs.FeLeasePath() For Output Lock Read Write As #fileNo
    If Err.Number = 0 Then
        Print #fileNo, CStr(Application.hWnd)
        m_leaseFile = fileNo
    Else
        Close #fileNo
    End If
    On Error GoTo 0
End Sub

Private Sub ReleaseLease()
    If m_leaseFile = 0 Then Exit Sub
    On Error Resume Next
    Close #m_leaseFile
    m_leaseFile = 0
    Kill PathFs.FeLeasePath()
    On Error GoTo 0
End Sub

' Written before StopAll so a worker that never sees its stop flag (the FE
' can finish quitting first) still learns the FE is gone, not merely busy.
Private Sub WriteShutdownFlag()
    Dim fileNo As Integer
    On Error Resume Next
    fileNo = FreeFile
    Open PathFs.FeShutdownFlagPath() For Output As #fileNo
    Print #fileNo, "closed_at=" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Close #fileNo
    On Error GoTo 0
End Sub

' Best effort removal of files left by long-gone sessions (a live session's
' lease is locked, so Kill skips it). Keeps TEMP from accumulating litter.
Private Sub SweepStaleSessionFiles()
    Dim folder As String
    On Error Resume Next
    folder = Environ$("TEMP") & Application.PathSeparator
    SweepStaleFilesInFolder folder, Array("xltoolrack_fe_gone_*.flag", _
                                          "xltoolrack_fe_lease_*.lock", _
                                          "xltoolrack_ch_*", _
                                          "xltoolrack_done_*.flag", _
                                          "xltoolrack_worker_*.xlsm")
    folder = PathFs.ChannelRootPath() & Application.PathSeparator
    SweepStaleFilesInFolder folder, Array("xltoolrack_ch_*")
    On Error GoTo 0
End Sub

Private Sub SweepStaleFilesInFolder(ByVal folder As String, ByVal patterns As Variant)
    Dim i As Long
    Dim name As String
    Dim fullPath As String
    On Error Resume Next
    For i = LBound(patterns) To UBound(patterns)
        name = Dir$(folder & CStr(patterns(i)))
        Do While Len(name) > 0
            fullPath = folder & name
            If FileDateTime(fullPath) < DateAdd("d", -1, Now) Then Kill fullPath
            name = Dir$()
        Loop
    Next i
    On Error GoTo 0
End Sub

Public Sub DetachForTest()
    If Environ$("XLTOOLRACK_TEST") <> "1" Then Err.Raise vbObjectError + 2520, "HostMain.DetachForTest", "test mode required"
    On Error Resume Next
    If Not m_events Is Nothing Then Set m_events.App = Nothing
    Set m_events = Nothing
    HostServices.Clear
    On Error GoTo 0
End Sub
