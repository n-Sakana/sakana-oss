VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ContextHostProbe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_started As Long
Private m_stopped As Long
Private m_lastTool As String
Private m_lastWorker As String
Private m_lastArgs As String
Private m_lastTarget As Object

Public Function StartJob(ByVal toolId As String, ByVal workerEntry As String, _
                         ByVal argsJson As String, ByVal target As Object) As String
    m_started = m_started + 1
    m_lastTool = toolId
    m_lastWorker = workerEntry
    m_lastArgs = argsJson
    Set m_lastTarget = target
    StartJob = "probe_job_" & CStr(m_started)
End Function

Public Sub StopJob(ByVal jobId As String)
    m_stopped = m_stopped + 1
End Sub

Public Function ReadPayload(ByVal jobId As String, ByVal address As String) As Variant
    ReadPayload = "payload"
End Function

Public Sub RegisterInput(ByVal jobId As String, ByVal sheet As Worksheet, ByVal address As String)
End Sub

Public Property Get StartedCount() As Long
    StartedCount = m_started
End Property

Public Property Get StoppedCount() As Long
    StoppedCount = m_stopped
End Property

Public Property Get LastTool() As String
    LastTool = m_lastTool
End Property

Public Property Get LastWorker() As String
    LastWorker = m_lastWorker
End Property

Public Property Get LastArgs() As String
    LastArgs = m_lastArgs
End Property

Public Property Get LastTarget() As Object
    Set LastTarget = m_lastTarget
End Property

