Attribute VB_Name = "HostServices"
Option Explicit
Option Private Module

Private m_jobs As Object

Public Sub SetJobs(ByVal jobs As Object)
    Set m_jobs = jobs
End Sub

Public Function Jobs() As Object
    Set Jobs = m_jobs
End Function

Public Sub Clear()
    Set m_jobs = Nothing
End Sub

