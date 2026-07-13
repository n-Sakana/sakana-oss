Attribute VB_Name = "PathFs"
Option Explicit
Option Private Module

' One stable id per FE session. Application.hWnd is NOT usable for this: it
' can return different values between add-in startup (Workbook_Open) and
' later calls, which once pointed workers at a lease file that was never
' created. Generated once and cached for the session.
Private m_sessionId As String

Private Const CHANNEL_FOLDER As String = "xltoolrack"

Public Function JoinPath(ByVal leftPath As String, ByVal rightPath As String) As String
    If Right$(leftPath, 1) = Application.PathSeparator Then
        JoinPath = leftPath & rightPath
    Else
        JoinPath = leftPath & Application.PathSeparator & rightPath
    End If
End Function

Public Function SessionId() As String
    If Len(m_sessionId) = 0 Then
        m_sessionId = Format$(Now, "yyyymmddhhnnss") & Format$(CLng(Timer * 1000!) Mod 100000, "00000")
    End If
    SessionId = m_sessionId
End Function

' Session-scoped flag the FE writes at shutdown. Workers treat it as a
' definite "the FE is closing" signal.
Public Function FeShutdownFlagPath() As String
    FeShutdownFlagPath = JoinPath(Environ$("TEMP"), "xltoolrack_fe_gone_" & SessionId() & ".flag")
End Function

' Lease file the FE keeps open (share-deny) for its whole lifetime. Workers
' probe the write lock: locked means the FE process is alive even while the
' user edits a cell; unlocked or missing means it is gone (including crashes,
' which release the lock instantly). Pure file IO, no COM.
Public Function FeLeasePath() As String
    FeLeasePath = JoinPath(Environ$("TEMP"), "xltoolrack_fe_lease_" & SessionId() & ".lock")
End Function

' Dedicated directory for the high-churn channel files. Keeping only these
' files in one narrow subtree allows a reversible Defender exclusion test
' without excluding TEMP itself or unrelated application data.
Public Function ChannelRootPath() As String
    Dim root As String
    root = JoinPath(Environ$("TEMP"), CHANNEL_FOLDER)
    If Len(Dir$(root, vbDirectory)) = 0 Then MkDir root
    ChannelRootPath = root
End Function

' Base path for one job channel's files (payload/input/stop/error).
Public Function ChannelBasePath(ByVal channelId As String) As String
    ChannelBasePath = JoinPath(ChannelRootPath(), "xltoolrack_ch_" & SessionId() & "_" & channelId)
End Function
