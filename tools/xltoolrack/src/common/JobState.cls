VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JobState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public JobId As String
Public ChannelId As String
Public SessionToken As String
Public ToolId As String
Public WorkerEntry As String
Public ArgsJson As String
Public RowNumber As Long
Public PayloadSheetName As String
Public ChannelBase As String
Public InputVersion As Long
Public DonePath As String
Public WorkerCopy As String
Public LogPath As String
Public LastVersion As Long
Public StateName As String
Public HoldUntil As Date
Public LastPayloadTick As Double
Public CachedPayloadAddress As String
Public CachedPayloadValues As Variant
Public CachedPayloadDirty As Boolean
Public Target As Workbook
Public BeApp As Object
Public BeBook As Object
Public InputTargetSheet As Worksheet
Public InputTargetAddress As String

Public Function IsActive() As Boolean
    IsActive = (StateName = "starting" Or StateName = "running" Or StateName = "stopping")
End Function
