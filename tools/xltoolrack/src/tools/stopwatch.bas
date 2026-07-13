Attribute VB_Name = "stopwatch"
Option Explicit
Option Private Module
'@name Multi Stopwatch
'@ribbon Multi Stopwatch
'@group Tools
'@maxjobs 3

Private Const DISPLAY_SHEET As String = "xtr_stopwatch"
Private m_rows As Object
Private m_workerStarted As Boolean
Private m_startedAt As Date
Private m_elapsed As Long
Private m_displayValues(1 To 3, 1 To 3) As Variant
Private m_displayDirty As Boolean

Public Sub Run(ByVal ctx As InfraContext)
    Dim sheet As Worksheet
    Dim i As Long
    Dim jobId As String

    Set sheet = EnsureDisplaySheet(ctx.Target)
    PrepareDisplay sheet
    Set m_rows = CreateObject("Scripting.Dictionary")
    m_rows.CompareMode = 1
    m_displayDirty = False
    For i = 1 To 3
        jobId = ctx.RunJob("stopwatch.Worker", "{}")
        m_rows.Add jobId, i + 2
        sheet.Cells(i + 2, 5).Value2 = jobId
        m_displayValues(i, 1) = 0
        m_displayValues(i, 2) = 0
        m_displayValues(i, 3) = "starting"
    Next i
    ctx.LogMessage "started 3 stopwatch jobs"
End Sub

Public Sub Worker(ByVal job As InfraJob)
    Dim values(1 To 1, 1 To 2) As Variant
    Dim nextElapsed As Long
    If Not m_workerStarted Then
        m_startedAt = Now
        m_elapsed = 0
        m_workerStarted = True
    End If
    nextElapsed = DateDiff("s", m_startedAt, Now) + 1
    If nextElapsed <= m_elapsed Then nextElapsed = m_elapsed + 1
    values(1, 1) = nextElapsed
    values(1, 2) = CDbl(Now)
    If job.Push("state", "A1:B1", values) Then m_elapsed = nextElapsed
End Sub

Public Sub OnResult(ByVal ctx As InfraContext, ByVal jobId As String, ByVal version As Long)
    Dim values As Variant
    Dim rowNumber As Long
    Dim rowIndex As Long
    If m_rows Is Nothing Then Exit Sub
    If Not m_rows.Exists(jobId) Then Exit Sub
    rowNumber = CLng(m_rows(jobId))
    rowIndex = rowNumber - 2
    values = ctx.ReadJob(jobId, "A1:B1")
    m_displayValues(rowIndex, 1) = CLng(values(1, 1))
    m_displayValues(rowIndex, 2) = version
    m_displayValues(rowIndex, 3) = "running"
    m_displayDirty = True
End Sub

Public Sub OnFlush(ByVal ctx As InfraContext)
    Dim sheet As Worksheet
    If Not m_displayDirty Then Exit Sub
    Set sheet = ctx.Target.Worksheets(DISPLAY_SHEET)
    sheet.Range("B3:D5").Value2 = m_displayValues
    m_displayDirty = False
End Sub

Private Function EnsureDisplaySheet(ByVal workbook As Workbook) As Worksheet
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = workbook.Worksheets(DISPLAY_SHEET)
    On Error GoTo 0
    If sheet Is Nothing Then
        Set sheet = workbook.Worksheets.Add(After:=workbook.Worksheets(workbook.Worksheets.Count))
        sheet.Name = DISPLAY_SHEET
    End If
    Set EnsureDisplaySheet = sheet
End Function

Private Sub PrepareDisplay(ByVal sheet As Worksheet)
    Dim i As Long
    sheet.Cells.Clear
    sheet.Range("A1:E1").Merge
    sheet.Range("A1").Value2 = "MULTI STOPWATCH"
    sheet.Range("A1").Font.Bold = True
    sheet.Range("A1").Font.Size = 20
    sheet.Range("A1").HorizontalAlignment = xlCenter
    sheet.Range("A1").Interior.Color = RGB(29, 44, 63)
    sheet.Range("A1").Font.Color = RGB(255, 255, 255)
    sheet.Range("A2:E2").Value = Array("Timer", "Seconds", "Version", "State", "Job")
    sheet.Range("A2:E2").Font.Bold = True
    sheet.Range("A2:E2").Interior.Color = RGB(78, 121, 167)
    sheet.Range("A2:E2").Font.Color = RGB(255, 255, 255)
    For i = 1 To 3
        sheet.Cells(i + 2, 1).Value2 = "Stopwatch " & CStr(i)
        sheet.Cells(i + 2, 2).Value2 = 0
        sheet.Cells(i + 2, 3).Value2 = 0
        sheet.Cells(i + 2, 4).Value2 = "starting"
        sheet.Cells(i + 2, 2).Font.Size = 28
        sheet.Cells(i + 2, 2).Font.Bold = True
        sheet.Rows(i + 2).RowHeight = 38
    Next i
    sheet.Columns("A").ColumnWidth = 18
    sheet.Columns("B").ColumnWidth = 16
    sheet.Columns("C:D").ColumnWidth = 12
    sheet.Columns("E").Hidden = True
    sheet.Range("A2:D5").Borders.LineStyle = xlContinuous
End Sub
