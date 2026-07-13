Attribute VB_Name = "pi_race"
Option Explicit
Option Private Module
'@name Parallel Pi Race
'@ribbon Parallel Pi Race
'@group Tools
'@maxjobs 4

Private Const DISPLAY_SHEET As String = "xtr_pi_race"
Private m_rows As Object
Private m_iteration As Double
Private m_sum As Double
Private m_addNext As Boolean
Private m_initialized As Boolean

Public Sub Run(ByVal ctx As InfraContext)
    Dim sheet As Worksheet
    Dim i As Long
    Dim jobId As String
    Set sheet = EnsureDisplaySheet(ctx.Target)
    PrepareDisplay sheet
    Set m_rows = CreateObject("Scripting.Dictionary")
    m_rows.CompareMode = 1
    For i = 1 To 3
        jobId = ctx.RunJob("pi_race.Worker", "{}")
        m_rows.Add jobId, i + 2
        sheet.Cells(i + 2, 6).Value2 = jobId
    Next i
    ctx.LogMessage "started 3 pi workers"
End Sub

Public Sub Worker(ByVal job As InfraJob)
    Dim startedAt As Single
    Dim batch As Long
    Dim denominator As Double
    Dim estimate As Double
    Dim values(1 To 1, 1 To 3) As Variant

    startedAt = Timer
    If Not m_initialized Then
        m_addNext = True
        m_initialized = True
    End If
    Do
        For batch = 1 To 5000
            denominator = m_iteration * 2# + 1#
            If m_addNext Then
                m_sum = m_sum + 1# / denominator
            Else
                m_sum = m_sum - 1# / denominator
            End If
            m_addNext = Not m_addNext
            m_iteration = m_iteration + 1#
        Next batch
    Loop While ElapsedSince(startedAt) < 0.75!

    estimate = 4# * m_sum
    values(1, 1) = m_iteration
    values(1, 2) = estimate
    values(1, 3) = CDbl(Now)
    ' A skipped publish is fine: the totals are cumulative and the next tick
    ' pushes fresher numbers anyway.
    Call job.Push("progress", "A1:C1", values)
End Sub

Public Sub OnResult(ByVal ctx As InfraContext, ByVal jobId As String, ByVal version As Long)
    Dim values As Variant
    Dim displayValues(1 To 1, 1 To 4) As Variant
    Dim sheet As Worksheet
    Dim rowNumber As Long
    If m_rows Is Nothing Then Exit Sub
    If Not m_rows.Exists(jobId) Then Exit Sub
    rowNumber = CLng(m_rows(jobId))
    values = ctx.ReadJob(jobId, "A1:C1")
    Set sheet = ctx.Target.Worksheets(DISPLAY_SHEET)
    displayValues(1, 1) = CDbl(values(1, 1))
    displayValues(1, 2) = CDbl(values(1, 2))
    displayValues(1, 3) = version
    displayValues(1, 4) = "running"
    sheet.Range("B" & CStr(rowNumber) & ":E" & CStr(rowNumber)).Value2 = displayValues
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
    sheet.Range("A1:F1").Merge
    sheet.Range("A1").Value2 = "PARALLEL PI RACE"
    sheet.Range("A1").Font.Bold = True
    sheet.Range("A1").Font.Size = 20
    sheet.Range("A1").HorizontalAlignment = xlCenter
    sheet.Range("A1").Interior.Color = RGB(89, 44, 92)
    sheet.Range("A1").Font.Color = RGB(255, 255, 255)
    sheet.Range("A2:F2").Value = Array("Worker", "Iterations", "Pi estimate", "Version", "State", "Job")
    sheet.Range("A2:F2").Font.Bold = True
    sheet.Range("A2:F2").Interior.Color = RGB(176, 122, 161)
    sheet.Range("A2:F2").Font.Color = RGB(255, 255, 255)
    For i = 1 To 3
        sheet.Cells(i + 2, 1).Value2 = "Core worker " & CStr(i)
        sheet.Cells(i + 2, 2).Value2 = 0
        sheet.Cells(i + 2, 3).Value2 = 0
        sheet.Cells(i + 2, 3).NumberFormat = "0.000000000000000"
        sheet.Cells(i + 2, 4).Value2 = 0
        sheet.Cells(i + 2, 5).Value2 = "starting"
        sheet.Rows(i + 2).RowHeight = 28
    Next i
    sheet.Columns("A").ColumnWidth = 18
    sheet.Columns("B").ColumnWidth = 18
    sheet.Columns("C").ColumnWidth = 22
    sheet.Columns("D:E").ColumnWidth = 12
    sheet.Columns("F").Hidden = True
    sheet.Range("A2:E5").Borders.LineStyle = xlContinuous
End Sub

Private Function ElapsedSince(ByVal startedAt As Single) As Single
    Dim value As Single
    value = Timer - startedAt
    If value < 0 Then value = value + 86400!
    ElapsedSince = value
End Function
