Attribute VB_Name = "life"
Option Explicit
Option Private Module
'@name Conway Life
'@ribbon Conway Life
'@group Tools
'@maxjobs 1

Private Const DISPLAY_SHEET As String = "xtr_life"
Private Const GRID_ADDRESS As String = "B3:U22"
Private Const GRID_SIZE As Long = 20
Private m_jobId As String
Private m_generation As Long
Private m_lastInputVersion As Long
Private m_board As Variant

Public Sub Run(ByVal ctx As InfraContext)
    Dim sheet As Worksheet
    Set sheet = EnsureDisplaySheet(ctx.Target)
    PrepareDisplay sheet
    m_jobId = ctx.RunJob("life.Worker", "{}")
    ctx.BindInput m_jobId, sheet, GRID_ADDRESS
    sheet.Range("W2").Value2 = "Generation"
    sheet.Range("W3").Value2 = 0
    ctx.LogMessage "life worker started"
End Sub

Public Sub Worker(ByVal job As InfraJob)
    Dim inputValues As Variant
    Dim outputValues(1 To GRID_SIZE, 1 To GRID_SIZE + 1) As Variant
    Dim rowNumber As Long
    Dim columnNumber As Long
    Dim neighbors As Long
    Dim alive As Boolean

    inputValues = job.ReadInput("A1:T20")
    ' Input versions represent the initial board or a real user edit. Keep
    ' subsequent generations in the worker so a delayed FE display cannot
    ' feed an old board back into the simulation.
    If Not IsEmpty(inputValues) Then
        If job.InputVersion > m_lastInputVersion Then
            m_board = inputValues
            m_lastInputVersion = job.InputVersion
        End If
    End If
    If IsEmpty(m_board) Then Exit Sub
    For rowNumber = 1 To GRID_SIZE
        For columnNumber = 1 To GRID_SIZE
            neighbors = NeighborCount(m_board, rowNumber, columnNumber)
            alive = (CellAlive(m_board, rowNumber, columnNumber))
            If alive Then
                outputValues(rowNumber, columnNumber) = IIf(neighbors = 2 Or neighbors = 3, 1, 0)
            Else
                outputValues(rowNumber, columnNumber) = IIf(neighbors = 3, 1, 0)
            End If
        Next columnNumber
    Next rowNumber
    outputValues(1, GRID_SIZE + 1) = m_generation + 1
    ' A rejected push means the FE is busy: skip this generation, keep running.
    If job.Push("board", "A1:U20", outputValues) Then
        m_generation = m_generation + 1
        m_board = outputValues
    End If
End Sub

Public Sub OnResult(ByVal ctx As InfraContext, ByVal jobId As String, ByVal version As Long)
    Dim values As Variant
    Dim board(1 To GRID_SIZE, 1 To GRID_SIZE) As Variant
    Dim rowNumber As Long
    Dim columnNumber As Long
    Dim sheet As Worksheet
    If jobId <> m_jobId Then Exit Sub
    values = ctx.ReadJob(jobId, "A1:U20")
    For rowNumber = 1 To GRID_SIZE
        For columnNumber = 1 To GRID_SIZE
            board(rowNumber, columnNumber) = values(rowNumber, columnNumber)
        Next columnNumber
    Next rowNumber
    Set sheet = ctx.Target.Worksheets(DISPLAY_SHEET)
    sheet.Range(GRID_ADDRESS).Value2 = board
    sheet.Range("W3").Value2 = CLng(values(1, GRID_SIZE + 1))
End Sub

Private Function CellAlive(ByVal values As Variant, ByVal rowNumber As Long, ByVal columnNumber As Long) As Boolean
    If rowNumber < 1 Or rowNumber > GRID_SIZE Then Exit Function
    If columnNumber < 1 Or columnNumber > GRID_SIZE Then Exit Function
    If IsNumeric(values(rowNumber, columnNumber)) Then CellAlive = (CDbl(values(rowNumber, columnNumber)) <> 0#)
End Function

Private Function NeighborCount(ByVal values As Variant, ByVal rowNumber As Long, ByVal columnNumber As Long) As Long
    Dim rowOffset As Long
    Dim columnOffset As Long
    For rowOffset = -1 To 1
        For columnOffset = -1 To 1
            If Not (rowOffset = 0 And columnOffset = 0) Then
                If CellAlive(values, rowNumber + rowOffset, columnNumber + columnOffset) Then
                    NeighborCount = NeighborCount + 1
                End If
            End If
        Next columnOffset
    Next rowOffset
End Function

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
    Dim grid As Range
    sheet.Cells.Clear
    sheet.Range("B1:W1").Merge
    sheet.Range("B1").Value2 = "CONWAY LIFE"
    sheet.Range("B1").Font.Bold = True
    sheet.Range("B1").Font.Size = 20
    sheet.Range("B1").HorizontalAlignment = xlCenter
    sheet.Range("B1").Interior.Color = RGB(23, 57, 53)
    sheet.Range("B1").Font.Color = RGB(255, 255, 255)
    Set grid = sheet.Range(GRID_ADDRESS)
    grid.ClearContents
    grid.Value2 = 0
    grid.NumberFormat = ";;;"
    grid.ColumnWidth = 2.4
    grid.RowHeight = 16
    grid.Borders.LineStyle = xlContinuous
    grid.Borders.Color = RGB(205, 215, 212)
    grid.FormatConditions.Delete
    With grid.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=1")
        .Interior.Color = RGB(70, 190, 145)
    End With
    sheet.Range("J12:L12").Value2 = 1
    sheet.Columns("W").ColumnWidth = 14
End Sub
