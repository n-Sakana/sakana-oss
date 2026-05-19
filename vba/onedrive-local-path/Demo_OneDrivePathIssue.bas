Attribute VB_Name = "Demo_OneDrivePathIssue"
Option Explicit

' Demonstrates why FSO/Dir operations fail when ThisWorkbook.Path is a
' SharePoint/OneDrive URL, and how convURLtoLocalPath fixes the base path.
'
' Import this module together with ConvURLtoLocalPath.bas, then run:
'     Demo_PathIssue_Compare

Private Const DEMO_FOLDER_NAME As String = "_vba_path_demo"
Private Const DEMO_FILE_NAME As String = "path-demo.txt"

Public Sub Demo_PathIssue_Compare()
    Dim ws As Worksheet
    Set ws = PrepareDemoSheet()

    Dim rawBase As String
    Dim fixedBase As String

    rawBase = ThisWorkbook.Path
    fixedBase = convURLtoLocalPath(ThisWorkbook, returnInputOnFail:=False)

    WriteHeader ws, rawBase, fixedBase

    RunFsoDemo ws, 6, "Raw ThisWorkbook.Path", rawBase
    RunFsoDemo ws, 7, "convURLtoLocalPath(ThisWorkbook)", fixedBase

    FormatDemoSheet ws
    ws.Activate
End Sub

Private Function PrepareDemoSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Path Demo")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "Path Demo"
    Else
        ws.Cells.Clear
    End If

    Set PrepareDemoSheet = ws
End Function

Private Sub WriteHeader(ByVal ws As Worksheet, ByVal rawBase As String, ByVal fixedBase As String)
    ws.Range("A1").Value = "OneDrive / SharePoint Path Demo"
    ws.Range("A3").Value = "ThisWorkbook.Name"
    ws.Range("B3").Value = ThisWorkbook.Name
    ws.Range("A4").Value = "ThisWorkbook.Path"
    ws.Range("B4").Value = rawBase
    ws.Range("A5").Value = "convURLtoLocalPath(ThisWorkbook)"
    ws.Range("B5").Value = fixedBase

    ws.Range("A6").Value = "Case"
    ws.Range("B6").Value = "Base path"
    ws.Range("C6").Value = "FolderExists"
    ws.Range("D6").Value = "CreateFolder"
    ws.Range("E6").Value = "CreateTextFile"
    ws.Range("F6").Value = "Dir(file)"
    ws.Range("G6").Value = "Result"
    ws.Range("H6").Value = "Error"
End Sub

Private Sub RunFsoDemo(ByVal ws As Worksheet, ByVal rowIndex As Long, _
                       ByVal caseName As String, ByVal basePath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim demoFolder As String
    Dim demoFile As String

    demoFolder = JoinPath(basePath, DEMO_FOLDER_NAME)
    demoFile = JoinPath(demoFolder, DEMO_FILE_NAME)

    ws.Cells(rowIndex, 1).Value = caseName
    ws.Cells(rowIndex, 2).Value = basePath

    If Len(basePath) = 0 Then
        WriteDemoResult ws, rowIndex, False, "Base path is empty"
        Exit Sub
    End If

    Dim ok As Boolean
    Dim errorText As String

    On Error Resume Next
    ok = fso.FolderExists(basePath)
    errorText = ErrorInfo()
    On Error GoTo 0
    ws.Cells(rowIndex, 3).Value = StatusText(ok, errorText)
    If Len(errorText) > 0 Then AppendError ws, rowIndex, "FolderExists: " & errorText

    On Error Resume Next
    If Not fso.FolderExists(demoFolder) Then fso.CreateFolder demoFolder
    ok = (Err.Number = 0)
    errorText = ErrorInfo()
    On Error GoTo 0
    ws.Cells(rowIndex, 4).Value = StatusText(ok, errorText)
    If Len(errorText) > 0 Then AppendError ws, rowIndex, "CreateFolder: " & errorText

    On Error Resume Next
    Dim textFile As Object
    Set textFile = fso.CreateTextFile(demoFile, True, True)
    If Err.Number = 0 Then
        textFile.WriteLine "Created by Demo_PathIssue_Compare"
        textFile.WriteLine "Case: " & caseName
        textFile.WriteLine "CreatedAt: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
        textFile.WriteLine "ThisWorkbook.Name: " & ThisWorkbook.Name
        textFile.WriteLine "ThisWorkbook.Path: " & ThisWorkbook.Path
        textFile.WriteLine "ResolvedBasePath: " & basePath
        textFile.Close
    End If
    ok = (Err.Number = 0)
    errorText = ErrorInfo()
    On Error GoTo 0
    ws.Cells(rowIndex, 5).Value = StatusText(ok, errorText)
    If Len(errorText) > 0 Then AppendError ws, rowIndex, "CreateTextFile: " & errorText

    On Error Resume Next
    Dim dirResult As String
    dirResult = Dir(demoFile)
    ok = (Err.Number = 0 And Len(dirResult) > 0)
    errorText = ErrorInfo()
    On Error GoTo 0
    ws.Cells(rowIndex, 6).Value = StatusText(ok, errorText)
    If Len(errorText) > 0 Then AppendError ws, rowIndex, "Dir: " & errorText

    WriteDemoResult ws, rowIndex, _
                    (ws.Cells(rowIndex, 4).Value = "OK" _
                     And ws.Cells(rowIndex, 5).Value = "OK" _
                     And ws.Cells(rowIndex, 6).Value = "OK"), _
                    vbNullString
End Sub

Private Sub WriteDemoResult(ByVal ws As Worksheet, ByVal rowIndex As Long, _
                            ByVal success As Boolean, ByVal message As String)
    If success Then
        ws.Cells(rowIndex, 7).Value = "SUCCESS"
    Else
        ws.Cells(rowIndex, 7).Value = "FAILED"
    End If

    If Len(message) > 0 Then AppendError ws, rowIndex, message
End Sub

Private Function ErrorInfo() As String
    If Err.Number <> 0 Then
        ErrorInfo = CStr(Err.Number) & ": " & Err.Description
        Err.Clear
    End If
End Function

Private Function StatusText(ByVal ok As Boolean, ByVal errorText As String) As String
    If ok And Len(errorText) = 0 Then
        StatusText = "OK"
    Else
        StatusText = "NG"
    End If
End Function

Private Sub AppendError(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal message As String)
    If Len(ws.Cells(rowIndex, 8).Value) > 0 Then
        ws.Cells(rowIndex, 8).Value = ws.Cells(rowIndex, 8).Value & vbLf & message
    Else
        ws.Cells(rowIndex, 8).Value = message
    End If
End Sub

Private Function JoinPath(ByVal basePath As String, ByVal childName As String) As String
    If Len(basePath) = 0 Then
        JoinPath = childName
    ElseIf Right$(basePath, 1) = "\" Or Right$(basePath, 1) = "/" Then
        JoinPath = basePath & childName
    Else
        JoinPath = basePath & "\" & childName
    End If
End Function

Private Sub FormatDemoSheet(ByVal ws As Worksheet)
    With ws
        .Columns("A:H").EntireColumn.AutoFit
        .Columns("B:B").ColumnWidth = 70
        .Columns("H:H").ColumnWidth = 45
        .Range("A1:H1").Font.Bold = True
        .Range("A6:H6").Font.Bold = True
        .Range("A6:H8").Borders.LineStyle = xlContinuous
        .Range("B4:B5").WrapText = True
        .Range("B7:B8").WrapText = True
        .Range("H7:H8").WrapText = True
        .Range("G7:G8").Font.Bold = True
        .Rows("1:8").VerticalAlignment = xlTop
    End With
End Sub
