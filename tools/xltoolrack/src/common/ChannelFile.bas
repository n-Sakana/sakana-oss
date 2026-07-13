Attribute VB_Name = "ChannelFile"
Option Explicit
Option Private Module

' File protocol shared by the worker side (publish results, read input) and
' the FE side (publish input, read results). Writers publish atomically by
' writing "<path>.tmp" and renaming it over the target; a reader that finds
' the file missing or unreadable skips that cycle and retries on the next.
' Cell encoding, one row per line, tab separated: N<digits> for numerics
' (Str$/Val, locale independent), S<text> for strings, E for empty.

Public Function PayloadPath(ByVal channelBase As String) As String
    PayloadPath = channelBase & "_payload.dat"
End Function

Public Function InputPath(ByVal channelBase As String) As String
    InputPath = channelBase & "_input.dat"
End Function

Public Function StopPath(ByVal channelBase As String) As String
    StopPath = channelBase & "_stop.flag"
End Function

Public Function ErrorPath(ByVal channelBase As String) As String
    ErrorPath = channelBase & "_error.dat"
End Function

Public Function AggregatePayloadPath(ByVal channelBase As String) As String
    Dim separatorAt As Long
    separatorAt = InStrRev(channelBase, "_")
    If separatorAt = 0 Then
        AggregatePayloadPath = channelBase & "_payloads.dat"
    Else
        AggregatePayloadPath = Left$(channelBase, separatorAt - 1) & "_payloads.dat"
    End If
End Function

Public Function FlagExists(ByVal path As String) As Boolean
    If Len(path) = 0 Then Exit Function
    On Error Resume Next
    FlagExists = (Len(Dir$(path)) > 0)
    On Error GoTo 0
End Function

Public Sub WriteFlag(ByVal path As String, ByVal text As String)
    Dim fileNo As Integer
    On Error Resume Next
    fileNo = FreeFile
    Open path For Output As #fileNo
    Print #fileNo, text
    Close #fileNo
    On Error GoTo 0
End Sub

Public Sub DeleteQuiet(ByVal path As String)
    If Len(path) = 0 Then Exit Sub
    On Error Resume Next
    If Len(Dir$(path)) > 0 Then Kill path
    On Error GoTo 0
End Sub

Public Function WriteRect(ByVal path As String, ByVal version As Long, ByVal address As String, ByVal values As Variant) As Boolean
    Dim fileNo As Integer
    Dim tmpPath As String
    Dim rowCount As Long, columnCount As Long
    Dim r As Long, c As Long
    Dim lineText As String
    Dim wrote As Boolean

    tmpPath = path & ".tmp"
    If IsArray(values) Then
        rowCount = UBound(values, 1) - LBound(values, 1) + 1
        columnCount = UBound(values, 2) - LBound(values, 2) + 1
    Else
        rowCount = 1
        columnCount = 1
    End If

    On Error GoTo Failed
    fileNo = FreeFile
    Open tmpPath For Output As #fileNo
    Print #fileNo, "v" & vbTab & CStr(version) & vbTab & address & vbTab & CStr(rowCount) & vbTab & CStr(columnCount)
    If IsArray(values) Then
        For r = LBound(values, 1) To UBound(values, 1)
            lineText = vbNullString
            For c = LBound(values, 2) To UBound(values, 2)
                If c > LBound(values, 2) Then lineText = lineText & vbTab
                lineText = lineText & EncodeCell(values(r, c))
            Next c
            Print #fileNo, lineText
        Next r
    Else
        Print #fileNo, EncodeCell(values)
    End If
    Close #fileNo
    fileNo = 0
    wrote = True

    On Error Resume Next
    If Len(Dir$(path)) > 0 Then Kill path
    Err.Clear
    Name tmpPath As path
    If Err.Number <> 0 Then
        Kill tmpPath
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    WriteRect = True
    Exit Function

Failed:
    On Error Resume Next
    If fileNo <> 0 Then Close #fileNo
    If Not wrote Then Kill tmpPath
    On Error GoTo 0
End Function

' All workers in one FE session publish into one latest-value snapshot. A
' dedicated lock serializes the worker-side read/replace cycle; the FE opens
' only the resulting aggregate file once per pump tick.
Public Function WriteAggregateRecord(ByVal channelBase As String, ByVal jobId As String, _
                                     ByVal version As Long, ByVal address As String, _
                                     ByVal values As Variant) As Boolean
    Dim path As String
    Dim lockPath As String
    Dim tmpPath As String
    Dim lockNo As Integer
    Dim inputNo As Integer
    Dim outputNo As Integer
    Dim lines As Collection
    Dim lineText As String
    Dim parts As Variant
    Dim i As Long
    Dim lockHeld As Boolean

    path = AggregatePayloadPath(channelBase)
    lockPath = path & ".lock"
    tmpPath = path & ".tmp"
    Set lines = New Collection
    On Error GoTo Failed

    lockNo = FreeFile
    Open lockPath For Binary Access Read Write Lock Read Write As #lockNo
    lockHeld = True

    If Len(Dir$(path)) > 0 Then
        inputNo = FreeFile
        Open path For Input Shared As #inputNo
        Do Until EOF(inputNo)
            Line Input #inputNo, lineText
            parts = Split(lineText, vbTab)
            If UBound(parts) < 1 Then
                lines.Add lineText
            ElseIf CStr(parts(0)) <> "r" Or CStr(parts(1)) <> jobId Then
                lines.Add lineText
            End If
        Loop
        Close #inputNo
        inputNo = 0
    End If

    outputNo = FreeFile
    Open tmpPath For Output As #outputNo
    For i = 1 To lines.Count
        Print #outputNo, CStr(lines(i))
    Next i
    Print #outputNo, EncodeAggregateRecord(jobId, version, address, values)
    Close #outputNo
    outputNo = 0

    If Len(Dir$(path)) > 0 Then Kill path
    Name tmpPath As path
    Close #lockNo
    lockNo = 0
    WriteAggregateRecord = True
    Exit Function

Failed:
    On Error Resume Next
    If inputNo <> 0 Then Close #inputNo
    If outputNo <> 0 Then Close #outputNo
    If lockNo <> 0 Then Close #lockNo
    If lockHeld Then
        If Len(Dir$(tmpPath)) > 0 Then Kill tmpPath
    End If
    On Error GoTo 0
End Function

Public Function ReadAggregate(ByVal channelBase As String, _
                              Optional ByRef profileExistsMs As Double = 0#, _
                              Optional ByRef profileOpenMs As Double = 0#, _
                              Optional ByRef profileLineInputMs As Double = 0#, _
                              Optional ByRef profileDecodeMs As Double = 0#, _
                              Optional ByRef profileCloseMs As Double = 0#, _
                              Optional ByVal profileRequested As Boolean = False) As Collection
    Dim path As String
    Dim fileNo As Integer
    Dim lineText As String
    Dim records As Collection
    Dim record As ChannelRecord
    Dim lines As Collection
    Dim rawLine As Variant
    Dim pathExists As Boolean
    Dim stageStarted As Double
    Set records = New Collection
    path = AggregatePayloadPath(channelBase)
    If profileRequested Then
        stageStarted = ProfileClockSeconds()
        pathExists = FlagExists(path)
        profileExistsMs = ElapsedMilliseconds(stageStarted, ProfileClockSeconds())
    Else
        pathExists = FlagExists(path)
    End If
    If Not pathExists Then
        Set ReadAggregate = records
        Exit Function
    End If
    On Error GoTo Failed
    If profileRequested Then
        stageStarted = ProfileClockSeconds()
        fileNo = FreeFile
        Open path For Input Shared As #fileNo
        profileOpenMs = ElapsedMilliseconds(stageStarted, ProfileClockSeconds())
        Set lines = New Collection
        ' Time each stage as one block. The normal read path above is unchanged,
        ' and profiling avoids per-record clock calls.
        stageStarted = ProfileClockSeconds()
        Do Until EOF(fileNo)
            Line Input #fileNo, lineText
            lines.Add lineText
        Loop
        profileLineInputMs = ElapsedMilliseconds(stageStarted, ProfileClockSeconds())
        stageStarted = ProfileClockSeconds()
        For Each rawLine In lines
            Set record = DecodeAggregateRecord(CStr(rawLine))
            If Not record Is Nothing Then records.Add record
        Next rawLine
        profileDecodeMs = ElapsedMilliseconds(stageStarted, ProfileClockSeconds())
    Else
        fileNo = FreeFile
        Open path For Input Shared As #fileNo
        Do Until EOF(fileNo)
            Line Input #fileNo, lineText
            Set record = DecodeAggregateRecord(lineText)
            If Not record Is Nothing Then records.Add record
        Loop
    End If
    If profileRequested Then stageStarted = ProfileClockSeconds()
    Close #fileNo
    If profileRequested Then profileCloseMs = ElapsedMilliseconds(stageStarted, ProfileClockSeconds())
    fileNo = 0
Failed:
    On Error Resume Next
    If fileNo <> 0 Then Close #fileNo
    On Error GoTo 0
    Set ReadAggregate = records
End Function

Public Sub DeleteAggregateRecord(ByVal channelBase As String, ByVal jobId As String)
    Dim path As String
    Dim lockPath As String
    Dim tmpPath As String
    Dim lockNo As Integer
    Dim inputNo As Integer
    Dim outputNo As Integer
    Dim lines As Collection
    Dim lineText As String
    Dim parts As Variant
    Dim i As Long
    Dim lockHeld As Boolean

    path = AggregatePayloadPath(channelBase)
    lockPath = path & ".lock"
    tmpPath = path & ".tmp"
    Set lines = New Collection
    On Error GoTo Failed
    lockNo = FreeFile
    Open lockPath For Binary Access Read Write Lock Read Write As #lockNo
    lockHeld = True
    If Len(Dir$(path)) = 0 Then GoTo CleanExit

    inputNo = FreeFile
    Open path For Input Shared As #inputNo
    Do Until EOF(inputNo)
        Line Input #inputNo, lineText
        parts = Split(lineText, vbTab)
        If UBound(parts) < 1 Then
            lines.Add lineText
        ElseIf CStr(parts(0)) <> "r" Or CStr(parts(1)) <> jobId Then
            lines.Add lineText
        End If
    Loop
    Close #inputNo
    inputNo = 0

    If lines.Count = 0 Then
        Kill path
    Else
        outputNo = FreeFile
        Open tmpPath For Output As #outputNo
        For i = 1 To lines.Count
            Print #outputNo, CStr(lines(i))
        Next i
        Close #outputNo
        outputNo = 0
        Kill path
        Name tmpPath As path
    End If

CleanExit:
    If lockNo <> 0 Then Close #lockNo
    Exit Sub
Failed:
    On Error Resume Next
    If inputNo <> 0 Then Close #inputNo
    If outputNo <> 0 Then Close #outputNo
    If lockNo <> 0 Then Close #lockNo
    If lockHeld Then
        If Len(Dir$(tmpPath)) > 0 Then Kill tmpPath
    End If
    On Error GoTo 0
End Sub

Public Function ReadRect(ByVal path As String, ByRef version As Long, ByRef address As String, ByRef values As Variant) As Boolean
    Dim fileNo As Integer
    Dim header As String
    Dim parts As Variant
    Dim rowCount As Long, columnCount As Long
    Dim r As Long, c As Long
    Dim lineText As String
    Dim cells As Variant
    Dim result As Variant

    If Not FlagExists(path) Then Exit Function
    On Error GoTo Failed
    fileNo = FreeFile
    Open path For Input Shared As #fileNo
    Line Input #fileNo, header
    parts = Split(header, vbTab)
    If UBound(parts) < 4 Then GoTo Failed
    If CStr(parts(0)) <> "v" Then GoTo Failed
    version = CLng(Val(parts(1)))
    address = CStr(parts(2))
    rowCount = CLng(Val(parts(3)))
    columnCount = CLng(Val(parts(4)))
    If rowCount < 1 Or columnCount < 1 Or rowCount > 10000 Or columnCount > 1000 Then GoTo Failed
    ReDim result(1 To rowCount, 1 To columnCount)
    For r = 1 To rowCount
        Line Input #fileNo, lineText
        cells = Split(lineText, vbTab)
        If UBound(cells) < columnCount - 1 Then GoTo Failed
        For c = 1 To columnCount
            result(r, c) = DecodeCell(CStr(cells(c - 1)))
        Next c
    Next r
    Close #fileNo
    values = result
    ReadRect = True
    Exit Function

Failed:
    On Error Resume Next
    If fileNo <> 0 Then Close #fileNo
    On Error GoTo 0
End Function

Public Sub WriteError(ByVal path As String, ByVal number As Long, ByVal description As String)
    Dim safeText As String
    safeText = Replace(Replace(Replace(description, vbTab, " "), vbCr, " "), vbLf, " ")
    Call WriteRect(path, 1, "error", Array2D(CStr(number), safeText))
End Sub

Public Function ReadError(ByVal path As String, ByRef number As Long, ByRef description As String) As Boolean
    Dim version As Long
    Dim address As String
    Dim values As Variant
    If Not ReadRect(path, version, address, values) Then Exit Function
    number = CLng(Val(CStr(values(1, 1))))
    description = CStr(values(1, 2))
    ReadError = True
End Function

Private Function Array2D(ByVal first As String, ByVal second As String) As Variant
    Dim result(1 To 1, 1 To 2) As Variant
    result(1, 1) = first
    result(1, 2) = second
    Array2D = result
End Function

Private Function EncodeAggregateRecord(ByVal jobId As String, ByVal version As Long, _
                                       ByVal address As String, ByVal values As Variant) As String
    Dim rowCount As Long
    Dim columnCount As Long
    Dim r As Long
    Dim c As Long
    Dim result As String
    If IsArray(values) Then
        rowCount = UBound(values, 1) - LBound(values, 1) + 1
        columnCount = UBound(values, 2) - LBound(values, 2) + 1
    Else
        rowCount = 1
        columnCount = 1
    End If
    result = "r" & vbTab & jobId & vbTab & CStr(version) & vbTab & address & vbTab & _
             CStr(rowCount) & vbTab & CStr(columnCount)
    If IsArray(values) Then
        For r = LBound(values, 1) To UBound(values, 1)
            For c = LBound(values, 2) To UBound(values, 2)
                result = result & vbTab & EncodeCell(values(r, c))
            Next c
        Next r
    Else
        result = result & vbTab & EncodeCell(values)
    End If
    EncodeAggregateRecord = result
End Function

Private Function DecodeAggregateRecord(ByVal lineText As String) As ChannelRecord
    Dim parts As Variant
    Dim rowCount As Long
    Dim columnCount As Long
    Dim values As Variant
    Dim r As Long
    Dim c As Long
    Dim fieldAt As Long
    Dim record As ChannelRecord
    On Error GoTo Failed
    parts = Split(lineText, vbTab)
    If UBound(parts) < 6 Then Exit Function
    If CStr(parts(0)) <> "r" Then Exit Function
    rowCount = CLng(Val(parts(4)))
    columnCount = CLng(Val(parts(5)))
    If rowCount < 1 Or columnCount < 1 Or rowCount > 10000 Or columnCount > 1000 Then Exit Function
    If UBound(parts) < 5 + rowCount * columnCount Then Exit Function
    ReDim values(1 To rowCount, 1 To columnCount)
    fieldAt = 6
    For r = 1 To rowCount
        For c = 1 To columnCount
            values(r, c) = DecodeCell(CStr(parts(fieldAt)))
            fieldAt = fieldAt + 1
        Next c
    Next r
    Set record = New ChannelRecord
    record.JobId = CStr(parts(1))
    record.Version = CLng(Val(parts(2)))
    record.Address = CStr(parts(3))
    record.Values = values
    Set DecodeAggregateRecord = record
    Exit Function
Failed:
End Function

Private Function EncodeCell(ByVal value As Variant) As String
    Select Case VarType(value)
        Case vbEmpty, vbNull
            EncodeCell = "E"
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbByte, vbDecimal
            EncodeCell = "N" & Trim$(Str$(CDbl(value)))
        Case vbDate
            EncodeCell = "N" & Trim$(Str$(CDbl(value)))
        Case vbBoolean
            EncodeCell = "N" & Trim$(Str$(CDbl(value)))
        Case Else
            EncodeCell = "S" & Replace(Replace(Replace(CStr(value), vbTab, " "), vbCr, " "), vbLf, " ")
    End Select
End Function

Private Function DecodeCell(ByVal token As String) As Variant
    If Len(token) = 0 Then
        DecodeCell = Empty
    ElseIf Left$(token, 1) = "N" Then
        DecodeCell = Val(Mid$(token, 2))
    ElseIf Left$(token, 1) = "S" Then
        DecodeCell = Mid$(token, 2)
    Else
        DecodeCell = Empty
    End If
End Function

Private Function ElapsedMilliseconds(ByVal startedAt As Double, ByVal endedAt As Double) As Double
    If endedAt < startedAt Then endedAt = endedAt + 86400#
    ElapsedMilliseconds = (endedAt - startedAt) * 1000#
End Function

Private Function ProfileClockSeconds() As Double
    ProfileClockSeconds = Timer
End Function
