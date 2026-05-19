Attribute VB_Name = "modConvURLtoLocalPathDebug"
Option Explicit

' Diagnostic module for the lightweight ConvURLtoLocalPath.bas module.
' Import this together with ConvURLtoLocalPath.bas, then run:
'     DebugThisWorkbookLocalPath

Private Const DEBUG_SHEET_NAME As String = "OneDrive Path Debug"
Private Const ADDED_SCOPE_SHEET_NAME As String = "OneDrive AddedScope Debug"
Private Const PATH_SEP As String = "\"

Public Sub DebugThisWorkbookLocalPath()
    convURLtoLocalPathDebug ThisWorkbook, True
    MsgBox "Diagnostic log was written." & vbCrLf & _
           "Sheet: " & DEBUG_SHEET_NAME & vbCrLf & _
           "Immediate: press Ctrl+G in the VBE", vbInformation
End Sub

Public Sub DebugThisWorkbookLocalPathLight()
    DebugThisWorkbookLocalPath
End Sub

Public Sub convURLtoLocalPathDebug(Optional ByVal source As Variant, _
                                   Optional ByVal writeToSheet As Boolean = True)
    Dim rows As New Collection
    On Error GoTo Failed

    Dim inputPath As String
    If IsMissing(source) Then
        inputPath = InputPathFromSource(ThisWorkbook)
    Else
        inputPath = InputPathFromSource(source)
    End If

    AddDebugRow rows, "Input", "Resolved input", inputPath
    AddDebugRow rows, "Input", "ThisWorkbook.Path", SafeWorkbookProperty("Path")
    AddDebugRow rows, "Input", "ThisWorkbook.FullName", SafeWorkbookProperty("FullName")
    AddDebugRow rows, "Input", "Is HTTPS URL", CStr(IsHttpsUrl(inputPath))
    AddDebugRow rows, "Input", "Use SyncEngineDatabase.db", "False"

    Dim normalizedUrl As String
    If IsHttpsUrl(inputPath) Then
        normalizedUrl = NormalizeUrl(inputPath)
        AddDebugRow rows, "Input", "Normalized URL", normalizedUrl
    End If

    AddSettingsRows rows

    Dim providers As Collection
    Set providers = convURLtoLocalPathProviderMap(True)
    AddDebugRow rows, "Providers", "Total providers", CStr(providers.Count)
    AddProviderRows rows, providers, normalizedUrl

    If Len(normalizedUrl) > 0 Then
        AddDebugRow rows, "Resolve", "requireExists=True", _
                    convURLtoLocalPath(inputPath, True, False, True)
        AddDebugRow rows, "Resolve", "requireExists=False", _
                    convURLtoLocalPath(inputPath, False, False, False)
        AddDebugRow rows, "Resolve", "Public function result", _
                    convURLtoLocalPath(inputPath, True, False, True)
    End If

    GoTo Done

Failed:
    AddDebugRow rows, "Debug", "ERROR", Err.Number & ": " & Err.Description

Done:
    On Error Resume Next
    PrintDebugRows rows
    If writeToSheet Then WriteDebugRowsToSheet rows
    On Error GoTo 0
End Sub

Public Sub DebugOneDriveAddedScopes()
    Dim rows As New Collection
    On Error GoTo Failed

    AddDebugRow rows, "Env", "OneDrive", Environ$("OneDrive"), _
                "exists=" & CStr(FolderExists(Environ$("OneDrive")))
    AddDebugRow rows, "Env", "OneDriveCommercial", Environ$("OneDriveCommercial"), _
                "exists=" & CStr(FolderExists(Environ$("OneDriveCommercial")))
    AddDebugRow rows, "Env", "OneDriveConsumer", Environ$("OneDriveConsumer"), _
                "exists=" & CStr(FolderExists(Environ$("OneDriveConsumer")))

    AddAddedScopeRows rows

    GoTo Done

Failed:
    AddDebugRow rows, "Debug", "ERROR", Err.Number & ": " & Err.Description

Done:
    On Error Resume Next
    PrintDebugRows rows
    WriteRowsToSheet rows, ADDED_SCOPE_SHEET_NAME
    MsgBox "AddedScope diagnostic log was written." & vbCrLf & _
           "Sheet: " & ADDED_SCOPE_SHEET_NAME, vbInformation
    On Error GoTo 0
End Sub

Private Function InputPathFromSource(ByVal source As Variant) As String
    On Error GoTo FromValue

    If IsObject(source) Then
        InputPathFromSource = CStr(CallByName(source, "Path", VbGet))
        If Len(InputPathFromSource) > 0 Then Exit Function

        Err.Clear
        InputPathFromSource = CStr(CallByName(source, "FullName", VbGet))
        Exit Function
    End If

FromValue:
    On Error Resume Next
    InputPathFromSource = CStr(source)
    On Error GoTo 0
End Function

Private Function SafeWorkbookProperty(ByVal propertyName As String) As String
    On Error Resume Next
    SafeWorkbookProperty = CStr(CallByName(ThisWorkbook, propertyName, VbGet))
    If Err.Number <> 0 Then SafeWorkbookProperty = Err.Number & ": " & Err.Description
    Err.Clear
    On Error GoTo 0
End Function

Private Sub AddSettingsRows(ByVal rows As Collection)
    Dim settingsRoot As String
    settingsRoot = CombinePath(Environ$("LOCALAPPDATA"), "Microsoft\OneDrive\settings")

    AddDebugRow rows, "Settings", "Root", settingsRoot
    AddDebugRow rows, "Settings", "Root exists", CStr(FolderExists(settingsRoot))

    If Not FolderExists(settingsRoot) Then Exit Sub

    Dim folderName As String
    Dim accountCount As Long
    folderName = Dir(CombinePath(settingsRoot, "*"), vbDirectory)

    Do While Len(folderName) > 0
        If folderName <> "." And folderName <> ".." Then
            If LCase$(folderName) = "personal" _
               Or LCase$(Left$(folderName, 8)) = "business" Then
                accountCount = accountCount + 1
                AddDebugRow rows, "Settings", "Account folder " & CStr(accountCount), _
                            CombinePath(settingsRoot, folderName)
            End If
        End If
        folderName = Dir
    Loop

    AddDebugRow rows, "Settings", "Account folder count", CStr(accountCount)
End Sub

Private Sub AddAddedScopeRows(ByVal rows As Collection)
    Dim settingsRoot As String
    settingsRoot = CombinePath(Environ$("LOCALAPPDATA"), "Microsoft\OneDrive\settings")

    AddDebugRow rows, "Settings", "Root", settingsRoot
    AddDebugRow rows, "Settings", "Root exists", CStr(FolderExists(settingsRoot))
    If Not FolderExists(settingsRoot) Then Exit Sub

    Dim accountFolder As String
    Dim accountName As String
    Dim folderName As String
    Dim fileName As String
    Dim fileText As String
    Dim lineText As Variant
    Dim tokens As Variant
    Dim index As Long
    Dim accounts As New Collection

    folderName = Dir(CombinePath(settingsRoot, "*"), vbDirectory)
    Do While Len(folderName) > 0
        If folderName <> "." And folderName <> ".." Then
            If LCase$(folderName) = "personal" _
               Or LCase$(Left$(folderName, 8)) = "business" Then
                accounts.Add Array(folderName, CombinePath(settingsRoot, folderName))
            End If
        End If
        folderName = Dir
    Loop

    Dim account As Variant
    For Each account In accounts
        accountName = CStr(account(0))
        accountFolder = CStr(account(1))
        AddDebugRow rows, "Account", accountName, accountFolder

        fileName = Dir(CombinePath(accountFolder, "*.ini"))
        Do While Len(fileName) > 0
            fileText = ReadTextFile(CombinePath(accountFolder, fileName))

            If InStr(1, fileText, "AddedScope = ", vbBinaryCompare) > 0 Then
                For Each lineText In SplitLines(fileText)
                    If Left$(CStr(lineText), 13) = "AddedScope = " Then
                        index = index + 1
                        tokens = ParseSettingLine(CStr(lineText))
                        AddOneAddedScopeRow rows, accountName, fileName, index, tokens
                    End If
                Next lineText
            End If

            fileName = Dir
        Loop
    Next account

    AddDebugRow rows, "AddedScope", "Total", CStr(index)
End Sub

Private Sub AddOneAddedScopeRow(ByVal rows As Collection, _
                                ByVal accountName As String, _
                                ByVal fileName As String, _
                                ByVal index As Long, _
                                ByVal tokens As Variant)
    Dim directUrl As String
    Dim folderId As String
    Dim relPath As String
    Dim firstLocal As String
    Dim tokenCount As String

    directUrl = DirectUrlFromTokens(tokens)
    folderId = TokenAt(tokens, 3)
    relPath = RelativePathFromToken(TokenAt(tokens, 11), PATH_SEP)
    firstLocal = FirstLocalPath(tokens)
    tokenCount = CStr(TokenCount(tokens))

    AddDebugRow rows, "AddedScope " & CStr(index), "source", _
                accountName & "\" & fileName, "tokens=" & tokenCount
    AddDebugRow rows, "AddedScope " & CStr(index), "directUrl", directUrl
    AddDebugRow rows, "AddedScope " & CStr(index), "folderId", folderId
    AddDebugRow rows, "AddedScope " & CStr(index), "relPath token 11", relPath
    AddDebugRow rows, "AddedScope " & CStr(index), "first local token", firstLocal, _
                "exists=" & CStr(PathExists(firstLocal))

    AddOneDriveRootCandidate rows, index, "OneDriveCommercial", relPath
    AddOneDriveRootCandidate rows, index, "OneDrive", relPath
    AddOneDriveRootCandidate rows, index, "OneDriveConsumer", relPath
End Sub

Private Sub AddOneDriveRootCandidate(ByVal rows As Collection, _
                                     ByVal index As Long, _
                                     ByVal envName As String, _
                                     ByVal relPath As String)
    Dim rootPath As String
    Dim candidate As String

    rootPath = Environ$(envName)
    If Len(rootPath) = 0 Then Exit Sub
    If Len(relPath) = 0 Then Exit Sub

    candidate = CombinePath(rootPath, relPath)
    AddDebugRow rows, "AddedScope " & CStr(index), _
                envName & " + relPath", candidate, _
                "exists=" & CStr(PathExists(candidate))
End Sub

Private Sub AddProviderRows(ByVal rows As Collection, _
                            ByVal providers As Collection, _
                            ByVal normalizedUrl As String)
    If providers Is Nothing Then Exit Sub

    Dim item As Variant
    Dim index As Long
    Dim localRoot As String
    Dim webRoot As String
    Dim relPart As String
    Dim candidate As String
    Dim reducedCandidate As String
    Dim isMatch As Boolean

    For Each item In providers
        index = index + 1
        localRoot = CStr(item(0))
        webRoot = CStr(item(1))
        isMatch = Len(normalizedUrl) > 0 And UrlStartsWith(normalizedUrl, webRoot)

        AddDebugRow rows, "Provider " & CStr(index), "webRoot", webRoot, _
                    "matches input=" & CStr(isMatch)
        AddDebugRow rows, "Provider " & CStr(index), "localRoot", localRoot, _
                    "exists=" & CStr(FolderExists(localRoot))

        If isMatch Then
            relPart = Mid$(normalizedUrl, Len(webRoot) + 1)
            relPart = Replace(relPart, "/", PATH_SEP)
            candidate = CombinePath(localRoot, relPart)
            reducedCandidate = CandidateWithoutFirstPathSegment(localRoot, relPart)
            AddDebugRow rows, "Provider " & CStr(index), "candidate", candidate, _
                        "exists=" & CStr(PathExists(candidate))
            If Len(reducedCandidate) > 0 Then
                AddDebugRow rows, "Provider " & CStr(index), _
                            "candidate without first segment", reducedCandidate, _
                            "exists=" & CStr(PathExists(reducedCandidate))
            End If
        End If
    Next item
End Sub

Private Sub AddDebugRow(ByVal rows As Collection, _
                        ByVal stage As String, _
                        ByVal name As String, _
                        ByVal value As String, _
                        Optional ByVal note As String = "")
    rows.Add Array(stage, name, value, note)
End Sub

Private Sub PrintDebugRows(ByVal rows As Collection)
    Dim row As Variant
    For Each row In rows
        Debug.Print CStr(row(0)) & " | " & CStr(row(1)) & _
                    " | " & CStr(row(2)) & " | " & CStr(row(3))
    Next row
End Sub

Private Sub WriteDebugRowsToSheet(ByVal rows As Collection)
    WriteRowsToSheet rows, DEBUG_SHEET_NAME
End Sub

Private Sub WriteRowsToSheet(ByVal rows As Collection, ByVal sheetName As String)
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If

    ws.Range("A1:D1").Value = Array("Stage", "Name", "Value", "Note")

    Dim i As Long
    Dim row As Variant
    i = 2
    For Each row In rows
        ws.Cells(i, 1).Value = CStr(row(0))
        ws.Cells(i, 2).Value = CStr(row(1))
        ws.Cells(i, 3).Value = CStr(row(2))
        ws.Cells(i, 4).Value = CStr(row(3))
        i = i + 1
    Next row

    ws.Columns("A:D").AutoFit
End Sub

Private Function ReadTextFile(ByVal filePath As String) As String
    On Error GoTo CloseFile

    Dim fileNo As Long
    fileNo = FreeFile
    Open filePath For Binary Access Read As #fileNo

    Dim bytes() As Byte
    ReDim bytes(0 To LOF(fileNo) - 1)
    Get #fileNo, , bytes

    Close #fileNo
    fileNo = 0

    If BytesLookLikeUtf16Le(bytes) Then
        ReadTextFile = Mid$(bytes, 2)
    Else
        ReadTextFile = StrConv(bytes, vbUnicode)
    End If

CloseFile:
    On Error Resume Next
    If fileNo <> 0 Then Close #fileNo
    On Error GoTo 0
End Function

Private Function BytesLookLikeUtf16Le(ByRef bytes() As Byte) As Boolean
    On Error GoTo Done

    If UBound(bytes) >= 1 Then
        If bytes(0) = &HFF And bytes(1) = &HFE Then
            BytesLookLikeUtf16Le = True
            Exit Function
        End If
    End If

    Dim i As Long
    Dim oddZero As Long
    Dim pairs As Long

    For i = 1 To UBound(bytes) Step 2
        pairs = pairs + 1
        If bytes(i) = 0 Then oddZero = oddZero + 1
        If pairs >= 200 Then Exit For
    Next i

    BytesLookLikeUtf16Le = pairs > 0 And oddZero > pairs * 0.6
Done:
End Function

Private Function SplitLines(ByVal text As String) As Variant
    text = Replace(text, vbCrLf, vbLf)
    text = Replace(text, vbCr, vbLf)
    SplitLines = Split(text, vbLf)
End Function

Private Function ParseSettingLine(ByVal lineText As String) As Variant
    Dim values() As String
    ReDim values(0 To 0)

    Dim count As Long
    Dim i As Long
    Dim j As Long
    Dim ch As String
    Dim token As String

    i = 1
    Do While i <= Len(lineText)
        ch = Mid$(lineText, i, 1)

        Select Case ch
            Case " ", vbTab
                i = i + 1

            Case """"
                j = InStr(i + 1, lineText, """")
                If j = 0 Then j = Len(lineText) + 1
                token = Mid$(lineText, i + 1, j - i - 1)
                AppendToken values, count, token
                i = j + 1

            Case Else
                j = i
                Do While j <= Len(lineText)
                    ch = Mid$(lineText, j, 1)
                    If ch = " " Or ch = vbTab Then Exit Do
                    j = j + 1
                Loop
                token = Mid$(lineText, i, j - i)
                AppendToken values, count, token
                i = j + 1
        End Select
    Loop

    ParseSettingLine = values
End Function

Private Sub AppendToken(ByRef values() As String, ByRef count As Long, ByVal value As String)
    If count = 0 Then
        values(0) = value
    Else
        ReDim Preserve values(0 To count)
        values(count) = value
    End If
    count = count + 1
End Sub

Private Function TokenAt(ByVal tokens As Variant, ByVal index As Long) As String
    On Error GoTo Done
    If IsArray(tokens) Then
        If index >= LBound(tokens) And index <= UBound(tokens) Then
            TokenAt = CStr(tokens(index))
        End If
    End If
Done:
End Function

Private Function TokenCount(ByVal tokens As Variant) As Long
    On Error GoTo Done
    If IsArray(tokens) Then TokenCount = UBound(tokens) - LBound(tokens) + 1
Done:
End Function

Private Function FirstLocalPath(ByVal tokens As Variant) As String
    On Error GoTo Done
    Dim i As Long
    For i = LBound(tokens) To UBound(tokens)
        If LooksLikeLocalPath(CStr(tokens(i))) Then
            FirstLocalPath = CStr(tokens(i))
            Exit Function
        End If
    Next i
Done:
End Function

Private Function FirstHttpsUrl(ByVal tokens As Variant) As String
    On Error GoTo Done
    Dim i As Long
    For i = LBound(tokens) To UBound(tokens)
        If IsHttpsUrl(CStr(tokens(i))) Then
            FirstHttpsUrl = CStr(tokens(i))
            Exit Function
        End If
    Next i
Done:
End Function

Private Function DirectUrlFromTokens(ByVal tokens As Variant) As String
    Dim tagName As String
    tagName = TokenAt(tokens, 0)

    Select Case tagName
        Case "libraryScope"
            DirectUrlFromTokens = TokenAt(tokens, 8)
        Case "AddedScope"
            DirectUrlFromTokens = TokenAt(tokens, 5)
        Case Else
            DirectUrlFromTokens = FirstHttpsUrl(tokens)
    End Select

    If Not IsHttpsUrl(DirectUrlFromTokens) Then _
        DirectUrlFromTokens = FirstHttpsUrl(tokens)
End Function

Private Function RelativePathFromToken(ByVal value As String, _
                                       ByVal separator As String) As String
    value = Trim$(value)
    If Len(value) = 0 Or value = " " Then Exit Function
    If IsHttpsUrl(value) Then Exit Function

    value = Replace(value, "/", separator)
    If separator = PATH_SEP Then value = TrimLeadingBackslash(value)
    RelativePathFromToken = value
End Function

Private Function LooksLikeLocalPath(ByVal value As String) As Boolean
    value = Trim$(value)
    LooksLikeLocalPath = (Len(value) >= 3 And Mid$(value, 2, 2) = ":\") _
                         Or Left$(value, 2) = "\\"
End Function

Private Function IsHttpsUrl(ByVal value As String) As Boolean
    value = LCase$(Trim$(value))
    IsHttpsUrl = (Left$(value, 8) = "https://")
End Function

Private Function NormalizeUrl(ByVal value As String) As String
    NormalizeUrl = TrimTrailingSlash(UrlDecode(Replace(Trim$(value), "\", "/")))
End Function

Private Function UrlStartsWith(ByVal value As String, ByVal prefix As String) As Boolean
    value = TrimTrailingSlash(NormalizeUrl(value))
    prefix = TrimTrailingSlash(NormalizeUrl(prefix))

    UrlStartsWith = (StrComp(value, prefix, vbTextCompare) = 0) _
                    Or (Len(value) > Len(prefix) _
                        And StrComp(Left$(value, Len(prefix) + 1), prefix & "/", vbTextCompare) = 0)
End Function

Private Function UrlDecode(ByVal value As String) As String
    Dim result As String
    Dim bytes() As Byte
    Dim byteCount As Long
    Dim i As Long

    i = 1
    Do While i <= Len(value)
        If Mid$(value, i, 1) = "%" And i + 2 <= Len(value) _
           And IsHexPair(Mid$(value, i + 1, 2)) Then
            byteCount = 0
            ReDim bytes(0 To 0)

            Do While i + 2 <= Len(value) _
                 And Mid$(value, i, 1) = "%" _
                 And IsHexPair(Mid$(value, i + 1, 2))
                If byteCount > 0 Then ReDim Preserve bytes(0 To byteCount)
                bytes(byteCount) = CByte("&H" & Mid$(value, i + 1, 2))
                byteCount = byteCount + 1
                i = i + 3
            Loop

            result = result & Utf8BytesToString(bytes)
        Else
            result = result & Mid$(value, i, 1)
            i = i + 1
        End If
    Loop

    UrlDecode = result
End Function

Private Function IsHexPair(ByVal value As String) As Boolean
    IsHexPair = Len(value) = 2 And IsHexChar(Left$(value, 1)) And IsHexChar(Right$(value, 1))
End Function

Private Function IsHexChar(ByVal value As String) As Boolean
    Dim code As Integer
    If Len(value) <> 1 Then Exit Function

    code = Asc(UCase$(value))
    IsHexChar = (code >= Asc("0") And code <= Asc("9")) _
                Or (code >= Asc("A") And code <= Asc("F"))
End Function

Private Function Utf8BytesToString(ByRef bytes() As Byte) As String
    Dim result As String
    Dim i As Long
    Dim codePoint As Long
    Dim byteValue As Long
    Dim upperBound As Long

    On Error GoTo Done
    upperBound = UBound(bytes)

    Do While i <= upperBound
        byteValue = CLng(bytes(i))

        If byteValue < &H80 Then
            codePoint = byteValue
            i = i + 1
        ElseIf (byteValue And &HE0) = &HC0 And i + 1 <= upperBound Then
            codePoint = ((byteValue And &H1F) * &H40) _
                        + (CLng(bytes(i + 1)) And &H3F)
            i = i + 2
        ElseIf (byteValue And &HF0) = &HE0 And i + 2 <= upperBound Then
            codePoint = ((byteValue And &HF) * &H1000) _
                        + ((CLng(bytes(i + 1)) And &H3F) * &H40) _
                        + (CLng(bytes(i + 2)) And &H3F)
            i = i + 3
        ElseIf (byteValue And &HF8) = &HF0 And i + 3 <= upperBound Then
            codePoint = ((byteValue And &H7) * &H40000) _
                        + ((CLng(bytes(i + 1)) And &H3F) * &H1000) _
                        + ((CLng(bytes(i + 2)) And &H3F) * &H40) _
                        + (CLng(bytes(i + 3)) And &H3F)
            i = i + 4
        Else
            codePoint = byteValue
            i = i + 1
        End If

        result = result & CodePointToString(codePoint)
    Loop

Done:
    Utf8BytesToString = result
End Function

Private Function CodePointToString(ByVal codePoint As Long) As String
    If codePoint <= &HFFFF Then
        CodePointToString = ChrWCompat(codePoint)
    ElseIf codePoint <= &H10FFFF Then
        codePoint = codePoint - &H10000
        CodePointToString = ChrWCompat(&HD800 + (codePoint \ &H400)) _
                            & ChrWCompat(&HDC00 + (codePoint Mod &H400))
    End If
End Function

Private Function ChrWCompat(ByVal codePoint As Long) As String
    If codePoint > &H7FFF Then
        ChrWCompat = ChrW$(codePoint - &H10000)
    Else
        ChrWCompat = ChrW$(codePoint)
    End If
End Function

Private Function CombinePath(ByVal basePath As String, ByVal childPath As String) As String
    basePath = TrimTrailingBackslash(basePath)
    childPath = TrimLeadingBackslash(childPath)

    If Len(basePath) = 0 Then
        CombinePath = childPath
    ElseIf Len(childPath) = 0 Then
        CombinePath = basePath
    Else
        CombinePath = basePath & PATH_SEP & childPath
    End If
End Function

Private Function CandidateWithoutFirstPathSegment(ByVal localRoot As String, _
                                                  ByVal relPath As String) As String
    Dim reducedRelPath As String
    reducedRelPath = PathWithoutFirstSegment(relPath)
    If Len(reducedRelPath) = 0 Then Exit Function

    CandidateWithoutFirstPathSegment = CombinePath(localRoot, reducedRelPath)
End Function

Private Function PathWithoutFirstSegment(ByVal value As String) As String
    value = TrimLeadingBackslash(value)

    Dim separatorPos As Long
    separatorPos = InStr(1, value, PATH_SEP, vbBinaryCompare)
    If separatorPos = 0 Then Exit Function

    PathWithoutFirstSegment = Mid$(value, separatorPos + 1)
End Function

Private Function TrimTrailingBackslash(ByVal value As String) As String
    value = Trim$(value)
    Do While Len(value) > 3 And Right$(value, 1) = PATH_SEP
        value = Left$(value, Len(value) - 1)
    Loop
    TrimTrailingBackslash = value
End Function

Private Function TrimLeadingBackslash(ByVal value As String) As String
    value = Trim$(value)
    Do While Len(value) > 0 And Left$(value, 1) = PATH_SEP
        value = Mid$(value, 2)
    Loop
    TrimLeadingBackslash = value
End Function

Private Function TrimTrailingSlash(ByVal value As String) As String
    value = Trim$(value)
    Do While Len(value) > 0 And Right$(value, 1) = "/"
        value = Left$(value, Len(value) - 1)
    Loop
    TrimTrailingSlash = value
End Function

Private Function FileExists(ByVal path As String) As Boolean
    On Error Resume Next
    Dim attr As Long
    attr = GetAttr(path)
    FileExists = (Err.Number = 0 And Not CBool(attr And vbDirectory))
    Err.Clear
    On Error GoTo 0
End Function

Private Function FolderExists(ByVal path As String) As Boolean
    On Error Resume Next
    Dim attr As Long
    attr = GetAttr(path)
    FolderExists = (Err.Number = 0 And CBool(attr And vbDirectory))
    Err.Clear
    On Error GoTo 0
End Function

Private Function PathExists(ByVal path As String) As Boolean
    On Error Resume Next
    Dim attr As Long
    attr = GetAttr(path)
    PathExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function
