Attribute VB_Name = "modConvURLtoLocalPath"
Option Explicit

' Converts a SharePoint/OneDrive URL returned by Workbook.Path into a local
' OneDrive sync path, without Win32 API declarations or shell automation.
'
' Public entry point:
'     localPath = convURLtoLocalPath(ThisWorkbook)
'     localPath = convURLtoLocalPath(ThisWorkbook.Path)
'     localPath = convURLtoLocalPath(ThisWorkbook.FullName)
'
' Notes:
' - Passing a Workbook-like object resolves its Path property.
' - Passing a String resolves that exact String.
' - The implementation reads OneDrive settings files only.
' - It does not write files, write registry values, run commands, or access
'   the network.

Private Const CACHE_SECONDS As Long = 30
Private Const PATH_SEP As String = "\"

Public Function convURLtoLocalPath(ByVal source As Variant, _
                                   Optional ByVal rebuildCache As Boolean = False, _
                                   Optional ByVal returnInputOnFail As Boolean = True, _
                                   Optional ByVal requireExists As Boolean = True) As String
    convURLtoLocalPath = ConvertURLtoLocalPathInternal(source, rebuildCache, _
                                                       returnInputOnFail, _
                                                       requireExists, True)
End Function

Public Function convURLtoLocalPathLight(ByVal source As Variant, _
                                        Optional ByVal rebuildCache As Boolean = False, _
                                        Optional ByVal returnInputOnFail As Boolean = True, _
                                        Optional ByVal requireExists As Boolean = True) As String
    convURLtoLocalPathLight = ConvertURLtoLocalPathInternal(source, rebuildCache, _
                                                            returnInputOnFail, _
                                                            requireExists, False)
End Function

Private Function ConvertURLtoLocalPathInternal(ByVal source As Variant, _
                                               ByVal rebuildCache As Boolean, _
                                               ByVal returnInputOnFail As Boolean, _
                                               ByVal requireExists As Boolean, _
                                               ByVal useDatabase As Boolean) As String
    Dim originalPath As String
    originalPath = GetInputPath(source)

    If Len(originalPath) = 0 Then Exit Function
    If Not IsHttpsUrl(originalPath) Then
        ConvertURLtoLocalPathInternal = originalPath
        Exit Function
    End If

    On Error GoTo Fail

    Dim providers As Collection
    Set providers = GetProviderCache(rebuildCache, useDatabase)

    Dim resolvedPath As String
    resolvedPath = ResolveWithProviders(NormalizeUrl(originalPath), providers, requireExists)

    If Len(resolvedPath) = 0 And Not rebuildCache Then
        Set providers = GetProviderCache(True, useDatabase)
        resolvedPath = ResolveWithProviders(NormalizeUrl(originalPath), providers, requireExists)
    End If

    If Len(resolvedPath) > 0 Then
        ConvertURLtoLocalPathInternal = resolvedPath
    ElseIf returnInputOnFail Then
        ConvertURLtoLocalPathInternal = originalPath
    End If

    Exit Function

Fail:
    If returnInputOnFail Then ConvertURLtoLocalPathInternal = originalPath
End Function

Public Sub convURLtoLocalPathDebug(Optional ByVal source As Variant, _
                                   Optional ByVal writeToSheet As Boolean = True, _
                                   Optional ByVal useDatabase As Boolean = True)
    Dim rows As New Collection
    On Error GoTo Failed

    Dim inputPath As String

    If IsMissing(source) Then
        inputPath = GetInputPath(ThisWorkbook)
    Else
        inputPath = GetInputPath(source)
    End If

    AddDebugRow rows, "Input", "Resolved input", inputPath
    AddDebugRow rows, "Input", "ThisWorkbook.Path", SafeWorkbookProperty("Path")
    AddDebugRow rows, "Input", "ThisWorkbook.FullName", SafeWorkbookProperty("FullName")
    AddDebugRow rows, "Input", "Is HTTPS URL", CStr(IsHttpsUrl(inputPath))
    AddDebugRow rows, "Input", "Use SyncEngineDatabase.db", CStr(useDatabase)

    Dim normalizedUrl As String
    If IsHttpsUrl(inputPath) Then
        normalizedUrl = NormalizeUrl(inputPath)
        AddDebugRow rows, "Input", "Normalized URL", normalizedUrl
    End If

    AddSettingsDebugRows rows, useDatabase

    Dim providers As Collection
    Set providers = BuildProvidersFromOneDriveSettings(useDatabase)
    AddDebugRow rows, "Providers", "Total providers", CStr(providers.Count)
    AddProviderDebugRows rows, providers, normalizedUrl

    If Len(normalizedUrl) > 0 Then
        AddDebugRow rows, "Resolve", "requireExists=True", _
                    ResolveWithProviders(normalizedUrl, providers, True)
        AddDebugRow rows, "Resolve", "requireExists=False", _
                    ResolveWithProviders(normalizedUrl, providers, False)
        AddDebugRow rows, "Resolve", "Public function result", _
                    ConvertURLtoLocalPathInternal(inputPath, True, False, True, useDatabase)
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

Public Sub DebugThisWorkbookLocalPath()
    convURLtoLocalPathDebug ThisWorkbook, True
    MsgBox "Diagnostic log was written." & vbCrLf & _
           "Sheet: OneDrive Path Debug" & vbCrLf & _
           "Immediate: press Ctrl+G in the VBE", vbInformation
End Sub

Public Sub DebugThisWorkbookLocalPathLight()
    convURLtoLocalPathDebug ThisWorkbook, True, False
    MsgBox "Light diagnostic log was written." & vbCrLf & _
           "Sheet: OneDrive Path Debug" & vbCrLf & _
           "Immediate: press Ctrl+G in the VBE", vbInformation
End Sub

Private Function GetInputPath(ByVal source As Variant) As String
    On Error GoTo FromValue

    If IsObject(source) Then
        GetInputPath = CStr(CallByName(source, "Path", VbGet))
        If Len(GetInputPath) > 0 Then Exit Function

        Err.Clear
        GetInputPath = CStr(CallByName(source, "FullName", VbGet))
        Exit Function
    End If

FromValue:
    On Error Resume Next
    GetInputPath = CStr(source)
    On Error GoTo 0
End Function

Private Function SafeWorkbookProperty(ByVal propertyName As String) As String
    On Error GoTo Done
    SafeWorkbookProperty = CStr(CallByName(ThisWorkbook, propertyName, VbGet))
Done:
End Function

Private Sub AddDebugRow(ByVal rows As Collection, _
                        ByVal stage As String, _
                        ByVal key As String, _
                        ByVal value As String, _
                        Optional ByVal note As String = vbNullString)
    rows.Add Array(stage, key, value, note)
End Sub

Private Sub PrintDebugRows(ByVal rows As Collection)
    Dim row As Variant
    For Each row In rows
        Debug.Print CStr(row(0)) & " | " & CStr(row(1)) & _
                    " | " & CStr(row(2)) & " | " & CStr(row(3))
    Next row
End Sub

Private Sub WriteDebugRowsToSheet(ByVal rows As Collection)
    On Error GoTo Done

    Const sheetName As String = "OneDrive Path Debug"

    Dim ws As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo Done

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If

    ws.Cells(1, 1).Value = "Stage"
    ws.Cells(1, 2).Value = "Key"
    ws.Cells(1, 3).Value = "Value"
    ws.Cells(1, 4).Value = "Note"

    Dim i As Long
    Dim row As Variant
    For i = 1 To rows.Count
        row = rows(i)
        ws.Cells(i + 1, 1).Value = CStr(row(0))
        ws.Cells(i + 1, 2).Value = CStr(row(1))
        ws.Cells(i + 1, 3).Value = CStr(row(2))
        ws.Cells(i + 1, 4).Value = CStr(row(3))
    Next i

    ws.Columns("A:D").AutoFit

Done:
End Sub

Private Function GetProviderCache(ByVal rebuildCache As Boolean, _
                                  ByVal useDatabase As Boolean) As Collection
    Static cachedProvidersWithDb As Collection
    Static cachedProvidersLight As Collection
    Static lastBuiltWithDb As Date
    Static lastBuiltLight As Date

    If useDatabase Then
        If cachedProvidersWithDb Is Nothing _
           Or rebuildCache _
           Or DateDiff("s", lastBuiltWithDb, Now) > CACHE_SECONDS Then
            Set cachedProvidersWithDb = BuildProvidersFromOneDriveSettings(True)
            lastBuiltWithDb = Now
        End If

        Set GetProviderCache = cachedProvidersWithDb
    Else
        If cachedProvidersLight Is Nothing _
           Or rebuildCache _
           Or DateDiff("s", lastBuiltLight, Now) > CACHE_SECONDS Then
            Set cachedProvidersLight = BuildProvidersFromOneDriveSettings(False)
            lastBuiltLight = Now
        End If

        Set GetProviderCache = cachedProvidersLight
    End If
End Function

Private Function BuildProvidersFromOneDriveSettings(Optional ByVal useDatabase As Boolean = True) As Collection
    Dim providers As New Collection

    Dim settingsRoot As String
    settingsRoot = CombinePath(Environ$("LOCALAPPDATA"), "Microsoft\OneDrive\settings")
    If Not FolderExists(settingsRoot) Then
        Set BuildProvidersFromOneDriveSettings = providers
        Exit Function
    End If

    Dim accountFolders As Collection
    Set accountFolders = GetOneDriveAccountFolders(settingsRoot)

    Dim account As Variant
    For Each account In accountFolders
        ReadAccountProviders providers, CStr(account(0)), CStr(account(1)), useDatabase
    Next account

    Set BuildProvidersFromOneDriveSettings = providers
End Function

Private Function GetOneDriveAccountFolders(ByVal settingsRoot As String) As Collection
    Dim result As New Collection

    Dim folderName As String
    folderName = Dir(CombinePath(settingsRoot, "*"), vbDirectory)

    Do While Len(folderName) > 0
        If folderName <> "." And folderName <> ".." Then
            If LCase$(folderName) = "personal" _
               Or LCase$(Left$(folderName, 8)) = "business" Then
                result.Add Array(CombinePath(settingsRoot, folderName), folderName)
            End If
        End If
        folderName = Dir
    Loop

    Set GetOneDriveAccountFolders = result
End Function

Private Sub AddSettingsDebugRows(ByVal rows As Collection, _
                                 ByVal useDatabase As Boolean)
    Dim settingsRoot As String
    settingsRoot = CombinePath(Environ$("LOCALAPPDATA"), "Microsoft\OneDrive\settings")

    AddDebugRow rows, "Settings", "Root", settingsRoot
    AddDebugRow rows, "Settings", "Root exists", CStr(FolderExists(settingsRoot))
    If Not FolderExists(settingsRoot) Then Exit Sub

    Dim accountFolders As Collection
    Set accountFolders = GetOneDriveAccountFolders(settingsRoot)
    AddDebugRow rows, "Settings", "Account folder count", CStr(accountFolders.Count)

    Dim account As Variant
    For Each account In accountFolders
        AddAccountDebugRows rows, CStr(account(0)), CStr(account(1)), useDatabase
    Next account
End Sub

Private Sub AddAccountDebugRows(ByVal rows As Collection, _
                                ByVal accountFolder As String, _
                                ByVal accountName As String, _
                                ByVal useDatabase As Boolean)
    On Error GoTo Failed

    Dim stage As String
    stage = "Account " & accountName

    AddDebugRow rows, stage, "Folder", accountFolder
    AddDebugRow rows, stage, "Folder exists", CStr(FolderExists(accountFolder))

    Dim globalPath As String
    globalPath = CombinePath(accountFolder, "global.ini")
    AddDebugRow rows, stage, "global.ini exists", CStr(FileExists(globalPath))

    Dim cidSource As String
    Dim globalCid As String
    globalCid = IniValue(globalPath, "cid")
    AddDebugRow rows, stage, "global.ini cid", globalCid

    Dim cid As String
    cid = ResolveAccountIniId(accountFolder, cidSource)
    AddDebugRow rows, stage, "cid source", cidSource
    AddDebugRow rows, stage, "cid", cid
    If Len(cid) = 0 Then Exit Sub

    Dim cidIniPath As String
    cidIniPath = CombinePath(accountFolder, cid & ".ini")
    AddDebugRow rows, stage, "<cid>.ini exists", CStr(FileExists(cidIniPath))
    If FileExists(cidIniPath) Then AddCidIniDebugRows rows, stage, cidIniPath

    Dim policies As Collection
    Set policies = ReadClientPolicies(accountFolder)
    AddDebugRow rows, stage, "ClientPolicy count", CStr(policies.Count)

    Dim policy As Variant
    For Each policy In policies
        AddDebugRow rows, stage & " ClientPolicy", CStr(policy(0)), _
                    "DavUrlNamespace=" & CStr(policy(1)) & _
                    " | SiteID=" & CStr(policy(2)) & _
                    " | WebID=" & CStr(policy(3)) & _
                    " | IrmLibraryId=" & CStr(policy(4))
    Next policy

    Dim datPath As String
    datPath = CombinePath(accountFolder, cid & ".dat")
    AddDebugRow rows, stage, "<cid>.dat exists", CStr(FileExists(datPath))

    Dim datFolders As New Collection
    If FileExists(datPath) Then ReadFoldersFromDat datPath, datFolders
    AddDebugRow rows, stage, "<cid>.dat folder count", CStr(datFolders.Count)

    Dim dbPath As String
    dbPath = CombinePath(accountFolder, "SyncEngineDatabase.db")
    AddDebugRow rows, stage, "SyncEngineDatabase.db exists", CStr(FileExists(dbPath))

    Dim dbFolders As New Collection
    If useDatabase And FileExists(dbPath) Then
        ReadFoldersFromDb dbPath, dbFolders, LCase$(accountName) = "personal"
    End If
    If useDatabase Then
        AddDebugRow rows, stage, "SyncEngineDatabase.db folder count", CStr(dbFolders.Count)
    Else
        AddDebugRow rows, stage, "SyncEngineDatabase.db folder count", "skipped"
    End If

    Dim accountProviders As New Collection
    ReadAccountProviders accountProviders, accountFolder, accountName, useDatabase
    AddDebugRow rows, stage, "Provider count from account", CStr(accountProviders.Count)

    Exit Sub

Failed:
    AddDebugRow rows, stage, "ERROR", Err.Number & ": " & Err.Description
End Sub

Private Sub AddCidIniDebugRows(ByVal rows As Collection, _
                               ByVal stage As String, _
                               ByVal cidIniPath As String)
    On Error GoTo Failed

    Dim iniText As String
    iniText = ReadTextFile(cidIniPath)
    AddDebugRow rows, stage, "<cid>.ini text length", CStr(Len(iniText))
    If Len(iniText) = 0 Then Exit Sub

    Dim sortedLines As Collection
    Set sortedLines = SortedOneDriveSettingLines(iniText)
    AddDebugRow rows, stage, "<cid>.ini tracked line count", CStr(sortedLines.Count)

    Dim lineText As Variant
    Dim tokens As Variant
    Dim index As Long
    Dim tagName As String

    For Each lineText In sortedLines
        index = index + 1
        tokens = ParseSettingLine(CStr(lineText))
        tagName = TokenAt(tokens, 0)

        Select Case tagName
            Case "libraryScope"
                AddDebugRow rows, stage & " cid.ini", _
                            tagName & " #" & CStr(index), _
                            "libNo=" & TokenAt(tokens, 2) & _
                            " | siteId=" & TokenAt(tokens, 10) & _
                            " | webId=" & TokenAt(tokens, 11) & _
                            " | libId=" & TokenAt(tokens, 12) & _
                            " | localRoot=" & TokenAt(tokens, 14) & _
                            " | localLooks=" & CStr(LooksLikeLocalPath(TokenAt(tokens, 14))) & _
                            " | directUrl=" & DirectUrlFromTokens(tokens)

            Case "libraryFolder"
                AddDebugRow rows, stage & " cid.ini", _
                            tagName & " #" & CStr(index), _
                            "libNo=" & TokenAt(tokens, 3) & _
                            " | folderId=" & TokenAt(tokens, 4) & _
                            " | localRoot=" & TokenAt(tokens, 6) & _
                            " | localLooks=" & CStr(LooksLikeLocalPath(TokenAt(tokens, 6))) & _
                            " | directUrl=" & DirectUrlFromTokens(tokens)

            Case "AddedScope"
                AddDebugRow rows, stage & " cid.ini", _
                            tagName & " #" & CStr(index), _
                            "folderId=" & TokenAt(tokens, 3) & _
                            " | siteId=" & TokenAt(tokens, 7) & _
                            " | webId=" & TokenAt(tokens, 8) & _
                            " | libId=" & TokenAt(tokens, 9) & _
                            " | linkId=" & TokenAt(tokens, 10) & _
                            " | relPath=" & TokenAt(tokens, 11) & _
                            " | directUrl=" & DirectUrlFromTokens(tokens)
        End Select
    Next lineText

    Exit Sub

Failed:
    AddDebugRow rows, stage, "<cid>.ini parse ERROR", _
                Err.Number & ": " & Err.Description
End Sub

Private Sub AddProviderDebugRows(ByVal rows As Collection, _
                                 ByVal providers As Collection, _
                                 ByVal normalizedUrl As String)
    If providers Is Nothing Then Exit Sub

    Dim item As Variant
    Dim index As Long
    Dim localRoot As String
    Dim webRoot As String
    Dim relPart As String
    Dim candidate As String
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
            AddDebugRow rows, "Provider " & CStr(index), "candidate", candidate, _
                        "exists=" & CStr(PathExists(candidate))
        End If
    Next item
End Sub

Private Sub ReadAccountProviders(ByVal providers As Collection, _
                                 ByVal accountFolder As String, _
                                 ByVal accountName As String, _
                                 Optional ByVal useDatabase As Boolean = True)
    On Error GoTo Done

    If Not FolderExists(accountFolder) Then Exit Sub

    Dim cid As String
    Dim cidSource As String
    cid = ResolveAccountIniId(accountFolder, cidSource)
    If Len(cid) = 0 Then Exit Sub

    Dim policies As Collection
    Set policies = ReadClientPolicies(accountFolder)
    If policies.Count = 0 Then Exit Sub

    Dim cidIniPath As String
    cidIniPath = CombinePath(accountFolder, cid & ".ini")
    If Not FileExists(cidIniPath) Then Exit Sub

    Dim folders As Collection
    Set folders = ReadOneDriveFolderMap(accountFolder, cid, accountName, useDatabase)

    If LCase$(accountName) = "personal" Then
        ReadPersonalProviders providers, accountFolder, cid, policies, folders
    Else
        ReadBusinessProviders providers, accountFolder, cid, policies, folders
    End If

Done:
End Sub

Private Function ResolveAccountIniId(ByVal accountFolder As String, _
                                     ByRef cidSource As String) As String
    Dim globalCid As String
    globalCid = IniValue(CombinePath(accountFolder, "global.ini"), "cid")

    If Len(globalCid) > 0 Then
        cidSource = "global.ini"
        ResolveAccountIniId = globalCid
        If FileExists(CombinePath(accountFolder, globalCid & ".ini")) Then Exit Function
        cidSource = "global.ini, but <cid>.ini missing; fallback .ini scan"
    Else
        cidSource = "fallback .ini scan"
    End If

    Dim fallbackCid As String
    fallbackCid = FindAccountIniId(accountFolder)
    If Len(fallbackCid) > 0 Then
        ResolveAccountIniId = fallbackCid
        Exit Function
    End If

    If Len(globalCid) > 0 Then ResolveAccountIniId = globalCid
End Function

Private Function FindAccountIniId(ByVal accountFolder As String) As String
    Dim candidates As New Collection

    Dim fileName As String
    fileName = Dir(CombinePath(accountFolder, "*.ini"))

    Do While Len(fileName) > 0
        If IsCandidateAccountIniFile(fileName) Then
            candidates.Add fileName
        End If

        fileName = Dir
    Loop

    Dim item As Variant
    For Each item In candidates
        Dim fileText As String
        fileText = ReadTextFile(CombinePath(accountFolder, CStr(item)))

        If LooksLikeAccountIniText(fileText) Then
            FindAccountIniId = Left$(CStr(item), Len(CStr(item)) - 4)
            Exit Function
        End If
    Next item
End Function

Private Function IsCandidateAccountIniFile(ByVal fileName As String) As Boolean
    Dim lowerName As String
    lowerName = LCase$(fileName)

    If lowerName = "global.ini" Then Exit Function
    If lowerName = "groupfolders.ini" Then Exit Function
    If lowerName Like "clientpolicy*.ini" Then Exit Function
    If lowerName Like "loguploadersettings*.ini" Then Exit Function

    IsCandidateAccountIniFile = (Right$(lowerName, 4) = ".ini")
End Function

Private Function LooksLikeAccountIniText(ByVal fileText As String) As Boolean
    LooksLikeAccountIniText = _
        InStr(1, fileText, "libraryScope = ", vbBinaryCompare) > 0 _
        Or InStr(1, fileText, "libraryFolder = ", vbBinaryCompare) > 0 _
        Or InStr(1, fileText, "AddedScope = ", vbBinaryCompare) > 0 _
        Or InStr(1, fileText, "library = ", vbBinaryCompare) > 0
End Function

Private Sub ReadPersonalProviders(ByVal providers As Collection, _
                                  ByVal accountFolder As String, _
                                  ByVal cid As String, _
                                  ByVal policies As Collection, _
                                  ByVal folders As Collection)
    Dim iniText As String
    iniText = ReadTextFile(CombinePath(accountFolder, cid & ".ini"))
    If Len(iniText) = 0 Then Exit Sub

    Dim mainLocalRoot As String
    Dim lineText As Variant
    Dim parts As Variant

    For Each lineText In SplitLines(iniText)
        If CStr(lineText) Like "library = *" Then
            parts = Split(CStr(lineText), """")
            If UBoundSafe(parts) >= 3 Then mainLocalRoot = CStr(parts(3))
            Exit For
        ElseIf CStr(lineText) Like "libraryScope = *" Then
            parts = Split(CStr(lineText), """")
            If UBoundSafe(parts) >= 9 Then mainLocalRoot = CStr(parts(9))
            If Len(mainLocalRoot) = 0 Then mainLocalRoot = FirstLocalPath(ParseSettingLine(CStr(lineText)))
            Exit For
        End If
    Next lineText

    Dim mainWebRoot As String
    mainWebRoot = PolicyValueByFile(policies, "ClientPolicy.ini", "DavUrlNamespace")
    If Len(mainLocalRoot) > 0 And Len(mainWebRoot) > 0 Then
        AddProvider providers, mainLocalRoot, CombineUrl(mainWebRoot, cid)
    End If

    ReadPersonalGroupFolders providers, accountFolder, mainLocalRoot, mainWebRoot, folders
End Sub

Private Sub ReadPersonalGroupFolders(ByVal providers As Collection, _
                                     ByVal accountFolder As String, _
                                     ByVal mainLocalRoot As String, _
                                     ByVal mainWebRoot As String, _
                                     ByVal folders As Collection)
    If Len(mainLocalRoot) = 0 Or Len(mainWebRoot) = 0 Then Exit Sub

    Dim groupPath As String
    groupPath = CombinePath(accountFolder, "GroupFolders.ini")
    If Not FileExists(groupPath) Then Exit Sub

    Dim groupText As String
    groupText = ReadTextFile(groupPath)
    If Len(groupText) = 0 Then Exit Sub

    Dim lineText As Variant
    Dim remoteCid As String
    Dim folderId As String
    Dim localName As String
    Dim remoteRelPath As String

    For Each lineText In SplitLines(groupText)
        If InStr(1, CStr(lineText), "_BaseUri = ", vbTextCompare) > 0 Then
            remoteCid = CStr(lineText)
            remoteCid = Mid$(remoteCid, InStrRev(remoteCid, "/") + 1)
            If InStr(1, remoteCid, "!", vbBinaryCompare) > 0 Then
                remoteCid = Left$(remoteCid, InStr(1, remoteCid, "!", vbBinaryCompare) - 1)
            End If
            folderId = Left$(CStr(lineText), InStr(1, CStr(lineText), "_", vbBinaryCompare) - 1)
        ElseIf Len(remoteCid) > 0 And InStr(1, CStr(lineText), "_Path = ", vbTextCompare) > 0 Then
            localName = FolderNameForId(folders, folderId)
            remoteRelPath = Mid$(CStr(lineText), Len(folderId) + 9)
            If Len(localName) > 0 And Len(remoteRelPath) > 0 Then
                AddProvider providers, CombinePath(mainLocalRoot, localName), _
                            CombineUrl(CombineUrl(mainWebRoot, remoteCid), remoteRelPath)
            End If
            remoteCid = vbNullString
            folderId = vbNullString
        End If
    Next lineText
End Sub

Private Sub ReadBusinessProviders(ByVal providers As Collection, _
                                  ByVal accountFolder As String, _
                                  ByVal cid As String, _
                                  ByVal policies As Collection, _
                                  ByVal folders As Collection)
    Dim iniText As String
    iniText = ReadTextFile(CombinePath(accountFolder, cid & ".ini"))
    If Len(iniText) = 0 Then Exit Sub

    Dim sortedLines As Collection
    Set sortedLines = SortedOneDriveSettingLines(iniText)

    Dim libToWebRoot As New Collection
    Dim mainLocalRoot As String
    Dim lineText As Variant
    Dim tokens As Variant

    For Each lineText In sortedLines
        tokens = ParseSettingLine(CStr(lineText))

        Select Case TokenAt(tokens, 0)
            Case "libraryScope"
                ProcessBusinessLibraryScope providers, policies, libToWebRoot, _
                                            mainLocalRoot, tokens

            Case "libraryFolder"
                ProcessBusinessLibraryFolder providers, libToWebRoot, folders, tokens

            Case "AddedScope"
                ProcessBusinessAddedScope providers, policies, folders, _
                                          mainLocalRoot, tokens
        End Select
    Next lineText
End Sub

Private Sub ProcessBusinessLibraryScope(ByVal providers As Collection, _
                                        ByVal policies As Collection, _
                                        ByVal libToWebRoot As Collection, _
                                        ByRef mainLocalRoot As String, _
                                        ByVal tokens As Variant)
    Dim localRoot As String
    localRoot = TokenAt(tokens, 14)
    If Not LooksLikeLocalPath(localRoot) Then localRoot = FirstLocalPath(tokens)

    Dim libNo As String
    Dim siteId As String
    Dim webId As String
    Dim libId As String
    Dim webRoot As String

    libNo = TokenAt(tokens, 2)
    siteId = TokenAt(tokens, 10)
    webId = TokenAt(tokens, 11)
    libId = TokenAt(tokens, 12)

    webRoot = DirectUrlFromTokens(tokens)

    If Len(webRoot) > 0 Then
        If Len(localRoot) > 0 And Len(mainLocalRoot) = 0 Then mainLocalRoot = localRoot
    ElseIf libNo = "0" Then
        webRoot = PolicyValueByFile(policies, "ClientPolicy.ini", "DavUrlNamespace")
        If Len(localRoot) > 0 Then mainLocalRoot = localRoot
    Else
        webRoot = PolicyValueByFile(policies, "ClientPolicy_" & libId & siteId & ".ini", "DavUrlNamespace")
    End If

    If Len(webRoot) = 0 Then
        webRoot = PolicyValueByIds(policies, siteId, webId, libId, "DavUrlNamespace")
    End If

    If Len(localRoot) > 0 And Len(webRoot) > 0 Then
        AddProvider providers, localRoot, webRoot
        AddKeyedValue libToWebRoot, libNo, webRoot
    End If
End Sub

Private Sub ProcessBusinessLibraryFolder(ByVal providers As Collection, _
                                         ByVal libToWebRoot As Collection, _
                                         ByVal folders As Collection, _
                                         ByVal tokens As Variant)
    Dim localRoot As String
    localRoot = TokenAt(tokens, 6)
    If Not LooksLikeLocalPath(localRoot) Then localRoot = FirstLocalPath(tokens)
    If Len(localRoot) = 0 Then Exit Sub

    Dim webRoot As String
    webRoot = DirectUrlFromTokens(tokens)
    Dim hasDirectUrl As Boolean
    hasDirectUrl = Len(webRoot) > 0

    If Len(webRoot) = 0 Then webRoot = KeyedValue(libToWebRoot, TokenAt(tokens, 3))
    If Len(webRoot) = 0 Then Exit Sub

    Dim remoteFolder As String
    remoteFolder = FolderPathForId(folders, CleanFolderId(TokenAt(tokens, 4)), "/")

    If hasDirectUrl Then
        AddProvider providers, localRoot, webRoot
    ElseIf Len(remoteFolder) > 0 Then
        AddProvider providers, localRoot, CombineUrl(webRoot, remoteFolder)
    Else
        AddProvider providers, localRoot, CombineUrl(webRoot, LastPathPart(localRoot))
    End If
End Sub

Private Sub ProcessBusinessAddedScope(ByVal providers As Collection, _
                                      ByVal policies As Collection, _
                                      ByVal folders As Collection, _
                                      ByVal mainLocalRoot As String, _
                                      ByVal tokens As Variant)
    If Len(mainLocalRoot) = 0 Then Exit Sub

    Dim siteId As String
    Dim webId As String
    Dim libId As String
    Dim linkId As String
    Dim relPath As String
    Dim webRoot As String

    siteId = TokenAt(tokens, 7)
    webId = TokenAt(tokens, 8)
    libId = TokenAt(tokens, 9)
    linkId = TokenAt(tokens, 10)
    relPath = TokenAt(tokens, 11)
    If relPath = " " Then relPath = vbNullString

    webRoot = DirectUrlFromTokens(tokens)
    Dim hasDirectUrl As Boolean
    hasDirectUrl = Len(webRoot) > 0

    If Len(webRoot) = 0 Then webRoot = PolicyValueByFile(policies, _
              "ClientPolicy_" & libId & siteId & linkId & ".ini", "DavUrlNamespace")
    If Len(webRoot) = 0 Then
        webRoot = PolicyValueByIds(policies, siteId, webId, libId, "DavUrlNamespace")
    End If
    If Len(webRoot) = 0 Then Exit Sub

    If Not hasDirectUrl And Len(relPath) > 0 Then webRoot = CombineUrl(webRoot, relPath)

    Dim localRelPath As String
    localRelPath = FolderPathForId(folders, CleanFolderId(TokenAt(tokens, 3)), PATH_SEP)
    If Len(localRelPath) = 0 Then localRelPath = RelativePathFromToken(TokenAt(tokens, 11), PATH_SEP)
    If Len(localRelPath) = 0 Then Exit Sub

    AddProvider providers, CombinePath(mainLocalRoot, localRelPath), webRoot
End Sub

Private Function ReadClientPolicies(ByVal accountFolder As String) As Collection
    Dim result As New Collection

    Dim policyFiles As Collection
    Set policyFiles = GetClientPolicyFiles(accountFolder)

    Dim fileName As Variant
    For Each fileName In policyFiles
        Dim text As String
        text = ReadTextFile(CombinePath(accountFolder, CStr(fileName)))

        If Len(text) > 0 Then
            Dim policy As Variant
            policy = Array(CStr(fileName), _
                           Trim$(GetTagValue(text, "DavUrlNamespace = ")), _
                           Trim$(GetTagValue(text, "SiteID = ")), _
                           Trim$(GetTagValue(text, "WebID = ")), _
                           Trim$(GetTagValue(text, "IrmLibraryId = ")))
            On Error Resume Next
            result.Add policy, LCase$(CStr(fileName))
            On Error GoTo 0
        End If
    Next fileName

    Set ReadClientPolicies = result
End Function

Private Function GetClientPolicyFiles(ByVal accountFolder As String) As Collection
    Dim result As New Collection

    Dim fileName As String
    fileName = Dir(CombinePath(accountFolder, "ClientPolicy*.ini"))

    Do While Len(fileName) > 0
        result.Add fileName
        fileName = Dir
    Loop

    Set GetClientPolicyFiles = result
End Function

Private Function PolicyValueByFile(ByVal policies As Collection, _
                                   ByVal fileName As String, _
                                   ByVal valueName As String) As String
    On Error GoTo Done
    Dim policy As Variant
    policy = policies(LCase$(fileName))
    PolicyValueByFile = PolicyValue(policy, valueName)
Done:
End Function

Private Function PolicyValueByIds(ByVal policies As Collection, _
                                  ByVal siteId As String, _
                                  ByVal webId As String, _
                                  ByVal libId As String, _
                                  ByVal valueName As String) As String
    Dim item As Variant
    For Each item In policies
        If StrComp(PolicyValue(item, "SiteID"), siteId, vbTextCompare) = 0 _
           And StrComp(PolicyValue(item, "WebID"), webId, vbTextCompare) = 0 _
           And StrComp(PolicyValue(item, "IrmLibraryId"), libId, vbTextCompare) = 0 Then
            PolicyValueByIds = PolicyValue(item, valueName)
            Exit Function
        End If
    Next item
End Function

Private Function PolicyValue(ByVal policy As Variant, ByVal valueName As String) As String
    Select Case valueName
        Case "FileName":        PolicyValue = CStr(policy(0))
        Case "DavUrlNamespace": PolicyValue = CStr(policy(1))
        Case "SiteID":          PolicyValue = CStr(policy(2))
        Case "WebID":           PolicyValue = CStr(policy(3))
        Case "IrmLibraryId":    PolicyValue = CStr(policy(4))
    End Select
End Function

Private Function SortedOneDriveSettingLines(ByVal iniText As String) As Collection
    Dim buckets As New Collection
    buckets.Add New Collection, "libraryScope"
    buckets.Add New Collection, "libraryFolder"
    buckets.Add New Collection, "AddedScope"

    Dim lineText As Variant
    For Each lineText In SplitLines(iniText)
        Dim tagName As String
        tagName = Left$(CStr(lineText), InStr(1, CStr(lineText) & " ", " ", vbBinaryCompare) - 1)

        Select Case tagName
            Case "libraryScope", "libraryFolder", "AddedScope"
                On Error Resume Next
                buckets(tagName).Add CStr(lineText), SortKeyFromLine(CStr(lineText))
                If Err.Number <> 0 Then
                    Err.Clear
                    buckets(tagName).Add CStr(lineText)
                End If
                On Error GoTo 0
        End Select
    Next lineText

    Dim result As New Collection
    Dim tag As Variant
    For Each tag In Array("libraryScope", "libraryFolder", "AddedScope")
        Dim bucket As Collection
        Set bucket = buckets(CStr(tag))

        Dim i As Long
        Do While bucket.Count > 0
            Dim value As String
            value = vbNullString

            On Error Resume Next
            value = CStr(bucket(CStr(i)))
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0

            If Len(value) > 0 Then
                result.Add value
                On Error Resume Next
                bucket.Remove CStr(i)
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0
            End If

            i = i + 1
            If i > 10000 Then Exit Do
        Loop
    Next tag

    Set SortedOneDriveSettingLines = result
End Function

Private Function SortKeyFromLine(ByVal lineText As String) As String
    Dim tokens As Variant
    tokens = ParseSettingLine(lineText)
    If TokenAt(tokens, 1) = "=" Then
        SortKeyFromLine = TokenAt(tokens, 2)
    Else
        SortKeyFromLine = TokenAt(tokens, 1)
    End If
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

Private Function ReadOneDriveFolderMap(ByVal accountFolder As String, _
                                       ByVal cid As String, _
                                       ByVal accountName As String, _
                                       Optional ByVal useDatabase As Boolean = True) As Collection
    Dim folders As New Collection

    Dim datPath As String
    datPath = CombinePath(accountFolder, cid & ".dat")
    If FileExists(datPath) Then
        ReadFoldersFromDat datPath, folders
    End If

    If useDatabase And folders.Count = 0 Then
        ReadFoldersFromDb CombinePath(accountFolder, "SyncEngineDatabase.db"), _
                          folders, LCase$(accountName) = "personal"
    End If

    Set ReadOneDriveFolderMap = folders
End Function

Private Sub ReadFoldersFromDat(ByVal filePath As String, ByVal folders As Collection)
    On Error GoTo CloseFile

    Dim fileNo As Long
    fileNo = FreeFile
    Open filePath For Binary Access Read As #fileNo

    Dim fileSize As Long
    fileSize = LOF(fileNo)
    If fileSize = 0 Then GoTo CloseFile

    Const headerSize As Long = 8
    Const idSize As Long = 39
    Const nameOffset As Long = 121
    Const checkToName As Long = headerSize + idSize + nameOffset + nameOffset
    Const chunkSize As Long = &H100000
    Const maxDirName As Long = 255

    Dim bytes(1 To chunkSize) As Byte
    Dim buffer As String
    Dim lastRecord As Long
    Dim i As Long
    Dim bytesCount As Long
    Dim dirId As String
    Dim parentId As String
    Dim dirName As String
    Dim idPattern As String
    Dim nullByte As String
    Dim folderMark As String
    Dim checkMark As String * 4

    nullByte = ChrB$(0)
    folderMark = ChrB$(2)
    MidB$(checkMark, 1) = ChrB$(1)
    idPattern = Replace(Space$(12), " ", "[a-fA-F0-9]") & "*"

    Dim stepSize As Long
    For stepSize = 16 To 8 Step -8
        lastRecord = 1

        Do
            Get fileNo, lastRecord, bytes
            buffer = bytes

            i = InStrB(stepSize + 1, buffer, checkMark)
            Do While i > 0 And i < chunkSize - checkToName
                If MidB$(buffer, i - stepSize, 1) = folderMark Then
                    i = i + headerSize

                    bytesCount = ClampLong(InStrB(i, buffer, nullByte) - i, 0, idSize)
                    dirId = StrConv(MidB$(buffer, i, bytesCount), vbUnicode)

                    i = i + idSize
                    bytesCount = ClampLong(InStrB(i, buffer, nullByte) - i, 0, idSize)
                    parentId = StrConv(MidB$(buffer, i, bytesCount), vbUnicode)

                    i = i + nameOffset
                    If dirId Like idPattern And parentId Like idPattern Then
                        bytesCount = InStr((i + 1) \ 2, buffer, vbNullChar) * 2 - i - 1
                        If bytesCount > maxDirName * 2 Then bytesCount = maxDirName * 2

                        If bytesCount < 0 Or i + bytesCount - 1 > chunkSize Then
                            i = i - checkToName
                            Exit Do
                        End If

                        dirName = MidB$(buffer, i, bytesCount)
                        AddFolder folders, dirId, parentId, dirName
                    End If
                End If

                i = InStrB(i + 1, buffer, checkMark)
            Loop

            lastRecord = lastRecord + chunkSize - stepSize
            If i > stepSize Then
                lastRecord = lastRecord - chunkSize + (i \ 2) * 2
            End If
        Loop Until lastRecord > fileSize

        If folders.Count > 0 Then Exit For
    Next stepSize

CloseFile:
    On Error Resume Next
    If fileNo <> 0 Then Close #fileNo
    On Error GoTo 0
End Sub

Private Sub ReadFoldersFromDb(ByVal filePath As String, _
                              ByVal folders As Collection, _
                              ByVal isPersonal As Boolean)
    On Error GoTo CloseFile
    If Not FileExists(filePath) Then Exit Sub

    Dim fileNo As Long
    fileNo = FreeFile
    Open filePath For Binary Access Read As #fileNo

    Dim fileSize As Long
    fileSize = LOF(fileNo)
    If fileSize = 0 Then GoTo CloseFile

    Const chunkSize As Long = &H100000
    Const minName As Long = 15
    Const maxSigByte As Byte = 9
    Const maxHeader As Long = 21
    Const minIdSize As Long = 12
    Const maxIdSize As Long = 48
    Const minThreeIdSizes As Long = minIdSize * 3
    Const maxThreeIdSizes As Long = maxIdSize * 3
    Const leadingBuffer As Long = maxHeader + maxThreeIdSizes
    Const headBytesOffset As Long = 15
    Const bangCode As Long = 33

    Dim curlyStart As String
    Dim quoteByte As String
    Dim bangByte As String
    curlyStart = ChrW$(&H7B22)
    quoteByte = ChrB$(&H22)
    bangByte = ChrB$(bangCode)

    Dim signature As String
    Dim idPattern As String
    idPattern = Replace(Space$(12), " ", "[a-fA-F0-9]")

    If isPersonal Then
        signature = bangByte
        idPattern = "*" & idPattern & "![a-fA-F0-9]*"
    Else
        signature = curlyStart
        idPattern = idPattern & "*"
    End If

    Dim bytes(1 To chunkSize) As Byte
    Dim buffer As String
    Dim lastRecord As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim idSize(1 To 4) As Long
    Dim nameSize As Long
    Dim nameStart As Long
    Dim nameEnd As Long
    Dim folderId As String
    Dim parentId As String
    Dim tempId As String
    Dim folderName As String

    lastRecord = 1
    Do
        Get #fileNo, lastRecord, bytes
        buffer = bytes

        i = InStrB(1, buffer, signature)
        Do While i > 0
            If isPersonal Then
                For j = i - 1 To i - maxIdSize Step -1
                    If j = 0 Then GoTo NextSignature
                    If bytes(j) < bangCode Then Exit For
                Next j
                If j < maxHeader Or i - j < minIdSize Then GoTo NextSignature
            Else
                j = InStrB(i + 2, buffer, quoteByte)
                If j = 0 Then Exit Do

                idSize(4) = j - i + 1
                If idSize(4) > maxIdSize Then GoTo NextSignature

                For j = i - 1 To i - maxThreeIdSizes Step -1
                    If j = 0 Then GoTo NextSignature
                    If bytes(j) < bangCode Then Exit For
                Next j
                If j < maxHeader Then GoTo NextSignature

                idSize(1) = i - j - 1
                If idSize(1) < minThreeIdSizes Then GoTo NextSignature
            End If

            k = j + 1
            For j = j To j - headBytesOffset + 1 Step -1
                If bytes(j) > maxSigByte Then GoTo NextSignature
            Next j

            If j <= 4 Then GoTo NextSignature
            If bytes(j) <= maxSigByte And bytes(j - 1) < &H80 Then j = j - 1
            If j <= 4 Then GoTo NextSignature
            If bytes(j) < minName Then j = j - 1
            If j <= 4 Then GoTo NextSignature
            If bytes(j) < minName Then j = j - 1
            If j <= 4 Then GoTo NextSignature

            nameSize = bytes(j)
            If nameSize Mod 2 = 0 Then GoTo NextSignature

            nameSize = (nameSize - 13) / 2
            If bytes(j - 1) > &H7F Then
                nameSize = (bytes(j - 1) - &H80) * &H40 + nameSize
                j = j - 1
            End If
            If j < 5 Then GoTo NextSignature
            If nameSize < 1 Or bytes(j - 4) = 0 Then GoTo NextSignature

            If isPersonal Then
                idSize(4) = (bytes(j - 1) - 13) / 2
                idSize(3) = (bytes(j - 2) - 13) / 2
                idSize(2) = (bytes(j - 3) - 13) / 2
                idSize(1) = (bytes(j - 4) - 13) / 2
                nameStart = k + idSize(1) + idSize(2) + idSize(3) + idSize(4)
            Else
                If bytes(j - 1) <> idSize(4) * 2 + 13 Then GoTo NextSignature
                idSize(3) = (bytes(j - 2) - 13) / 2
                idSize(2) = (bytes(j - 3) - 13) / 2
                idSize(1) = idSize(1) - idSize(2) - idSize(3)
                nameStart = i + idSize(4)
            End If

            For j = 1 To 4
                If idSize(j) < minIdSize Or idSize(j) > maxIdSize Then GoTo NextSignature
            Next j

            nameEnd = nameStart + nameSize - 1
            If nameEnd > chunkSize Then Exit Do

            folderId = BinaryAnsiToString(MidB$(buffer, k, idSize(1)))
            If Not folderId Like idPattern Then GoTo NextSignature

            k = k + idSize(1)
            parentId = BinaryAnsiToString(MidB$(buffer, k, idSize(2)))
            If Not parentId Like idPattern Then GoTo NextSignature

            If isPersonal Then
                k = k + idSize(2)
                tempId = BinaryAnsiToString(MidB$(buffer, k, idSize(3)))
                If Not tempId Like idPattern Then GoTo NextSignature

                tempId = BinaryAnsiToString(MidB$(buffer, k + idSize(3), idSize(4)))
                If Not tempId Like idPattern Then GoTo NextSignature
            End If

            folderName = DecodeDbFolderName(buffer, bytes, nameStart, nameSize)
            AddFolder folders, folderId, parentId, folderName
            i = nameEnd

NextSignature:
            i = InStrB(i + 1, buffer, signature)
        Loop

        If i = 0 Then
            lastRecord = lastRecord + chunkSize - leadingBuffer
        ElseIf i > leadingBuffer Then
            lastRecord = lastRecord + i - leadingBuffer
        Else
            lastRecord = lastRecord + i
        End If
    Loop Until lastRecord > fileSize

CloseFile:
    On Error Resume Next
    If fileNo <> 0 Then Close #fileNo
    On Error GoTo 0
End Sub

Private Function DecodeDbFolderName(ByVal buffer As String, _
                                    ByRef sourceBytes() As Byte, _
                                    ByVal startByte As Long, _
                                    ByVal byteCount As Long) As String
    Dim raw As String
    raw = MidB$(buffer, startByte, byteCount)

    Dim i As Long
    For i = startByte To startByte + byteCount - 1
        If sourceBytes(i) > &H7F Then
            DecodeDbFolderName = Utf8BinaryStringToString(raw)
            Exit Function
        End If
    Next i

    DecodeDbFolderName = BinaryAnsiToString(raw)
End Function

Private Function BinaryAnsiToString(ByVal value As String) As String
    If LenB(value) = 0 Then Exit Function
    BinaryAnsiToString = StrConv(value, vbUnicode)
End Function

Private Function Utf8BinaryStringToString(ByVal value As String) As String
    If LenB(value) = 0 Then Exit Function

    Dim bytes() As Byte
    ReDim bytes(0 To LenB(value) - 1)

    Dim i As Long
    For i = 1 To LenB(value)
        bytes(i - 1) = AscB(MidB$(value, i, 1))
    Next i

    Utf8BinaryStringToString = Utf8BytesToString(bytes)
End Function

Private Sub AddFolder(ByVal folders As Collection, _
                      ByVal dirId As String, _
                      ByVal parentId As String, _
                      ByVal dirName As String)
    If Len(dirId) = 0 Or Len(dirName) = 0 Then Exit Sub
    On Error Resume Next
    folders.Add Array(parentId, dirName), dirId
    On Error GoTo 0
End Sub

Private Function FolderNameForId(ByVal folders As Collection, ByVal dirId As String) As String
    On Error GoTo Done
    FolderNameForId = CStr(folders(CleanFolderId(dirId))(1))
Done:
End Function

Private Function FolderPathForId(ByVal folders As Collection, _
                                 ByVal dirId As String, _
                                 ByVal separator As String) As String
    Dim parts As New Collection
    Dim currentId As String
    currentId = CleanFolderId(dirId)

    Dim depth As Long
    Do While Len(currentId) > 0 And depth < 128
        On Error GoTo Done
        Dim item As Variant
        item = folders(currentId)
        On Error GoTo 0

        If Len(CStr(item(1))) > 0 Then parts.Add CStr(item(1))
        currentId = CleanFolderId(CStr(item(0)))
        depth = depth + 1
    Loop

Done:
    Dim result As String
    Dim i As Long
    For i = parts.Count To 1 Step -1
        If Len(result) = 0 Then
            result = CStr(parts(i))
        Else
            result = result & separator & CStr(parts(i))
        End If
    Next i

    FolderPathForId = result
End Function

Private Function CleanFolderId(ByVal value As String) As String
    value = Trim$(value)
    If InStr(1, value, "+", vbBinaryCompare) > 0 Then
        value = Left$(value, InStr(1, value, "+", vbBinaryCompare) - 1)
    End If
    If Len(value) > 32 And value Like Replace(Space$(12), " ", "[a-fA-F0-9]") & "*" Then
        value = Left$(value, 32)
    End If
    CleanFolderId = value
End Function

Private Sub AddProvider(ByVal providers As Collection, ByVal localRoot As String, ByVal webRoot As String)
    localRoot = TrimTrailingBackslash(localRoot)
    webRoot = TrimTrailingSlash(NormalizeUrl(webRoot))

    If Len(localRoot) = 0 Or Len(webRoot) = 0 Then Exit Sub
    If Not LooksLikeLocalPath(localRoot) Then Exit Sub
    If Not IsHttpsUrl(webRoot) Then Exit Sub

    On Error Resume Next
    providers.Add Array(localRoot, webRoot), LCase$(localRoot & "|" & webRoot)
    On Error GoTo 0
End Sub

Private Function ResolveWithProviders(ByVal normalizedUrl As String, _
                                      ByVal providers As Collection, _
                                      ByVal requireExists As Boolean) As String
    If providers Is Nothing Then Exit Function

    Dim item As Variant
    Dim localRoot As String
    Dim webRoot As String
    Dim relPart As String
    Dim candidate As String
    Dim fallback As String

    For Each item In providers
        localRoot = CStr(item(0))
        webRoot = CStr(item(1))

        If UrlStartsWith(normalizedUrl, webRoot) Then
            relPart = Mid$(normalizedUrl, Len(webRoot) + 1)
            relPart = Replace(relPart, "/", PATH_SEP)
            candidate = CombinePath(localRoot, relPart)

            If Not requireExists Or PathExists(candidate) Then
                ResolveWithProviders = candidate
                Exit Function
            End If

            If Len(fallback) = 0 Then fallback = candidate
        End If
    Next item

    If Not requireExists Then ResolveWithProviders = fallback
End Function

Private Function IniValue(ByVal filePath As String, ByVal keyName As String) As String
    IniValue = Trim$(GetTagValue(ReadTextFile(filePath), keyName & " = "))
    If Len(IniValue) = 0 Then IniValue = Trim$(GetTagValue(ReadTextFile(filePath), keyName & "="))
End Function

Private Function GetTagValue(ByVal text As String, ByVal tag As String) As String
    Dim p As Long
    p = InStr(1, text, tag, vbTextCompare)
    If p = 0 Then Exit Function

    p = p + Len(tag)

    Dim e As Long
    e = InStr(p, text, vbLf, vbBinaryCompare)
    If e = 0 Then
        GetTagValue = Trim$(Mid$(text, p))
    Else
        GetTagValue = Trim$(Replace(Mid$(text, p, e - p), vbCr, vbNullString))
    End If
End Function

Private Function ReadTextFile(ByVal filePath As String) As String
    On Error GoTo CleanUp
    If Not FileExists(filePath) Then Exit Function

    Dim fileNo As Long
    fileNo = FreeFile
    Open filePath For Binary Access Read As #fileNo

    If LOF(fileNo) = 0 Then GoTo CleanUp

    Dim bytes() As Byte
    ReDim bytes(0 To LOF(fileNo) - 1)
    Get #fileNo, , bytes
    Close #fileNo
    fileNo = 0

    If BytesLookLikeUtf16Le(bytes) Then
        ReadTextFile = bytes
        If Len(ReadTextFile) > 0 Then
            If AscW(Left$(ReadTextFile, 1)) = -257 _
               Or AscW(Left$(ReadTextFile, 1)) = 65279 Then
                ReadTextFile = Mid$(ReadTextFile, 2)
            End If
        End If
    Else
        ReadTextFile = StrConv(bytes, vbUnicode)
    End If

CleanUp:
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
    Dim checked As Long
    Dim nullCount As Long

    For i = 1 To UBound(bytes) Step 2
        checked = checked + 1
        If bytes(i) = 0 Then nullCount = nullCount + 1
        If checked >= 200 Then Exit For
    Next i

    If checked > 0 Then BytesLookLikeUtf16Le = (nullCount * 2 >= checked)

Done:
End Function

Private Function SplitLines(ByVal text As String) As Variant
    text = Replace(text, vbCrLf, vbLf)
    text = Replace(text, vbCr, vbLf)
    SplitLines = Split(text, vbLf)
End Function

Private Function TokenAt(ByVal tokens As Variant, ByVal index As Long) As String
    On Error GoTo Done
    If IsArray(tokens) Then
        If index >= LBound(tokens) And index <= UBound(tokens) Then
            TokenAt = CStr(tokens(index))
        End If
    End If
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
    value = Trim$(value)
    IsHttpsUrl = (LCase$(Left$(value, 8)) = "https://")
End Function

Private Function NormalizeUrl(ByVal value As String) As String
    value = Replace(value, "\", "/")
    NormalizeUrl = UrlDecode(value)
End Function

Private Function UrlStartsWith(ByVal value As String, ByVal prefix As String) As Boolean
    value = TrimTrailingSlash(value)
    prefix = TrimTrailingSlash(prefix)

    If Len(value) < Len(prefix) Then Exit Function
    If StrComp(Left$(value, Len(prefix)), prefix, vbTextCompare) <> 0 Then Exit Function

    UrlStartsWith = (Len(value) = Len(prefix) Or Mid$(value, Len(prefix) + 1, 1) = "/")
End Function

Private Function UrlDecode(ByVal value As String) As String
    Dim result As String
    Dim bytes() As Byte
    Dim byteCount As Long
    Dim i As Long
    Dim ch As String

    i = 1
    Do While i <= Len(value)
        ch = Mid$(value, i, 1)

        If ch = "%" And i + 2 <= Len(value) And IsHexPair(Mid$(value, i + 1, 2)) Then
            byteCount = 0
            Erase bytes

            Do While i <= Len(value) - 2
                If Mid$(value, i, 1) <> "%" Then Exit Do
                If Not IsHexPair(Mid$(value, i + 1, 2)) Then Exit Do

                ReDim Preserve bytes(0 To byteCount)
                bytes(byteCount) = CByte("&H" & Mid$(value, i + 1, 2))
                byteCount = byteCount + 1
                i = i + 3
            Loop

            result = result & Utf8BytesToString(bytes)
        Else
            result = result & ch
            i = i + 1
        End If
    Loop

    UrlDecode = result
End Function

Private Function IsHexPair(ByVal value As String) As Boolean
    If Len(value) <> 2 Then Exit Function
    IsHexPair = IsHexChar(Left$(value, 1)) And IsHexChar(Right$(value, 1))
End Function

Private Function IsHexChar(ByVal value As String) As Boolean
    If Len(value) <> 1 Then Exit Function

    Dim code As Long
    code = AscW(value)

    IsHexChar = (code >= AscW("0") And code <= AscW("9")) _
                Or (code >= AscW("A") And code <= AscW("F")) _
                Or (code >= AscW("a") And code <= AscW("f"))
End Function

Private Function Utf8BytesToString(ByRef bytes() As Byte) As String
    On Error GoTo Done

    Dim i As Long
    Dim cp As Long
    Dim b0 As Long

    i = LBound(bytes)
    Do While i <= UBound(bytes)
        b0 = bytes(i)

        If b0 < &H80 Then
            cp = b0
            i = i + 1
        ElseIf (b0 And &HE0) = &HC0 And i + 1 <= UBound(bytes) Then
            cp = ((b0 And &H1F) * &H40) + (bytes(i + 1) And &H3F)
            i = i + 2
        ElseIf (b0 And &HF0) = &HE0 And i + 2 <= UBound(bytes) Then
            cp = ((b0 And &HF) * &H1000) _
                 + ((bytes(i + 1) And &H3F) * &H40) _
                 + (bytes(i + 2) And &H3F)
            i = i + 3
        ElseIf (b0 And &HF8) = &HF0 And i + 3 <= UBound(bytes) Then
            cp = ((b0 And &H7) * &H40000) _
                 + ((bytes(i + 1) And &H3F) * &H1000) _
                 + ((bytes(i + 2) And &H3F) * &H40) _
                 + (bytes(i + 3) And &H3F)
            i = i + 4
        Else
            cp = AscW("?")
            i = i + 1
        End If

        Utf8BytesToString = Utf8BytesToString & CodePointToString(cp)
    Loop

Done:
End Function

Private Function CodePointToString(ByVal codePoint As Long) As String
    If codePoint < 0 Then Exit Function

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

Private Function CombineUrl(ByVal baseUrl As String, ByVal childUrl As String) As String
    baseUrl = TrimTrailingSlash(NormalizeUrl(baseUrl))
    childUrl = TrimLeadingSlash(NormalizeUrl(childUrl))

    If Len(baseUrl) = 0 Then
        CombineUrl = childUrl
    ElseIf Len(childUrl) = 0 Then
        CombineUrl = baseUrl
    Else
        CombineUrl = baseUrl & "/" & childUrl
    End If
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

Private Function TrimLeadingSlash(ByVal value As String) As String
    value = Trim$(value)
    Do While Len(value) > 0 And Left$(value, 1) = "/"
        value = Mid$(value, 2)
    Loop
    TrimLeadingSlash = value
End Function

Private Function LastPathPart(ByVal value As String) As String
    value = TrimTrailingBackslash(value)
    LastPathPart = Mid$(value, InStrRev(value, PATH_SEP) + 1)
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

Private Function ClampLong(ByVal value As Long, ByVal lowerBound As Long, ByVal upperBound As Long) As Long
    If value < lowerBound Then
        ClampLong = lowerBound
    ElseIf value > upperBound Then
        ClampLong = upperBound
    Else
        ClampLong = value
    End If
End Function

Private Sub AddKeyedValue(ByVal coll As Collection, ByVal key As String, ByVal value As String)
    If Len(key) = 0 Or Len(value) = 0 Then Exit Sub
    On Error Resume Next
    coll.Add value, key
    On Error GoTo 0
End Sub

Private Function KeyedValue(ByVal coll As Collection, ByVal key As String) As String
    On Error Resume Next
    KeyedValue = CStr(coll(key))
    On Error GoTo 0
End Function

Private Function UBoundSafe(ByRef values As Variant) As Long
    On Error GoTo Fail
    UBoundSafe = UBound(values)
    Exit Function
Fail:
    UBoundSafe = -1
End Function
