Attribute VB_Name = "TableRecordEditor"
Option Explicit

Private m_openForms As Collection

Public Sub OpenActiveTableRecordEditor()
    Dim tbl As ListObject
    Set tbl = TableRecordEditorFindActiveTable()

    If tbl Is Nothing Then
        MsgBox "Select a cell inside an Excel table first.", vbExclamation, "Table Record Editor"
        Exit Sub
    End If

    TableRecordEditorOpen tbl
End Sub

Public Sub OpenNamedTableRecordEditor(Optional ByVal tableName As String = "")
    Dim tbl As ListObject

    If Len(Trim$(tableName)) = 0 Then
        Set tbl = TableRecordEditorFindActiveTable()
    Else
        Set tbl = TableRecordEditorFindTableByName(ActiveWorkbook, tableName)
    End If

    If tbl Is Nothing Then
        MsgBox "Table not found.", vbExclamation, "Table Record Editor"
        Exit Sub
    End If

    TableRecordEditorOpen tbl
End Sub

Public Sub TableRecordEditorOpen(ByVal tbl As ListObject)
    If tbl Is Nothing Then Err.Raise 5, "TableRecordEditorOpen", "Table is required."

    If m_openForms Is Nothing Then Set m_openForms = New Collection

    Dim frm As frmTableRecordEditor
    Set frm = New frmTableRecordEditor

    Dim key As String
    key = CStr(ObjPtr(frm))
    frm.HostKey = key
    frm.Init tbl

    m_openForms.Add frm, key
    frm.Show vbModeless
End Sub

Public Sub TableRecordEditorRelease(ByVal key As String)
    On Error Resume Next
    If Not m_openForms Is Nothing Then m_openForms.Remove key
    On Error GoTo 0
End Sub

Private Function TableRecordEditorFindActiveTable() As ListObject
    If ActiveWorkbook Is Nothing Then Exit Function
    If ActiveSheet Is Nothing Then Exit Function

    Dim rng As Range
    On Error Resume Next
    Set rng = Selection
    On Error GoTo 0

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim tbl As ListObject
    If Not rng Is Nothing Then
        For Each tbl In ws.ListObjects
            If Not Intersect(rng, tbl.Range) Is Nothing Then
                Set TableRecordEditorFindActiveTable = tbl
                Exit Function
            End If
        Next tbl
    End If

    If ws.ListObjects.Count = 1 Then
        Set TableRecordEditorFindActiveTable = ws.ListObjects(1)
    End If
End Function

Private Function TableRecordEditorFindTableByName(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    If wb Is Nothing Then Exit Function

    Dim ws As Worksheet
    Dim tbl As ListObject

    For Each ws In wb.Worksheets
        For Each tbl In ws.ListObjects
            If StrComp(tbl.Name, tableName, vbTextCompare) = 0 Then
                Set TableRecordEditorFindTableByName = tbl
                Exit Function
            End If
        Next tbl
    Next ws
End Function
