VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E0D-00AA006002F3} frmTableRecordEditor
   Caption         =   "Table Record Editor"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10800
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmTableRecordEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_recordList As MSForms.ListBox
Private WithEvents m_splitter As MSForms.Label
Private WithEvents m_resizeHandle As MSForms.Label

Private m_topFrame As MSForms.Frame
Private m_fieldsFrame As MSForms.Frame
Private m_detailFrame As MSForms.Frame
Private m_detailLabel As MSForms.Label
Private m_detailText As MSForms.TextBox

Private m_table As ListObject
Private m_visibleColumns As Collection
Private m_fieldEditors As Collection
Private m_detailColumnIndex As Long
Private m_currentRowIndex As Long
Private m_topHeight As Single
Private m_suppressEvents As Boolean
Private m_layingOut As Boolean
Private m_hostKey As String

Private m_draggingSplitter As Boolean
Private m_splitStartY As Single
Private m_splitStartTopHeight As Single
Private m_draggingResize As Boolean
Private m_resizeStartX As Single
Private m_resizeStartY As Single
Private m_resizeStartWidth As Single
Private m_resizeStartHeight As Single

Private Const M As Single = 6
Private Const LABEL_W As Single = 112
Private Const ROW_H As Single = 22
Private Const SPLITTER_H As Single = 7
Private Const HANDLE_SIZE As Single = 14
Private Const MIN_FORM_W As Single = 420
Private Const MIN_FORM_H As Single = 300

Public Property Let HostKey(ByVal value As String)
    m_hostKey = value
End Property

Public Sub Init(ByVal tbl As ListObject)
    If tbl Is Nothing Then Err.Raise 5, "frmTableRecordEditor.Init", "Table is required."

    Set m_table = tbl
    Set m_visibleColumns = New Collection
    Set m_fieldEditors = New Collection

    CollectVisibleColumns
    If m_visibleColumns.Count = 0 Then Err.Raise 5, "frmTableRecordEditor.Init", "The table has no visible columns."

    m_detailColumnIndex = CLng(m_visibleColumns(m_visibleColumns.Count))
    m_topHeight = 245

    Me.Caption = "Table Record Editor - " & m_table.Name
    Me.Width = 720
    Me.Height = 520
    Me.BackColor = &HFFFFFF

    BuildShellControls
    BuildFieldControls
    PopulateRecordList
    LayoutControls

    If GetRecordCount() > 0 Then
        m_recordList.ListIndex = 0
        LoadRecord 1
    End If
End Sub

Private Sub CollectVisibleColumns()
    Dim i As Long
    For i = 1 To m_table.ListColumns.Count
        If Not m_table.ListColumns(i).Range.EntireColumn.Hidden Then
            m_visibleColumns.Add i
        End If
    Next i
End Sub

Private Sub BuildShellControls()
    Set m_topFrame = Me.Controls.Add("Forms.Frame.1", "fraTop", True)
    m_topFrame.Caption = ""
    m_topFrame.SpecialEffect = fmSpecialEffectFlat

    Set m_fieldsFrame = m_topFrame.Controls.Add("Forms.Frame.1", "fraFields", True)
    m_fieldsFrame.Caption = ""
    m_fieldsFrame.SpecialEffect = fmSpecialEffectFlat
    m_fieldsFrame.ScrollBars = fmScrollBarsVertical
    m_fieldsFrame.KeepScrollBarsVisible = fmScrollBarsVertical

    Set m_detailFrame = m_topFrame.Controls.Add("Forms.Frame.1", "fraDetail", True)
    m_detailFrame.Caption = ""
    m_detailFrame.SpecialEffect = fmSpecialEffectFlat

    Set m_splitter = Me.Controls.Add("Forms.Label.1", "lblSplitter", True)
    m_splitter.Caption = ""
    m_splitter.BackColor = &HD0D0D0
    m_splitter.SpecialEffect = fmSpecialEffectFlat
    m_splitter.BorderStyle = fmBorderStyleNone
    m_splitter.MousePointer = 7

    Set m_recordList = Me.Controls.Add("Forms.ListBox.1", "lstRecords", True)
    m_recordList.Font.Name = "Meiryo UI"
    m_recordList.Font.Size = 9
    m_recordList.IntegralHeight = False
    m_recordList.MultiSelect = fmMultiSelectSingle

    Set m_resizeHandle = Me.Controls.Add("Forms.Label.1", "lblResizeHandle", True)
    m_resizeHandle.Caption = ""
    m_resizeHandle.BackColor = &HD0D0D0
    m_resizeHandle.SpecialEffect = fmSpecialEffectFlat
    m_resizeHandle.BorderStyle = fmBorderStyleNone
    m_resizeHandle.MousePointer = 8
End Sub

Private Sub BuildFieldControls()
    Dim y As Single
    y = M

    Dim i As Long
    For i = 1 To m_visibleColumns.Count
        Dim columnIndex As Long
        columnIndex = CLng(m_visibleColumns(i))

        If columnIndex <> m_detailColumnIndex Then
            AddFieldRow columnIndex, y
            y = y + ROW_H + M
        End If
    Next i

    If y < m_fieldsFrame.Height Then y = m_fieldsFrame.Height + M
    m_fieldsFrame.ScrollHeight = y + M

    Set m_detailLabel = m_detailFrame.Controls.Add("Forms.Label.1", "lblDetail", True)
    m_detailLabel.Caption = m_table.ListColumns(m_detailColumnIndex).Name
    m_detailLabel.Font.Name = "Meiryo UI"
    m_detailLabel.Font.Size = 9

    Set m_detailText = m_detailFrame.Controls.Add("Forms.TextBox.1", "txtDetail", True)
    m_detailText.Font.Name = "Meiryo UI"
    m_detailText.Font.Size = 9
    m_detailText.MultiLine = True
    m_detailText.EnterKeyBehavior = True
    m_detailText.WordWrap = True
    m_detailText.ScrollBars = fmScrollBarsVertical
    ConfigureTextBoxFromColumn m_detailText, m_detailColumnIndex, True

    Dim detailEditor As TableRecordEditorField
    Set detailEditor = New TableRecordEditorField
    detailEditor.Bind Me, m_detailText, m_detailColumnIndex, True
    m_fieldEditors.Add detailEditor
End Sub

Private Sub AddFieldRow(ByVal columnIndex As Long, ByVal y As Single)
    Dim safeName As String
    safeName = CStr(columnIndex)

    Dim lbl As MSForms.Label
    Set lbl = m_fieldsFrame.Controls.Add("Forms.Label.1", "lblField" & safeName, True)
    lbl.Caption = m_table.ListColumns(columnIndex).Name
    lbl.Left = M
    lbl.Top = y + 3
    lbl.Width = LABEL_W
    lbl.Height = ROW_H
    lbl.Font.Name = "Meiryo UI"
    lbl.Font.Size = 9

    Dim txt As MSForms.TextBox
    Set txt = m_fieldsFrame.Controls.Add("Forms.TextBox.1", "txtField" & safeName, True)
    txt.Left = M + LABEL_W
    txt.Top = y
    txt.Height = ROW_H
    txt.Font.Name = "Meiryo UI"
    txt.Font.Size = 9
    txt.MultiLine = False
    ConfigureTextBoxFromColumn txt, columnIndex, False

    Dim editor As TableRecordEditorField
    Set editor = New TableRecordEditorField
    editor.Bind Me, txt, columnIndex, False
    m_fieldEditors.Add editor
End Sub

Private Sub ConfigureTextBoxFromColumn(ByVal txt As MSForms.TextBox, ByVal columnIndex As Long, ByVal isDetail As Boolean)
    If isDetail Then
        txt.TextAlign = fmTextAlignLeft
    ElseIf ShouldRightAlign(columnIndex) Then
        txt.TextAlign = fmTextAlignRight
    Else
        txt.TextAlign = fmTextAlignLeft
    End If
End Sub

Private Function ShouldRightAlign(ByVal columnIndex As Long) As Boolean
    Dim fmt As String
    fmt = ColumnNumberFormat(columnIndex)

    If IsNumericFormat(fmt) Then
        ShouldRightAlign = True
        Exit Function
    End If

    If Not m_table.DataBodyRange Is Nothing Then
        Dim v As Variant
        v = m_table.DataBodyRange.Cells(1, columnIndex).Value
        If IsNumeric(v) Then ShouldRightAlign = True
    End If
End Function

Private Sub LayoutControls()
    If m_topFrame Is Nothing Then Exit Sub
    If m_layingOut Then Exit Sub

    m_layingOut = True
    On Error GoTo CleanUp

    Dim cw As Single
    Dim ch As Single
    cw = Me.InsideWidth
    ch = Me.InsideHeight

    If cw < MIN_FORM_W Then cw = MIN_FORM_W
    If ch < MIN_FORM_H Then ch = MIN_FORM_H

    m_topHeight = ClampS(m_topHeight, 95, ch - 115)

    m_topFrame.Left = M
    m_topFrame.Top = M
    m_topFrame.Width = cw - M * 2
    m_topFrame.Height = m_topHeight

    m_splitter.Left = M
    m_splitter.Top = m_topFrame.Top + m_topFrame.Height + 3
    m_splitter.Width = cw - M * 2
    m_splitter.Height = SPLITTER_H

    m_recordList.Left = M
    m_recordList.Top = m_splitter.Top + m_splitter.Height + 4
    m_recordList.Width = cw - M * 2
    m_recordList.Height = ch - m_recordList.Top - M - HANDLE_SIZE
    If m_recordList.Height < 45 Then m_recordList.Height = 45

    m_resizeHandle.Left = cw - HANDLE_SIZE
    m_resizeHandle.Top = ch - HANDLE_SIZE
    m_resizeHandle.Width = HANDLE_SIZE
    m_resizeHandle.Height = HANDLE_SIZE

    Dim innerW As Single
    Dim innerH As Single
    innerW = m_topFrame.Width - M * 2
    innerH = m_topFrame.Height - M * 2

    Dim leftW As Single
    leftW = Int(innerW * 0.66)
    If leftW < 210 Then leftW = 210
    If leftW > innerW - 130 Then leftW = innerW - 130

    m_fieldsFrame.Left = M
    m_fieldsFrame.Top = M
    m_fieldsFrame.Width = leftW
    m_fieldsFrame.Height = innerH

    m_detailFrame.Left = m_fieldsFrame.Left + m_fieldsFrame.Width + M
    m_detailFrame.Top = M
    m_detailFrame.Width = innerW - m_fieldsFrame.Width - M
    m_detailFrame.Height = innerH
    If m_detailFrame.Width < 120 Then m_detailFrame.Width = 120

    LayoutFieldTextBoxes
    LayoutDetailBox

CleanUp:
    m_layingOut = False
End Sub

Private Sub LayoutFieldTextBoxes()
    If m_fieldEditors Is Nothing Then Exit Sub

    Dim editor As TableRecordEditorField
    For Each editor In m_fieldEditors
        If Not editor.IsDetail Then
            editor.TextBox.Width = MaxS(60, m_fieldsFrame.Width - LABEL_W - M * 3)
        End If
    Next editor
End Sub

Private Sub LayoutDetailBox()
    If m_detailLabel Is Nothing Then Exit Sub

    m_detailLabel.Left = M
    m_detailLabel.Top = M
    m_detailLabel.Width = m_detailFrame.Width - M * 2
    m_detailLabel.Height = 18

    m_detailText.Left = M
    m_detailText.Top = m_detailLabel.Top + m_detailLabel.Height + 4
    m_detailText.Width = m_detailFrame.Width - M * 2
    m_detailText.Height = m_detailFrame.Height - m_detailText.Top - M
    If m_detailText.Height < 30 Then m_detailText.Height = 30
End Sub

Private Sub PopulateRecordList()
    m_recordList.Clear
    m_recordList.ColumnCount = m_visibleColumns.Count
    m_recordList.ColumnWidths = BuildColumnWidths()

    Dim rowCount As Long
    rowCount = GetRecordCount()
    If rowCount = 0 Then Exit Sub

    Dim r As Long
    Dim c As Long
    For r = 1 To rowCount
        m_recordList.AddItem DisplayCell(r, CLng(m_visibleColumns(1)))
        For c = 2 To m_visibleColumns.Count
            m_recordList.List(r - 1, c - 1) = DisplayCell(r, CLng(m_visibleColumns(c)))
        Next c
    Next r
End Sub

Private Function BuildColumnWidths() As String
    Dim parts() As String
    ReDim parts(1 To m_visibleColumns.Count)

    Dim i As Long
    For i = 1 To m_visibleColumns.Count
        Dim columnIndex As Long
        columnIndex = CLng(m_visibleColumns(i))

        Dim w As Single
        On Error Resume Next
        w = m_table.ListColumns(columnIndex).Range.Width
        On Error GoTo 0
        If w <= 0 Then w = 80

        w = ClampS(w, 55, 190)
        If columnIndex = m_detailColumnIndex Then w = ClampS(w, 90, 260)
        parts(i) = CStr(CLng(w)) & " pt"
    Next i

    BuildColumnWidths = Join(parts, ";")
End Function

Private Function GetRecordCount() As Long
    If m_table Is Nothing Then Exit Function
    If m_table.DataBodyRange Is Nothing Then Exit Function
    GetRecordCount = m_table.DataBodyRange.Rows.Count
End Function

Private Sub LoadRecord(ByVal rowIndex As Long)
    If rowIndex < 1 Or rowIndex > GetRecordCount() Then Exit Sub

    m_suppressEvents = True
    m_currentRowIndex = rowIndex

    Dim editor As TableRecordEditorField
    For Each editor In m_fieldEditors
        editor.SetText DisplayCell(rowIndex, editor.ColumnIndex)
    Next editor

    m_suppressEvents = False
End Sub

Public Sub FieldBoxChanged(ByVal columnIndex As Long, ByVal textValue As String)
    If m_suppressEvents Then Exit Sub
    If m_currentRowIndex < 1 Then Exit Sub
    If m_table Is Nothing Then Exit Sub
    If m_table.DataBodyRange Is Nothing Then Exit Sub

    On Error GoTo ErrHandler

    Dim v As Variant
    v = ParseInputValue(textValue, columnIndex)
    m_table.DataBodyRange.Cells(m_currentRowIndex, columnIndex).Value = v
    UpdateRecordListCell m_currentRowIndex, columnIndex
    Exit Sub

ErrHandler:
    MsgBox "Failed to write the edited value to the table." & vbCrLf & Err.Description, vbExclamation, "Table Record Editor"
    m_suppressEvents = True
    LoadRecord m_currentRowIndex
    m_suppressEvents = False
End Sub

Private Sub UpdateRecordListCell(ByVal rowIndex As Long, ByVal columnIndex As Long)
    If m_recordList Is Nothing Then Exit Sub
    If rowIndex < 1 Or rowIndex > m_recordList.ListCount Then Exit Sub

    Dim visibleIndex As Long
    visibleIndex = VisibleColumnOffset(columnIndex)
    If visibleIndex < 0 Then Exit Sub

    m_recordList.List(rowIndex - 1, visibleIndex) = DisplayCell(rowIndex, columnIndex)
End Sub

Private Function VisibleColumnOffset(ByVal columnIndex As Long) As Long
    Dim i As Long
    For i = 1 To m_visibleColumns.Count
        If CLng(m_visibleColumns(i)) = columnIndex Then
            VisibleColumnOffset = i - 1
            Exit Function
        End If
    Next i
    VisibleColumnOffset = -1
End Function

Private Function DisplayCell(ByVal rowIndex As Long, ByVal columnIndex As Long) As String
    If m_table.DataBodyRange Is Nothing Then Exit Function

    Dim cell As Range
    Set cell = m_table.DataBodyRange.Cells(rowIndex, columnIndex)

    Dim v As Variant
    v = cell.Value

    If IsError(v) Then
        DisplayCell = cell.Text
        Exit Function
    End If

    If IsNull(v) Or IsEmpty(v) Then
        DisplayCell = ""
        Exit Function
    End If

    If Len(CStr(v)) = 0 Then
        DisplayCell = ""
        Exit Function
    End If

    Dim fmt As String
    fmt = CStr(cell.NumberFormat)

    If fmt <> "General" And (IsDate(v) Or IsNumeric(v)) Then
        On Error Resume Next
        DisplayCell = Application.WorksheetFunction.Text(v, fmt)
        If Err.Number = 0 Then
            On Error GoTo 0
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
    End If

    DisplayCell = CStr(v)
End Function

Private Function ParseInputValue(ByVal textValue As String, ByVal columnIndex As Long) As Variant
    If Len(textValue) = 0 Then
        ParseInputValue = vbNullString
        Exit Function
    End If

    Dim fmt As String
    fmt = ColumnNumberFormat(columnIndex)

    If IsDateFormat(fmt) And IsDate(textValue) Then
        ParseInputValue = CDate(textValue)
        Exit Function
    End If

    Dim normalized As String
    normalized = NormalizeNumberText(textValue)
    If (IsNumericFormat(fmt) Or CellCurrentlyNumeric(columnIndex)) And IsNumeric(normalized) Then
        ParseInputValue = CDbl(normalized)
        Exit Function
    End If

    ParseInputValue = textValue
End Function

Private Function NormalizeNumberText(ByVal textValue As String) As String
    Dim s As String
    s = Trim$(textValue)
    s = Replace(s, ",", "")
    s = Replace(s, "$", "")
    s = Replace(s, ChrW$(&HA5), "")
    s = Replace(s, " ", "")
    NormalizeNumberText = s
End Function

Private Function CellCurrentlyNumeric(ByVal columnIndex As Long) As Boolean
    If m_currentRowIndex < 1 Then Exit Function
    If m_table.DataBodyRange Is Nothing Then Exit Function

    Dim v As Variant
    v = m_table.DataBodyRange.Cells(m_currentRowIndex, columnIndex).Value
    If IsError(v) Then Exit Function
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    CellCurrentlyNumeric = IsNumeric(v)
End Function

Private Function ColumnNumberFormat(ByVal columnIndex As Long) As String
    On Error Resume Next

    Dim fmt As Variant
    If Not m_table.DataBodyRange Is Nothing Then
        fmt = m_table.DataBodyRange.Cells(1, columnIndex).NumberFormat
    End If

    If Err.Number <> 0 Or IsEmpty(fmt) Or IsNull(fmt) Then
        Err.Clear
        fmt = m_table.ListColumns(columnIndex).Range.NumberFormat
    End If

    If Err.Number <> 0 Or IsEmpty(fmt) Or IsNull(fmt) Then
        ColumnNumberFormat = "General"
    Else
        ColumnNumberFormat = CStr(fmt)
    End If

    On Error GoTo 0
End Function

Private Function IsNumericFormat(ByVal fmt As String) As Boolean
    fmt = LCase$(fmt)
    IsNumericFormat = (InStr(fmt, "0") > 0 Or InStr(fmt, "#") > 0 Or InStr(fmt, "currency") > 0 Or InStr(fmt, "accounting") > 0)
End Function

Private Function IsDateFormat(ByVal fmt As String) As Boolean
    fmt = LCase$(fmt)
    IsDateFormat = (InStr(fmt, "yy") > 0 Or InStr(fmt, "dd") > 0 Or InStr(fmt, "yyyy") > 0 Or InStr(fmt, "m/d") > 0 Or InStr(fmt, "d/m") > 0)
End Function

Private Function ClampS(ByVal value As Single, ByVal minValue As Single, ByVal maxValue As Single) As Single
    If value < minValue Then
        ClampS = minValue
    ElseIf value > maxValue Then
        ClampS = maxValue
    Else
        ClampS = value
    End If
End Function

Private Function MaxS(ByVal a As Single, ByVal b As Single) As Single
    If a > b Then MaxS = a Else MaxS = b
End Function

Private Sub m_recordList_Click()
    If m_recordList.ListIndex < 0 Then Exit Sub
    LoadRecord m_recordList.ListIndex + 1
End Sub

Private Sub m_splitter_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button <> 1 Then Exit Sub
    m_draggingSplitter = True
    m_splitStartY = Y
    m_splitStartTopHeight = m_topHeight
End Sub

Private Sub m_splitter_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not m_draggingSplitter Then Exit Sub
    m_topHeight = m_splitStartTopHeight + (Y - m_splitStartY)
    LayoutControls
End Sub

Private Sub m_splitter_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    m_draggingSplitter = False
End Sub

Private Sub m_resizeHandle_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button <> 1 Then Exit Sub
    m_draggingResize = True
    m_resizeStartX = X
    m_resizeStartY = Y
    m_resizeStartWidth = Me.Width
    m_resizeStartHeight = Me.Height
End Sub

Private Sub m_resizeHandle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not m_draggingResize Then Exit Sub

    Dim newW As Single
    Dim newH As Single
    newW = m_resizeStartWidth + (X - m_resizeStartX)
    newH = m_resizeStartHeight + (Y - m_resizeStartY)

    If newW < MIN_FORM_W Then newW = MIN_FORM_W
    If newH < MIN_FORM_H Then newH = MIN_FORM_H

    Me.Width = newW
    Me.Height = newH
    LayoutControls
End Sub

Private Sub m_resizeHandle_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    m_draggingResize = False
End Sub

Private Sub UserForm_Resize()
    LayoutControls
End Sub

Private Sub UserForm_Terminate()
    If Len(m_hostKey) > 0 Then TableRecordEditorRelease m_hostKey
End Sub
