VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableRecordEditorField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_textBox As MSForms.TextBox
Private m_parent As Object
Private m_columnIndex As Long
Private m_isDetail As Boolean

Public Sub Bind(ByVal parentForm As Object, ByVal textBox As MSForms.TextBox, ByVal columnIndex As Long, ByVal isDetail As Boolean)
    Set m_parent = parentForm
    Set m_textBox = textBox
    m_columnIndex = columnIndex
    m_isDetail = isDetail
End Sub

Public Property Get ColumnIndex() As Long
    ColumnIndex = m_columnIndex
End Property

Public Property Get IsDetail() As Boolean
    IsDetail = m_isDetail
End Property

Public Property Get TextBox() As MSForms.TextBox
    Set TextBox = m_textBox
End Property

Public Sub SetText(ByVal value As String)
    m_textBox.Text = value
    m_textBox.SelStart = 0
End Sub

Private Sub m_textBox_Change()
    If m_parent Is Nothing Then Exit Sub
    m_parent.FieldBoxChanged m_columnIndex, m_textBox.Text
End Sub

Private Sub m_textBox_Enter()
    On Error Resume Next
    m_textBox.SelStart = 0
    m_textBox.SelLength = Len(m_textBox.Text)
    On Error GoTo 0
End Sub

Private Sub Class_Terminate()
    Set m_textBox = Nothing
    Set m_parent = Nothing
End Sub
