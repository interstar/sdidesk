VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TableDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This is the code which manages a TableEditor
' Doesn't actually create the Editor (that's given by
' the main form) but does know about and manipulate it

Private Editor As MSFlexGrid
Private EditableCell As TextBox
Private RawText As RichTextBox

Public mainForm As Form

Public noRows As Integer
Public noCols As Integer

Public startRow As Integer ' for selecting a range
Public startCol As Integer

Public comment As String

Public Sub init(te As MSFlexGrid, ec As TextBox, rt As RichTextBox, mf As Form)
    Set Editor = te
    Set EditableCell = ec
    Set RawText = rt
    Set mainForm = mf
End Sub


Public Sub fillFromTable(t As table)
    Dim i As Integer, j As Integer
    Me.noRows = t.noRows
    Me.noCols = t.noCols
    
    Editor.rows = t.noRows + 2
    Editor.Cols = t.noCols + 2
    Editor.Row = 0
    For j = 0 To t.noCols
        Editor.col = j + 1
        Editor.text = CStr(t.atHeader(j))
    Next j
    For i = 0 To t.noRows
        For j = 0 To t.noCols
            Editor.Row = i + 1
            Editor.col = j + 1
            Editor.text = CStr(t.at(i, j))
        Next j
    Next i
    
    comment = t.comment
    
End Sub

Public Function toTable() As table
    Dim i As Integer, j As Integer
    Dim t As New table
    
    Call t.setUp(Me.noRows, Me.noCols)
    
    Editor.Row = 0
    For j = 0 To t.noCols
        Editor.col = j + 1
        Call t.setHeader(j, Editor.text)
    Next j
    
    For i = 0 To t.noRows
        For j = 0 To t.noCols
            Editor.col = j + 1
            Editor.Row = i + 1
            Call t.putIn(i, j, Editor.text)
        Next j
    Next i
    t.comment = comment
    Set toTable = t
End Function

Public Function updatePage(p As Page) As Page
    Dim t As table
    Set t = toTable()
    p.raw = t.spitAsPrettyPersist
    
    Set t = Nothing
End Function

Public Sub cellEdit()
    EditableCell.Visible = True
    EditableCell.Width = Editor.CellWidth
    EditableCell.Height = Editor.CellHeight
    EditableCell.Top = Editor.CellTop + Editor.Top
    
    EditableCell.Left = Editor.CellLeft + Editor.Left + Me.mainForm.ViewFrame.Left
    EditableCell.text = Editor.text
    EditableCell.SelStart = 0
    EditableCell.SelLength = Len(EditableCell.text)
    EditableCell.ZOrder
    EditableCell.SetFocus
    
End Sub

Public Sub startRange()
    startRow = Editor.Row
    startCol = Editor.col
End Sub


Public Sub editableCellKeyDown(KeyCode As Integer)
    If KeyCode = vbKeyReturn Then
        Editor.text = EditableCell.text
        If Editor.Row = Editor.rows - 1 Then
            Editor.Row = Editor.Row
        Else
            Editor.Row = Editor.Row + 1
        End If
        Editor.SetFocus
        EditableCell.Visible = False
    End If
End Sub

'Public Sub changed(p As Page)
'     Dim s As String
'     Dim t As table
     
   '  p.raw = n.spitAsPrettyPersist
   '  Set p.myNetwork = New Network
    ' Call p.myNetwork.init(1, 200, 0.75)
   '  p.myNetwork.parseFromPrettyPersist (p.raw)
   '  Call draw(p.myNetwork, mode)
   '  RawText.text = p.myNetwork.spitAsPrettyPersist
   '  mainForm.MagicNotebook.changesSaved = False
'End Sub

Private Sub Class_Terminate()
    Set Editor = Nothing
    Set EditableCell = Nothing
    Set RawText = Nothing
    Set mainForm = Nothing
End Sub
