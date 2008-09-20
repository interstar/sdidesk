VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   675
      Left            =   6780
      TabIndex        =   5
      Top             =   3480
      Width           =   1035
   End
   Begin VB.TextBox EditableCell 
      Height          =   435
      Left            =   6720
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   300
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3615
      Left            =   3120
      TabIndex        =   1
      Top             =   4320
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   6376
      _Version        =   393217
      TextRTF         =   $"testTable.frx":0000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Edit"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   4380
      Width           =   1395
   End
   Begin VB.Frame ViewFrame 
      Caption         =   "Frame1"
      Height          =   4035
      Left            =   180
      TabIndex        =   3
      Top             =   180
      Width           =   6315
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3735
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   6588
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public t As Table
Public td As TableDisplay

Private Sub Command1_Click()
    Call td.fillFromTable(t)
End Sub

Private Sub Command2_Click()
    Set t = td.toTable()
    RichTextBox1.Text = t.spitAsPrettyPersist()
    
End Sub

Private Sub Form_Load()
    Dim x As String
    Set t = New Table
    Set td = New TableDisplay
    x = "a,, b,, c" & vbCrLf & "____" & vbCrLf & _
    "1,, 2,, 3" & vbCrLf & "4,, 5,, 6" & vbCrLf & _
    vbCrLf & vbCrLf & " a comment "
    
    RichTextBox1.Text = x
    Call t.parseFromDoubleCommaString(x)

    Call td.init(MSFlexGrid1, EditableCell, RichTextBox1, Form1)
    Call td.fillFromTable(t)
End Sub

Private Sub MSFlexGrid1_DblClick()
    Call Me.td.cellEdit
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, shift As Integer, x As Single, y As Single)
    Call td.startRange
End Sub

Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call Me.td.cellEdit
    End If
    If KeyCode = vbKeyDelete Then
        TableEditor.Text = ""
    End If
End Sub


Private Sub EditableCell_KeyDown(KeyCode As Integer, shift As Integer)
    Call td.editableCellKeyDown(KeyCode)
End Sub

Private Sub EditableCell_LostFocus()
    EditableCell.Visible = False
End Sub

Private Sub EditorButton_Click()
   Call controller.actionEdit(MagicNotebook.getControllableModel.getCurrentPage.pageName, False)
End Sub

