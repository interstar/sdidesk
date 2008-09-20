VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form utlaaForm 
   Caption         =   "Form1"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Test VCollection"
      Height          =   555
      Left            =   1680
      TabIndex        =   2
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test LineArgAnalyser"
      Height          =   555
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   1515
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4515
      Left            =   60
      TabIndex        =   0
      Top             =   720
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   7964
      _Version        =   393217
      TextRTF         =   $"utlaaForm.frx":0000
   End
End
Attribute VB_Name = "utlaaForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub clearText()
    utlaaForm.RichTextBox1.Text = ""
End Sub

Public Sub pt(s As String)
    utlaaForm.RichTextBox1.Text = utlaaForm.RichTextBox1.Text + s + vbCrLf
End Sub

Private Sub Command1_Click()
    clearText
    pt testPL()
End Sub


Private Sub Command2_Click()
    clearText
    pt testVC
End Sub
