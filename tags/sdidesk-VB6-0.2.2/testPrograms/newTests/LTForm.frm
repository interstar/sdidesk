VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form LTForm 
   Caption         =   "Form1"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   ScaleHeight     =   8955
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Test NativeLinkWrapper"
      Height          =   675
      Left            =   6300
      TabIndex        =   6
      Top             =   180
      Width           =   1035
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Test StringTool"
      Height          =   675
      Left            =   5400
      TabIndex        =   5
      Top             =   180
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Test WikiToHtml"
      Height          =   675
      Left            =   4440
      TabIndex        =   4
      Top             =   180
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Test DoubleBrackets"
      Height          =   495
      Left            =   2580
      TabIndex        =   3
      Top             =   180
      Width           =   1755
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Test StandardLinkProcessor"
      Height          =   855
      Left            =   1320
      TabIndex        =   2
      Top             =   180
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test WikiMarkupGopher"
      Height          =   855
      Left            =   180
      TabIndex        =   1
      Top             =   180
      Width           =   975
   End
   Begin RichTextLib.RichTextBox RichText1 
      Height          =   7575
      Left            =   60
      TabIndex        =   0
      Top             =   1320
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   13361
      _Version        =   393217
      TextRTF         =   $"LTForm.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "LTForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub pt(s As String)
    LTForm.RichText1.text = LTForm.RichText1.text + vbCrLf + s
End Sub

Public Sub clearText()
    LTForm.RichText1.text = ""
End Sub

Private Sub Command1_Click()
    Call clearText
    pt testWmg("")
End Sub

Private Sub Command2_Click()
    Call clearText
    pt testSLP("")
End Sub

Private Sub Command3_Click()
    Call clearText
    pt testDB()
End Sub

Private Sub Command4_Click()
    Call clearText
    pt testW2H
End Sub

Private Sub Command5_Click()
    Call clearText
    pt testSt
End Sub

Private Sub Command6_Click()
    Call clearText
    pt testNlw
End Sub
