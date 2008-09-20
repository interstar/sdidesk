VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   555
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   8055
      Left            =   120
      TabIndex        =   0
      Top             =   1380
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   14208
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"ExTForm.frx":0000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
