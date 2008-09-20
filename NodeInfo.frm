VERSION 5.00
Begin VB.Form NodeInfo 
   BackColor       =   &H0000FF00&
   Caption         =   "Node Information"
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   3360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2340
      TabIndex        =   4
      Top             =   1440
      Width           =   795
   End
   Begin VB.CommandButton OkButton 
      Caption         =   "OK"
      Height          =   495
      Left            =   180
      TabIndex        =   3
      Top             =   1440
      Width           =   675
   End
   Begin VB.CheckBox DeleteCheckBox 
      BackColor       =   &H008080FF&
      Caption         =   "Delete this node  ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   180
      TabIndex        =   2
      Top             =   840
      Width           =   2955
   End
   Begin VB.TextBox NodeName 
      Height          =   435
      Left            =   180
      TabIndex        =   0
      Top             =   300
      Width           =   2955
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "New name for this node"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   180
      TabIndex        =   1
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "NodeInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public manager As NetworkInfoManager

Private Sub CancelButton_Click()
    Me.Hide
End Sub

Private Sub OkButton_Click()
    manager.confirmChangesToNode
End Sub
