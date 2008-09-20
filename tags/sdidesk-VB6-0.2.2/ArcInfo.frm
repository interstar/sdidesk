VERSION 5.00
Begin VB.Form ArcInfo 
   BackColor       =   &H0080C0FF&
   Caption         =   "Arc Information"
   ClientHeight    =   2610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   3240
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox DeleteCheckBox 
      BackColor       =   &H008080FF&
      Caption         =   "Delete this arc ?"
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
      TabIndex        =   5
      Top             =   1440
      Width           =   2895
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   1980
      Width           =   795
   End
   Begin VB.CommandButton OkButton 
      Caption         =   "OK"
      Height          =   495
      Left            =   180
      TabIndex        =   3
      Top             =   1980
      Width           =   675
   End
   Begin VB.CheckBox ArcDirectionality 
      BackColor       =   &H0000FFFF&
      Caption         =   "Directional ?"
      CausesValidation=   0   'False
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
      TabIndex        =   2
      Top             =   900
      Width           =   2895
   End
   Begin VB.TextBox ArcName 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   2955
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "New name for this Arc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   3255
   End
End
Attribute VB_Name = "ArcInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public manager As NetworkInfoManager

Private Sub CancelButton_Click()
    Me.Hide
End Sub

Private Sub OkButton_Click()
    manager.confirmChangesToArc
End Sub
