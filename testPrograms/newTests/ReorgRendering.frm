VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form ReorgRendering 
   Caption         =   "Form1"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Export Html"
      Height          =   435
      Left            =   4080
      TabIndex        =   4
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cooked"
      Height          =   435
      Left            =   2400
      TabIndex        =   3
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Prepared"
      Height          =   435
      Left            =   1320
      TabIndex        =   2
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Raw"
      Height          =   435
      Left            =   300
      TabIndex        =   1
      Top             =   60
      Width           =   915
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   7155
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   12621
      _Version        =   393217
      TextRTF         =   $"ReorgRendering.frx":0000
   End
End
Attribute VB_Name = "ReorgRendering"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim model As ModelLevel
Dim factory As SdiDeskConfigurationFactory
Dim p As Page

Private Sub Command1_Click()
    Call clear
    Set p = model.getWikiAnnotatedDataStore.store.loadRaw("Links")
    
    pt p.pageName
    pt p.createdDate
    pt p.lastEdited
    pt p.raw
End Sub

Private Sub Command2_Click()
    Call clear
    Set p = model.getWikiAnnotatedDataStore.store.loadRaw("Links")
    Call p.prepare(factory.getPagePreparer, True)
    pt p.prepared
End Sub

Private Sub Command3_Click()
    Call clear
    Set p = model.getWikiAnnotatedDataStore.store.loadRaw("Links")
    Call p.cook(factory.getPagePreparer, factory.getNativePageCooker, True)
    pt p.cooked
End Sub


Private Sub Form_Load()
    Set factory = New SdiDeskConfigurationFactory
    Set model = factory.getModelLevel
End Sub

Private Sub pt(s As String)
   Text1.text = Text1.text + s + vbCrLf
End Sub

Private Sub clear()
    Text1.text = ""
End Sub
