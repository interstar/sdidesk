VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form1"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog FileChooser 
      Left            =   180
      Top             =   2340
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   4560
      TabIndex        =   7
      Top             =   5880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "View Current PageName"
      Height          =   555
      Left            =   3660
      TabIndex        =   6
      Top             =   6360
      Width           =   1035
   End
   Begin VB.CommandButton Command3 
      Caption         =   "View Param Page Vars"
      Height          =   555
      Left            =   2520
      TabIndex        =   5
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "View Arguments"
      Height          =   555
      Left            =   1380
      TabIndex        =   4
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "View Command"
      Height          =   555
      Left            =   180
      TabIndex        =   3
      Top             =   6360
      Width           =   1155
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3795
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   5235
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "I export pages, or sets of pages as flat HTML."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   5235
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HtmlExporter by Phil Jones"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   26.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   1155
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5235
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private wpe As WebPageExporter
Private laa As LineArgAnaliser

Public Function outFile() As String
    FileChooser.CancelError = True
    FileChooser.InitDir = wpe.getPsi
    Dim doit As Boolean
    doit = False
   
    On Error GoTo Cancelled ' most likely cause of error
        Call FileChooser.ShowSave
        chooseFile = FileChooser.fileName
        Exit Function
        
Cancelled:
     ' cancelled (most likely)
    chooseDirectory = ""
End Function

Private Sub Command1_Click()
    Form1.Label3.Caption = "Command " & vbCrLf & command
End Sub

Private Sub Command2_Click()
    Set laa = New LineArgAnaliser
    laa.analise (command)
    Form1.Label3.Caption = "Command Args" & vbCrLf & laa.toString
End Sub

Private Sub Command3_Click()
    Set laa = New LineArgAnaliser
    laa.analise (command)
    
    Dim store As New FileSystemPageStore
    Dim path As String, paramPage As String
    Dim prep As PagePreparer
    Set prep = POLICY_getFactory().getPagePreparer
    
    path = laa.argDict.Item("psi")
 '   path = "C:\PROGRAM FILES\MICROSOFT VISUAL STUDIO\VB98\PROJECTS\DESKTOPWIKI\"
    paramPage = laa.argDict.Item("param")
 '   paramPage = "Exports/Dummy"
    path = store.pathFromFileName(path)
    path = store.ensureTrailingSlash(path)
    
    store.setDataDirectory (path)
    
    Dim p As page
    Set p = store.asPageStore.loadRaw(paramPage)
    Call p.prepare(prep, False)
    Form1.Label3.Caption = "ParamPage Variables " & vbCrLf & p.varsToString
    
End Sub

Private Sub Command4_Click()
    Set laa = New LineArgAnaliser
    laa.analise (command)
    
    Form1.Label3.Caption = "PageName = " & laa.argDict.Item("page")
 
End Sub

Private Sub Form_Load()
    Set wpe = New WebPageExporter
    
End Sub
