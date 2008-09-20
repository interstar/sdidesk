VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form WADSMainForm 
   BackColor       =   &H008080FF&
   Caption         =   "SdiDesk - (Version 0.2.2  ... another week, another release)"
   ClientHeight    =   8535
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7950
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid TableEditor 
      Height          =   6255
      Left            =   120
      TabIndex        =   27
      Top             =   1200
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   11033
      _Version        =   393216
   End
   Begin VB.TextBox EditableCell 
      Height          =   360
      Left            =   2520
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   1500
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.DirListBox DirListBox 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   120
      TabIndex        =   24
      Top             =   1920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog DirectoryChooser 
      Left            =   0
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Choose a new directory"
      FileName        =   "none"
      InitDir         =   "App"
   End
   Begin VB.Frame CategoryFrame 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   60
      TabIndex        =   9
      Top             =   7620
      Width           =   7575
      Begin VB.CommandButton HistoryButton 
         BackColor       =   &H0080C0FF&
         Caption         =   "History"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   60
         Width           =   975
      End
      Begin VB.Label DateCreatedLabel 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Created : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   19
         Top             =   60
         Width           =   2655
      End
      Begin VB.Label DateLastEditedLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Edited"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   18
         Top             =   60
         Width           =   2775
      End
   End
   Begin VB.Frame ViewFrame 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   60
      TabIndex        =   8
      Top             =   1140
      Width           =   7635
      Begin SHDocVwCtl.WebBrowser HtmlView 
         Height          =   6015
         Left            =   60
         TabIndex        =   13
         Top             =   60
         Width           =   7395
         ExtentX         =   13044
         ExtentY         =   10610
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin VB.PictureBox NetworkCanvas 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         DrawWidth       =   2
         Height          =   6015
         Left            =   120
         MousePointer    =   2  'Cross
         ScaleHeight     =   5955
         ScaleWidth      =   7395
         TabIndex        =   16
         Top             =   120
         Visible         =   0   'False
         Width           =   7455
      End
      Begin RichTextLib.RichTextBox RawText 
         Height          =   6135
         Left            =   120
         TabIndex        =   12
         Top             =   60
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   10821
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"WADSMainForm.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame NavFrame 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8775
      Begin VB.CommandButton FindButton 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Find"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8100
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   60
         Width           =   615
      End
      Begin VB.CommandButton RecentButton 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Recent"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton GoButton 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Go!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   60
         Width           =   615
      End
      Begin VB.CommandButton AllButton 
         BackColor       =   &H00C0FFC0&
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2700
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   60
         Width           =   615
      End
      Begin VB.CommandButton StartPageButton 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1140
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   60
         Width           =   615
      End
      Begin VB.CommandButton BackButton 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Width           =   495
      End
      Begin VB.TextBox PageNameText 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3360
         TabIndex        =   2
         Top             =   60
         Width           =   4035
      End
      Begin VB.CommandButton ForwardButton 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   1
         Top             =   60
         Width           =   495
      End
   End
   Begin VB.Frame PageFrame 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   60
      TabIndex        =   4
      Top             =   600
      Width           =   8775
      Begin VB.CommandButton ExportButton 
         BackColor       =   &H0080FFFF&
         Caption         =   "Export"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5340
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   60
         Width           =   915
      End
      Begin VB.ComboBox HistoryList 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6420
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   60
         Width           =   2295
      End
      Begin VB.CommandButton NewNetworkButton 
         BackColor       =   &H0080FFFF&
         Caption         =   "New Net"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   60
         Width           =   975
      End
      Begin VB.CommandButton SaveButton 
         BackColor       =   &H0080FFFF&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1860
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton NewButton 
         BackColor       =   &H0080FFFF&
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton EditorButton 
         BackColor       =   &H0080FFFF&
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton RawButton 
         BackColor       =   &H0080FFFF&
         Caption         =   "Raw"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton PresentationButton 
         BackColor       =   &H0080FFFF&
         Caption         =   "Preview"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3420
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         Width           =   975
      End
   End
   Begin VB.Menu menuFile 
      Caption         =   "&File"
      NegotiatePosition=   1  'Left
      Begin VB.Menu menuNewPage 
         Caption         =   "&New Page"
         Shortcut        =   ^N
      End
      Begin VB.Menu menuNewNet 
         Caption         =   "New Network"
         Shortcut        =   ^M
      End
      Begin VB.Menu menuSavePage 
         Caption         =   "&Save Page"
         Shortcut        =   ^S
      End
      Begin VB.Menu menuSep1 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu menuDelete 
         Caption         =   "&Delete Page"
         Shortcut        =   ^D
      End
      Begin VB.Menu menuSep4 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu menuDirectoryChooser 
         Caption         =   "Change Directory"
         Shortcut        =   {F1}
      End
      Begin VB.Menu menuSepSDJI 
         Caption         =   "-"
      End
      Begin VB.Menu menuExit 
         Caption         =   "Exit"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu menuPage 
      Caption         =   "&Page"
      NegotiatePosition=   1  'Left
      Begin VB.Menu menuEdit 
         Caption         =   "&Edit"
         Shortcut        =   ^E
      End
      Begin VB.Menu menuPreview 
         Caption         =   "Preview"
         Shortcut        =   ^U
      End
      Begin VB.Menu menuRaw 
         Caption         =   "&Raw"
         Shortcut        =   ^R
      End
      Begin VB.Menu menuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu menuHistory 
         Caption         =   "&History"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu menuStandard 
      Caption         =   "&Navigate"
      NegotiatePosition=   1  'Left
      Begin VB.Menu menuBack 
         Caption         =   "Back"
         Shortcut        =   ^J
      End
      Begin VB.Menu menuForward 
         Caption         =   "Forward"
         Shortcut        =   ^K
      End
      Begin VB.Menu menuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu menuStart 
         Caption         =   "S&tart Page"
         Shortcut        =   ^T
      End
      Begin VB.Menu menuRecentChanges 
         Caption         =   "Recent Changes"
         Shortcut        =   ^B
      End
      Begin VB.Menu menuAll 
         Caption         =   "&All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu menuExport 
      Caption         =   "Export"
      Begin VB.Menu menuExportThisPageHtml 
         Caption         =   "Export This Page (as HTML)"
      End
      Begin VB.Menu menuShowExporters 
         Caption         =   "Show Exporters"
      End
      Begin VB.Menu menuShowExports 
         Caption         =   "Show Exports"
      End
      Begin VB.Menu menuShowCrawlers 
         Caption         =   "Show Crawlers"
      End
   End
   Begin VB.Menu menuSettings 
      Caption         =   "Settings"
      Begin VB.Menu menuInterMap 
         Caption         =   "InterMap"
      End
      Begin VB.Menu menuLinkTypeDefinitions 
         Caption         =   "Link Types"
      End
      Begin VB.Menu menuCrawlers 
         Caption         =   "Crawlers"
      End
      Begin VB.Menu menuExports 
         Caption         =   "Exports"
      End
      Begin VB.Menu menuSep7 
         Caption         =   "-"
      End
      Begin VB.Menu menuBackLinks 
         Caption         =   "BackLinks"
      End
      Begin VB.Menu menuSep8 
         Caption         =   "-"
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "&Help"
      NegotiatePosition=   3  'Right
      Begin VB.Menu menuHelpIndex 
         Caption         =   "He&lp Index"
         Shortcut        =   ^L
      End
      Begin VB.Menu menuAbout 
         Caption         =   "About SdiDesk"
      End
      Begin VB.Menu menuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu menuShowOutlinks 
         Caption         =   "Show Outlinks"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu menuPageVariables 
         Caption         =   "Show Page Variables"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu menuShowPrepared 
         Caption         =   "Show Prepared"
         Shortcut        =   +{F3}
      End
      Begin VB.Menu menuShowHtml 
         Caption         =   "Show HTML"
         Shortcut        =   +{F4}
      End
      Begin VB.Menu menuShowInterMap 
         Caption         =   "Show InterMap"
         Shortcut        =   +{F5}
      End
   End
End
Attribute VB_Name = "WADSMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' SdiDesk 0.2.1

' SdiDesk is Copyright (c) Philip Jones, 2004-2005. All rights reserved.

'  This program is free software; you can redistribute it and/or modify
'  it under the terms of the GNU General Public License as published by
'  the Free Software Foundation; either version 2 of the License, or
'  (at your option) any later version.

'  This program is distributed in the hope that it will be useful,
'  but WITHOUT ANY WARRANTY; without even the implied warranty of
'  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'  GNU General Public License for more details.

'  You should have received a copy of the GNU General Public License
'  along with this program; if not, write to the Free Software
'  Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'  or from the web-site : http://www.gnu.org/licenses/gpl.html

'  Author's Note : this copyright only covers the SdiDesk source-code
'  by myself, Phil Jones, and any other contributers. It isn't intended to
'  apply to any runtime libraries belonging to third parties that you may have
'  also received as part of this package, and on which SdiDesk may depend.
'  These libraries are distributed on the understanding that I, having paid
'  for my copy of Visual Basic, am entitled to distribute them with any
'  application I develop using it. That may not apply to you. And this
'  license explicitly shouldn't be interpretted as granting you any rights
'  over them.

Option Explicit

' Terminology Note :
' VSE stands for "Visual Structure Editor" ... the canvas for editing networks
' WADS stands for "Wiki Annotated Data Store" (now an interface to the
' sub-part of the program

' the configuration factory should be the ONLY object
' that currently knows which concrete classes implement
' most of the major interfaces like ModelLevel, PageStore, Page etc.

Private factory As SdiDeskConfigurationFactory

' MVC "model"
Public MagicNotebook As ModelLevel

' MVC View
Public vm As ViewerManager ' use to show / hide / resize all viewers
Public vse As VseCanvas ' where we draw networks for the Visual Structure Editor
Public td As TableDisplay ' where we edit tables
' + functions of this form are part of view

' MVC controller
Dim docHandler As DocumentHandler ' used to trap events from the WebBrowser control
Public controller As ControlLevel ' all user actions (button clicks, command line) go via this


Public Sub showRawPage(p As Page)
   WADSMainForm.PageNameText = p.pageName
   WADSMainForm.DateCreatedLabel.Caption = "Created : " + CStr(p.createdDate)
   WADSMainForm.DateLastEditedLabel.Caption = "Last Edited : " + CStr(p.lastEdited)
   MagicNotebook.getSingleUserState.isLoading = True
   WADSMainForm.RawText.text = p.raw
   MagicNotebook.getSingleUserState.isLoading = False
   WADSMainForm.HtmlView.Document.body.innerHTML = p.cooked
   
   Call vm.showRaw ' viewer manager controls the visibility

End Sub

Public Sub showCookedPage(p As Page)
   WADSMainForm.PageNameText = p.pageName
   WADSMainForm.DateCreatedLabel.Caption = "Created : " + CStr(p.createdDate)
   WADSMainForm.DateLastEditedLabel.Caption = "Last Edited : " + CStr(p.lastEdited)
   MagicNotebook.getSingleUserState.isLoading = True
   WADSMainForm.RawText.text = p.raw
   MagicNotebook.getSingleUserState.isLoading = False
   
   ' actually this decision should be somewhere else too
   If p.isNetwork Then
     Dim n As Network
     Set n = POLICY_getFactory().wrapPageInNetwork(p)
     Call vm.showVse ' viewer manager controls visibility
     Call Me.vse.draw(n, vse.mode)
     Call MagicNotebook.getControllableModel.setCurrentPage(n)
   Else
     ' show the text in the WebBrowser
     WADSMainForm.HtmlView.Document.body.innerHTML = p.cooked
     Call vm.showHtml
    
     Call waitPageLoad
     If Not docHandler Is Nothing Then
       Call docHandler.recalc
     End If
         
   End If
   
End Sub

Private Sub AllButton_Click()
   Call controller.actionAll(False)
End Sub

Private Sub BackButton_Click()
    Call controller.actionBack
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




Private Sub ExportButton_Click()
    Call controller.actionInstantExport(MagicNotebook.getControllableModel.getCurrentPage.pageName)
End Sub

Private Sub FindButton_Click()
  Call controller.processCommand("#find " + PageNameText.text, False)
End Sub

' Actual form control stuff

Private Sub Form_Load()
  
  ' set html viewer to blank
  WADSMainForm.HtmlView.Navigate2 ("about:blank")
  
  ' setup the factory
  Set factory = POLICY_getFactory()
  
  ' init model level
  Set MagicNotebook = factory.getModelLevel
  
  Call MagicNotebook.setForm(Me)  ' dependency injection
    
  Dim wads As WikiAnnotatedDataStore
  Dim sysConf As SystemConfigurations
  Set wads = MagicNotebook.getWikiAnnotatedDataStore
  Set sysConf = MagicNotebook.getSystemConfigurations
    
  sysConf.configPage = "ConfigPage"
  sysConf.startPage = "StartPage"
  sysConf.helpIndexPage = "HelpIndex"
  sysConf.allPage = "AllPages"
  sysConf.recentChangesPage = "RecentChanges"
  
  
  ' init viwer
  Set docHandler = New DocumentHandler
  Set docHandler.HtmlView = WADSMainForm.HtmlView
    
  Set vm = New ViewerManager
  Call vm.init(WADSMainForm.RawText, WADSMainForm.HtmlView, WADSMainForm.NetworkCanvas, WADSMainForm.TableEditor)
  
  Set vse = New VseCanvas
  Call vse.init(WADSMainForm.NetworkCanvas, WADSMainForm.RawText, MagicNotebook.getControllableModel.getPageCooker, WADSMainForm)
  
  Set td = New TableDisplay
  Call td.init(WADSMainForm.TableEditor, WADSMainForm.EditableCell, WADSMainForm.RawText, WADSMainForm)
  
  ' connect the comboBox on the form with the
  ' navigation history in the model
  Dim nh As NavigationHistory
  Set nh = MagicNotebook.getSingleUserState.history
  Call nh.setComboBox(WADSMainForm.HistoryList)
  Set nh = Nothing
  
  ' init controller
  Set controller = New ControlLevel
  Call controller.init(WADSMainForm, MagicNotebook)
  
  ' let's go
  Call MagicNotebook.getControllableModel.loadNewPage(MagicNotebook.getSystemConfigurations.startPage)
  Call waitPageLoad
  MagicNotebook.getSingleUserState.changesSaved = True
  
  Call MagicNotebook.getExportSubsystem.refreshExportManager(MagicNotebook)
  
  Call showCooked
  MagicNotebook.getSingleUserState.history.append ("StartPage")
  Call setSize(Screen.Width, Screen.Height - 420)
  Left = 0
  Top = 0
End Sub


Public Sub setSize(w As Long, h As Long)
    WADSMainForm.Width = w
    WADSMainForm.Height = h
End Sub


Private Sub Form_Resize()
  On Error Resume Next
  ' Resize controls if Form is not minimized
  If Me.WindowState <> vbMinimized Then
    Dim w As Long, h As Long
    w = WADSMainForm.Width
    h = WADSMainForm.Height
    
    If w >= 8000 And h >= 8000 Then
      ' only if bigger than minimal size
      
      ' set the frames
      
      ' left
      NavFrame.Left = 60
      PageFrame.Left = 60
      ViewFrame.Left = 60
      CategoryFrame.Left = 60
    
      ' widths
      NavFrame.Width = w - 260
      PageFrame.Width = w - 260
      ViewFrame.Width = w - 260
      CategoryFrame.Width = w - 260
    
      ' vertical
      CategoryFrame.Top = h - 1150
      Dim c As Long
      c = WADSMainForm.CategoryFrame.Top - 100
      WADSMainForm.ViewFrame.Height = c - WADSMainForm.ViewFrame.Top
      
      
      ' set the dimensions of the content objects
      
      ' viewers
      Call vm.resize(ViewFrame.Width, ViewFrame.Height, 100)
       
      ' other components
      FindButton.Left = NavFrame.Width - 650
      GoButton.Left = NavFrame.Width - 1300
      PageNameText.Width = NavFrame.Width - 4700
    Else
      WADSMainForm.Width = 8000
      WADSMainForm.Height = 8000
    End If
  End If
End Sub

Public Sub loadPage(n As String)
  Dim pt As String
  pt = MagicNotebook.getControllableModel.loadNewPage(n)
  Select Case pt
  Case "normal":
    showCooked
  Case "newPage":
    showRaw
  Case "network"
    showCooked
  Case "table"
    showCooked
  Case Else
     MsgBox ("came back as something else WADSMainForm:loadPage")
  End Select
  MagicNotebook.getSingleUserState.changesSaved = True
End Sub

Private Sub showRaw()
  Call showRawPage(MagicNotebook.getControllableModel.getCurrentPage)
End Sub

Public Sub showCooked()
  Call showCookedPage(MagicNotebook.getControllableModel.getCurrentPage)
End Sub

Private Sub Form_Terminate()
    Call controller.saveGuard(EditedState)
End Sub

Private Sub ForwardButton_Click()
    Call controller.actionForward
End Sub

Private Sub GoButton_Click()
  Call controller.processCommand(PageNameText.text, False)
End Sub

Private Sub HelpButton_Click()
   Call controller.actionHelp(False)
End Sub

Private Sub HistoryButton_Click()
    Call controller.actionPageHistory(MagicNotebook.getControllableModel.getCurrentPage.pageName, False)
End Sub

Private Sub HistoryList_Click()
    If (HistoryList.text <> "Future" And HistoryList.text <> "History") Then
        Call controller.processCommand(HistoryList.text, False)
    End If
End Sub

Private Sub HtmlView_DocumentComplete(ByVal pDisp As Object, url As Variant)
' this is here because I was getting a blank page after
' loading a page through the interupt. May cause problems though
' as it's a hack

   Call showCooked
End Sub



Private Sub menuAbout_Click()
  Call controller.actionAbout(False)
End Sub

Private Sub menuAll_Click()
  Call controller.actionAll(False)
End Sub

Private Sub menuBack_Click()
  Call controller.actionBack
End Sub

Private Sub menuBackLinks_Click()
  If menuBackLinks.Checked = True Then
    menuBackLinks.Checked = False
    MagicNotebook.getSingleUserState.backlinks = False
  Else
    menuBackLinks.Checked = True
    MagicNotebook.getSingleUserState.backlinks = True
  End If
End Sub

Private Sub menuCrawlers_Click()
  Call controller.actionLoad("CrawlerDefinitions", False)
End Sub

Private Sub menuDelete_Click()
  Call controller.actionDelete(MagicNotebook.getControllableModel.getCurrentPage.pageName)
End Sub

Public Function chooseDirectory() As String
    DirectoryChooser.CancelError = True
    DirectoryChooser.InitDir = App.path
    Dim doit As Boolean
    doit = False
   
    On Error GoTo Cancelled ' most likely cause of error
        Call DirectoryChooser.ShowSave
        doit = True
Cancelled:

    If doit = True Then
        Dim fName As String
        fName = DirectoryChooser.fileName
  Else
     ' cancelled (most likely)
     fName = ""
  End If
  
  chooseDirectory = fName
End Function

Private Sub menuDirectoryChooser_Click()
   ' this stuff is here, rather than as an action,
   ' because it depends on the DirectoryChooser.
   ' Should probably refactor to an action
   
    Dim fName As String
    fName = chooseDirectory()
    If fName <> "" Then
        Call MagicNotebook.getLocalFileSystem.changeDirectory(fName)
    Else
        ' do nothing, probably a cancel
    End If
End Sub

Private Sub menuDynamicExport_Click(index As Integer)
 ' hi
 
End Sub

Private Sub menuEdit_Click()
  Call controller.actionEdit(MagicNotebook.getControllableModel.getCurrentPage.pageName, False)
End Sub

Private Sub menuExit_Click()
   Call controller.actionExit
End Sub

Private Sub menuExports_Click()
   Call controller.actionLoad("ExportDefinitions", False)
End Sub

Private Sub menuExportThisPageHtml_Click()
    Dim dirName As String
    dirName = chooseDirectory()
    Call controller.actionExportOne("web,, " + MagicNotebook.getSingleUserState.currentPage.pageName + ",, " + dirName)
End Sub

Private Sub menuForward_Click()
   Call controller.actionForward
End Sub

Private Sub menuHelpIndex_Click()
   Call controller.actionHelp(False)
End Sub

Private Sub menuHistory_Click()
  Call controller.actionPageHistory(MagicNotebook.getControllableModel.getCurrentPage.pageName, False)
End Sub

Private Sub menuInterMap_Click()
  Call controller.actionLoad("InterMap", False)
End Sub

Private Sub menuLinkTypeDefinitions_Click()
  Call controller.actionLoad("LinkTypeDefinitions", False)
End Sub

Private Sub menuNewNet_Click()
  Call controller.actionNewNetwork(False)
End Sub

Private Sub menuNewPage_Click()
  Call controller.actionNew(False)
End Sub

Private Sub menuPageVariables_Click()
   MsgBox (MagicNotebook.getControllableModel.getCurrentPage.varsToString)
End Sub

Private Sub menuPreview_Click()
   Call controller.actionPreview(False)
End Sub



Private Sub menuRaw_Click()
   Call controller.actionRaw(MagicNotebook.getControllableModel.getCurrentPage.pageName, False)
End Sub

Private Sub menuRecentChanges_Click()
   Call controller.actionRecent(False)
End Sub


Private Sub menuSavePage_Click()
   Call controller.actionSave(MagicNotebook.getControllableModel.getCurrentPage.pageName, False)
End Sub

Private Sub menuShowCrawlers_Click()
    Call controller.actionCrawlers(False)
End Sub

Private Sub menuShowExporters_Click()
    Call controller.actionExporters(False)
End Sub

Private Sub menuShowExports_Click()
    Call controller.actionExports(False)
End Sub

Private Sub menuShowHtml_Click()
  MsgBox (MagicNotebook.getControllableModel.getCurrentPage.cooked)
End Sub

Private Sub menuShowInterMap_Click()
    MsgBox (MagicNotebook.getSystemConfigurations.interMap.toString())
End Sub

Private Sub menuShowOutlinks_Click()
  Dim mg As New WikiMarkupGopher, s As String
  s = mg.getAllTargets(MagicNotebook.getControllableModel.getCurrentPage.raw, MagicNotebook)
  MsgBox (s)
  Set mg = Nothing
End Sub

Private Sub menuShowPrepared_Click()
   MsgBox (MagicNotebook.getControllableModel.getCurrentPage.prepared)
End Sub

Private Sub menuStart_Click()
   Call controller.actionStart(False)
End Sub

Private Sub NetworkCanvas_MouseDown(Button As Integer, shift As Integer, x As Single, y As Single)
   Call vse.MouseDownOnCanvas(MagicNotebook.getControllableModel.getCurrentPage(), Button, shift, x, y)
End Sub

Private Sub NetworkCanvas_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
   If typeName(MagicNotebook.getControllableModel.getCurrentPage()) = "Network" Then
     Dim net As Network
     Set net = MagicNotebook.getControllableModel.getCurrentPage()
     If ((net.hitNodeDetect(x, y) > -1) Or (net.hitArcDetect(x, y).x > -1)) Then
       NetworkCanvas.MousePointer = 4
     Else
       NetworkCanvas.MousePointer = 2
     End If
   End If
End Sub

Private Sub NetworkCanvas_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)
    If typeName(MagicNotebook.getControllableModel.getCurrentPage()) = "Network" Then
        Call vse.endDrag(MagicNotebook.getControllableModel.getCurrentPage(), Button, shift, CLng(x), CLng(y))
    End If
End Sub

Private Sub NewButton_Click()
 Call controller.actionNew(False)
End Sub

Private Sub NewNetworkButton_Click()
   Call controller.actionNewNetwork(False)
End Sub

Private Sub PageNameText_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       Call controller.processCommand(PageNameText.text, False)
   End If
End Sub


Private Sub PresentationButton_Click()
   Call controller.actionPreview(False)
End Sub

Private Sub RawButton_Click()
    Call controller.actionRaw(MagicNotebook.getControllableModel.getCurrentPage.pageName, False)
End Sub



Private Sub RawText_Change()
   If MagicNotebook.getSingleUserState.isLoading = False Then
     MagicNotebook.getSingleUserState.changesSaved = False
   End If
End Sub

Private Sub RecentButton_Click()
   Call controller.actionRecent(False)
End Sub

Private Sub SaveButton_Click()
   Call controller.actionSave(MagicNotebook.getControllableModel.getCurrentPage.pageName, False)
End Sub

Private Sub StartPageButton_Click()
    Call controller.actionStart(False)
End Sub

Public Sub waitPageLoad()
  ' Make sure the Page has loaded
  Do
    DoEvents
  Loop While HtmlView.Busy
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload ArcInfo
    Unload NodeInfo
End Sub

Private Sub TableEditor_DblClick()
    Call Me.td.cellEdit
End Sub

Private Sub TableEditor_MouseDown(Button As Integer, shift As Integer, x As Single, y As Single)
    Call td.startRange
End Sub

Private Sub TableEditor_KeyDown(KeyCode As Integer, shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call Me.td.cellEdit
    End If
    If KeyCode = vbKeyDelete Then
        TableEditor.text = ""
    End If
End Sub


