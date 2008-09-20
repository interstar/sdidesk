VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ViewerManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' this object manages the visibility of the various viewers available,
' and arranges which will be visible, invisible.

Public TableEditor As MSFlexGrid
Public NetworkCanvas As PictureBox
Public HtmlView As WebBrowser
Public RawText As RichTextBox

Public Enum ViewerManagerMode
    vmmNetwork
    vmmTable
    vmmHtml
    vmmRaw
End Enum

Public mode As ViewerManagerMode

Public Sub init(raw As RichTextBox, html As WebBrowser, net As PictureBox, table As MSFlexGrid)
    Set RawText = raw
    Set HtmlView = html
    Set NetworkCanvas = net
    Set TableEditor = table
End Sub

Public Sub hideAll()
    RawText.Visible = False
    HtmlView.Visible = False
    NetworkCanvas.Visible = False
    TableEditor.Visible = False
End Sub

Public Sub showRaw()
    Call hideAll
    RawText.Visible = True
    RawText.SetFocus
    mode = vmmRaw
End Sub

Public Sub showHtml()
    Call hideAll
    HtmlView.Visible = True
    'HtmlView.SetFocus (doesn't seem to work)
    mode = vmmHtml
End Sub

Public Sub showVse()
    Call hideAll
    NetworkCanvas.Visible = True
    NetworkCanvas.SetFocus
    mode = vmmNetwork
End Sub

Public Sub showTable()
    Call hideAll
    TableEditor.Visible = True
    TableEditor.SetFocus
    mode = vmmTable
End Sub

Public Sub resize(vfWidth As Integer, vfHeight As Integer, dfb As Integer)

      HtmlView.Left = 60
      HtmlView.Width = (vfWidth - 140)
      HtmlView.Height = vfHeight - dfb
      
      RawText.Left = 60
      RawText.Width = vfWidth - 140
      RawText.Height = vfHeight - dfb
      
      NetworkCanvas.Left = 60
      NetworkCanvas.Width = vfWidth - 140
      NetworkCanvas.Height = vfHeight - dfb
  
      TableEditor.Left = 60
      TableEditor.Width = vfWidth - 140
      TableEditor.Height = vfHeight - dfb
  
End Sub
