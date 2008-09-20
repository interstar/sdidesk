VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DocumentHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' this class handles the HTML document in the document viewer
' it collects the HTML links and accepts the

Public HtmlView As WebBrowser
Private links As Object ' these will be the link elements of the document
Public noLinks As Integer
Private linkObjects() As Object

Public Sub recalc()
  Dim x As Object, i As Integer
  Set links = HtmlView.Document.links
  
  i = 0
  ' get the length of the collection (!)
  For Each x In links: i = i + 1: Next x
  noLinks = i

  ReDim linkObjects(noLinks + 1)
  For i = 0 To noLinks - 1
    Dim anEvent As New HtmlEvent
    anEvent.Event_Details Me, "HTML_Click", links(i).id
    Set linkObjects(i) = links(i)
    Set linkObjects(i).onclick = anEvent
    Set anEvent = Nothing
  Next i
  
End Sub

Public Sub HTML_Click(a() As Variant)
  Call WADSMainForm.waitPageLoad
  If CStr(a(0)) = "external" Then
  Else
    Call WADSMainForm.controller.processCommand(CStr(a(0)), False)
  End If
End Sub

