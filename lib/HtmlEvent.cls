VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HtmlEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'needed when we trap the "clicked on a link" event in the HTML viewer

Private handler As Object ' what handles the event
Private action As String ' the action to call
Private id As String

Public Sub Event_Details(ByRef aHandler As Object, ByRef anAction As String, anId As String)
  Set handler = aHandler
  action = anAction
  id = anId
End Sub

Public Sub HTML_Event()
Attribute HTML_Event.VB_UserMemId = 0
  ' HTML EVENT TRIGGERED
  ' This MUST be the Default Procedure for the Class Module
  ' The routine will be processed when the event occurs
  'MsgBox ("hi " + id)
  Dim a(0) As Variant
  a(0) = id
  CallByName handler, action, VbMethod, a
End Sub

Private Sub Class_Terminate()
  On Error Resume Next
  Set handler = Nothing
End Sub


