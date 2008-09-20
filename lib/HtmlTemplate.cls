VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HtmlTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This is an object which holds template information
' to wrap an ExportHTML page

' It also knows how to parse a string

Private myStyleSheet As String
Private myHeader As String
Private myFooter As String

Public varDict As VCollection

Public Property Get styleSheet() As String
    styleSheet = myStyleSheet
End Property

Public Property Get header() As String
    header = myHeader
End Property

Public Property Get footer() As String
    footer = myFooter
End Property

Public Sub init(p As Page)
  ' expecting the three things to be on a page, separated by ----
  ' var definitions should come afterwards
  
  Dim s As String
  
  s = p.raw
  If s = "new page" Or s = "" Then
    myStyleSheet = "<style></style>"
    myHeader = "<body> " + vbCrLf
    myFooter = ""
  Else
    Dim parts() As String
    parts = Split(s, "----")
    myStyleSheet = parts(0)
    myHeader = parts(1)
    myFooter = parts(2)
      
  End If
  
  Set varDict = p.getDataDictionary

End Sub

Private Sub Class_Terminate()
    Set varDict = Nothing
End Sub
