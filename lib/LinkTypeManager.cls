VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LinkTypeManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' stores the collection of linkTypes and colours

Public c As VCollection ' will store in the format of key = typeName, val = colourDef


Public Sub init()
  Set c = New VCollection
End Sub


Public Function setupLinkTypes(raw As String) As String
  ' we're expecting raw to be in table format
  ' be
  ' typeName,, colourDef
  ' typeName,, colourDef
  ' etc.
    
  ' reset the collection
  Call Me.init
  
  ' Put in a default colour which may be over-ridden by data from raw
  Call c.Add("#006633", "normal")
  
  ' now parse the data from raw into a table, and then use to
  ' fill the collection
  Dim t As New table
  Call t.parseFromDoubleCommaString(raw)
  
  Dim st As New StringTool
  Dim i As Integer, typeName As String, colDef As String
  For i = 0 To t.noRows - 1
      typeName = st.strip(CStr(t.at(i, 0)))
      colDef = st.strip(CStr(t.at(i, 1)))
      If c.hasKey(typeName) Then
         c.Remove (typeName)
      End If
      
      Call c.Add(colDef, typeName)
  Next i
  Set t = Nothing
  Set st = Nothing
End Function

Public Function getColour(tn As String) As String
  If c.hasKey(tn) Then
    getColour = c.Item(tn)
  Else
    getColour = c.Item("normal")
  End If
End Function

Public Function toString() As String
  toString = c.toString()
End Function

Public Function toStyleDefs() As String
  MsgBox ("LinkTypeManager:toStyleDefs unwritten")
'  For Each s2 In d
'    Dim s As String
'    s = CStr(s2)
'    If s <> "" Then
'      Dim parts() As String
'      parts = Split(s, ",, ")
'      build = build + "." + parts(0) + " {color:" + parts(1) + "}" + vbCrLf
'
      
'    End If
'  Next s2
  
'  build = build + "</style>" + vbCrLf + "</head>" + vbCrLf
'  styleSheet = build ' may need to change this later
'  setupLinkTypes = build
  toStyleDefs = ""
End Function
