VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollectionTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Visual Basic seems to have lousy tools for
' looking at collections and arrays, here are
' some useful ones

Public s As String


' Find if a collection has a key, return true if it does
Public Function collectionOfStringsHasKey(c As Collection, k As String) As Boolean
   Dim b As Boolean
   b = False
   On Error GoTo notHere
      Dim a As Variant
      a = c.item(k)
      b = True
notHere:
   collectionOfStringsHasKey = b
End Function

' Find if a collection has a key, return true if it does
Public Function collectionOfObjectsHasKey(c As Collection, k As String) As Boolean
   Dim b As Boolean
   b = False
   On Error GoTo notHere
      Dim a As Object
      Set a = c.item(k)
      b = True
      Set a = Nothing
notHere:
   collectionOfObjectsHasKey = b
End Function


Public Function inspectToString(c As Collection) As String
  Dim s As String
  s = ""
  Dim i As Variant
  If Not c Is Nothing Then
      For Each i In c
        s = s + "*" + CStr(i) + "*" + vbCrLf
      Next i
  End If
  inspectToString = s
End Function



