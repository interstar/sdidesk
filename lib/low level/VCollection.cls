VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' VB6s collections suck, bigtime.

Private col As Collection
Private myKeys As Collection

Public Sub add(value As Variant, key As String)
    Call col.add(value, key)
    Call myKeys.add(key, key)
End Sub

Public Function Count() As Integer
    Count = col.Count
End Function

Public Function Item(k As String) As Variant
    Item = col.Item(k)
End Function

Public Sub Remove(k As Variant)
    Call col.Remove(k)
    Call myKeys.Remove(k)
End Sub

' Find if a collection has a key, return true if it does
Public Function hasKey(k As String) As Boolean
   Dim a As Variant, b As Boolean
   b = False
   On Error GoTo notHere
      a = col.Item(k)
      b = True
notHere:
   hasKey = b
End Function


Public Function toCollection() As Collection
    ' return the collection for "for each"ing
    Set toCollection = col
End Function

Public Function keyCollection() As Collection
    Set keyCollection = myKeys
End Function

Public Function toString() As String
  Dim s As String, i As Variant
  s = ""
  For Each i In myKeys
   
        s = s & "* (" & CStr(i) & ", " & CStr(col.Item(CStr(i))) & ")" & vbCrLf
  Next i
  toString = s
End Function

Public Function keysToString() As String
  Dim s As String, i As Variant
  s = ""
  For Each i In myKeys
        s = s + "* " + CStr(i) + vbCrLf
  Next i
  keysToString = s
End Function


Private Sub Class_Initialize()
    Set col = New Collection
    Set myKeys = New Collection
End Sub

Private Sub Class_Terminate()
    Set col = Nothing
    Set myKeys = Nothing
End Sub

