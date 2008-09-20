VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' VB6s collections suck, bigtime.

Private col As Collection
Private keys As Collection

Public Sub Add(v As Object, k As String)
    Call col.Add(v, k)
    Call keys.Add(k, k)
End Sub

Public Function Count() As Integer
    Count = col.Count
End Function

Public Function Item(k As Variant) As Object
    Set Item = col.Item(k)
End Function

Public Sub Remove(k As Variant)
    Call col.Remove(k)
    Call keys.Remove(k)
End Sub

' Find if a collection has a key, return true if it does
Public Function hasKey(k As String) As Boolean
   Dim a As Object, b As Boolean
   b = False
   On Error GoTo notHere
      Set a = col.Item(k)
      b = True
      Set a = Nothing
notHere:
   hasKey = b
End Function

Public Function toCollection() As Collection
    ' return the collection for "for each"ing
    Set toCollection = col
End Function



Private Sub Class_Initialize()
    Set col = New Collection
    Set keys = New Collection
End Sub

Private Sub Class_Terminate()
    Set col = Nothing
    Set keys = Nothing
End Sub


