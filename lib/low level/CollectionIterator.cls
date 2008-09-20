VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollectionIterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Iterator

' a collection implementation of iterator

Dim c As Collection
Dim walker As Integer

Public Sub init(aCollection As Collection)
    Set c = aCollection
    Call Iterator_zero
End Sub

Private Sub Class_Terminate()
    Set c = Nothing
End Sub

Private Function Iterator_eof() As Boolean
    Iterator_eof = walker > c.Count
End Function

Private Function Iterator_length() As Integer
    Iterator_length = c.Count
End Function

Private Function Iterator_nextObject() As Object
    Set Iterator_nextObject = c.Item(walker)
    walker = walker + 1
End Function

Private Function Iterator_nextVariant() As Variant
    Iterator_nextVariant = c.Item(walker)
    walker = walker + 1
End Function

Private Sub Iterator_zero()
    walker = 1
End Sub

