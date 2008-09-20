VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayIterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Iterator for an array

Option Explicit
Implements Iterator

Private walker As Integer
Private a

Public Sub init(anArray)
    a = anArray
End Sub

Private Function Iterator_eof() As Boolean
    Iterator_eof = (walker > UBound(a))
End Function

Private Function Iterator_nextObject() As Object
    Set Iterator_nextObject = a(walker)
    walker = walker + 1
End Function

Private Function Iterator_nextVariant() As Variant
    Iterator_nextVariant = a(walker)
    walker = walker + 1
End Function

Private Sub Iterator_zero()
    walker = LBound(a)
End Sub

Private Function Iterator_length() As Integer
    Iterator_length = (UBound(a) - LBound(a)) + 1
End Function
