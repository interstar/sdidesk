VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Iterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' An iterator interface for arrays and collections

Public Function eof() As Boolean
' is end of the thing being walked?
End Function

' stupid type-checking

Public Function nextVariant() As Variant
' return the next item if it's a variant
End Function

Public Function nextObject() As Object
' return the next item if it's an object
End Function

Public Sub zero()
' reset the walker to zero
End Sub

Public Function length() As Integer
' the length of the sequence to be iterated
End Function
