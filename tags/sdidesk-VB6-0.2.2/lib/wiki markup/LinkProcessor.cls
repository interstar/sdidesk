VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LinkProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This interface analyses a line to extract links
' and wraps them all with the LinkWrapper

' these operate on the level of the individual line
Public Function wrapAllLinks(l As String, wrapper As LinkWrapper) As String
End Function

Public Function getAllLinks(l As String) As OCollection
End Function


' but sometimes the processor should be able to tackle a whole
' document. We'll leave it up to the processer how it does it
Public Function getAllLinksInBigDocument(doc As String) As OCollection
End Function
