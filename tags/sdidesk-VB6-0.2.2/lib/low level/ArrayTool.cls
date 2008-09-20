VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' couple of low level routines for arrays

' this copies a subrange of one array of strings to another
Public Function copyStringArray(a1() As String, a2() As String, s1 As Integer, e1 As Integer, s2 As Integer)
   Dim a2c As Integer, i As Integer
   a2c = s2
   For i = s1 To e1
     a2(a2c) = a1(i)
     a2c = a2c + 1
   Next i
End Function

' this copies a subrange of one array of varianst to another
Public Function copyVariantArray(a1(), a2(), s1 As Integer, e1 As Integer, s2 As Integer)
   Dim a2c As Integer, i As Integer
   a2c = s2
   For i = s1 To e1
     a2(a2c) = a1(i)
     a2c = a2c + 1
   Next i
End Function


' return a string which lists the contents of an array
' will break if the contents break cstr

Public Function inspectArray(a1())
   Dim l As Integer, s As String
   l = UBound(a1) + 1
   s = ""
   Dim i As Integer
   For i = 0 To l - 1
     s = s + CStr(a1(i)) + ",, "
   Next i
   inspectArray = s
   
End Function

