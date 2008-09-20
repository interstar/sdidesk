VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InterWikiMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This is a mapping between other wiki names and their URLs

Private map As VCollection
Private nameMap As VCollection

Public Sub add(url As String, key As String)
    Call map.add(url, key)
    Call nameMap.add(key, key)
End Sub

Public Function getUrl(key As String) As String
    If map.hasKey(key) Then
        getUrl = CStr(map.Item(key))
    Else
        getUrl = "ERROR"
    End If
End Function

Public Sub parseFromString(s As String)
    Dim lines() As String
    Dim parts() As String
    Dim v As Variant
    lines = Split(s, vbCrLf)
    For Each v In lines
        parts = Split(CStr(v), " ")
        If UBound(parts) > 0 Then
            Call Me.add(parts(1), parts(0))
        End If
    Next v
End Sub

Public Function toString() As String
    Dim s As String
    Dim v As Variant
    For Each v In nameMap.toCollection
        s = s & v & ", " & Me.getUrl(CStr(v)) & vbCrLf
    Next v
    toString = s
End Function

Private Sub Class_Initialize()
    Set map = New VCollection
    Set nameMap = New VCollection
End Sub

Private Sub Class_Terminate()
    Set map = Nothing
    Set nameMap = Nothing
End Sub



