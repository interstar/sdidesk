VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LineArgAnaliser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Used by Exporter programs, this class analyses the line args

Public argDict As VCollection

Public Function asVCollection() As VCollection
    Set asVCollection = argDict
End Function

Public Sub analise(s As String)
    Dim st As New StringTool
    Dim parts() As String
    Dim k As String, v As String
    parts = st.mySplit(s, " ", "'")
    Dim i As Integer, ok As Boolean
    
    For i = 0 To UBound(parts) Step 2
            ok = True
            If Left(parts(i), 1) = "-" Then
                k = st.trimLeft(parts(i))
            Else
                ok = False
            End If
            If Left(parts(i + 1), 1) = "'" And Right(parts(i + 1), 1) = "'" Then
                v = st.trimLeft(parts(i + 1))
                v = st.trimRight(v)
            Else
                ok = False
            End If
            If ok = True Then
                Call argDict.add(v, k)
            End If
    Next i
    Set st = Nothing
End Sub

Public Function toString() As String
    Dim v As Variant
    Dim s As String
    s = ""
    For Each v In argDict.keyCollection
        s = s & "(" & v & ", " & argDict.Item(CStr(v)) & ") " & vbCrLf
    Next v
    toString = s
End Function

Private Sub Class_Initialize()
    Set argDict = New VCollection
End Sub

Private Sub Class_Terminate()
    Set argDict = Nothing
End Sub
