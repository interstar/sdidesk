VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' a few standard string processing functions

Public Function trimRight(s As String) As String
    trimRight = Left(s, Len(s) - 1)
End Function

Public Function trimLeft(s As String) As String
    trimLeft = Right(s, Len(s) - 1)
End Function

Public Function stripLeft(s2 As String, c As String) As String
  Dim s As String
  s = s2
  While Left(s, 1) = c
    s = trimLeft(s)
  Wend
  stripLeft = s
End Function


Public Function stripRight(s2 As String, c As String) As String
  Dim s As String
  s = s2
  While Right(s, 1) = c
    s = trimRight(s)
  Wend
  stripRight = s
End Function

Public Function strip(s2 As String) As String
  Dim s As String
  s = s2
  While Right(s, 1) = " " Or Right(s, 1) = vbCrLf
        s = trimRight(s)
  Wend
  While Left(s, 1) = " " Or Left(s, 1) = vbCrLf
        s = trimLeft(s)
  Wend
  
  strip = s
  
End Function


Public Function losta(s As String, sep As String) As Integer
  ' length of string to array
  ' returns the length of an array formed by breaking string at sep
  Dim parts() As String
  parts = split(s, sep)
  losta = UBound(parts) + 1
End Function

Public Function star(s As String, sep As String, b As Integer, e As Integer)
' string to array range
' turns a string into an array (splits on sep)
' and returns a new string made up of the desired range
' for example
' - get the first element : star(s,sep,1,1)
' - get the last n elements : star(s,sep,losta(s,sep)-n,losta(s,sep))
  Dim parts() As String, build As String, i As Integer
  build = ""
  parts = split(s, sep)
  For i = b To e
    build = build + parts(i) + sep
  Next i
  star = build
End Function

Public Function leftsa(s As String, sep As String, i As Integer) As String
  leftsa = star(s, sep, 0, i - 1)
End Function

Public Function rightsa(s As String, sep As String, i As Integer) As String
  rightsa = star(s, sep, losta(s, sep) - i, losta(s, sep) - 1)
End Function

Public Function stripHead(s As String, sep As String, n As Integer) As String
  ' remove the first n from the front
  stripHead = rightsa(s, " ", losta(s, " ") - n)
End Function

Public Sub seeAscii(s As String)
' diagnostic functions
  Dim i As Integer, b As String
  For i = 1 To Len(s)
    b = b + (CStr(i) + " : " + Mid(s, i, 1) + " : " + CStr(Asc(Mid(s, i, 1)))) + vbCrLf
  Next i
  MsgBox (b)
End Sub

Public Function removeDoubleChar(s As String, c As String) As String
   ' useful for removing double \\
   removeDoubleChar = Replace(s, c + c, c)
End Function

Public Function mySplit(s As String, sep As String, esc As String) As String()
    Dim parts() As String
    If esc <> "" Then
        Dim i As Integer
        Dim s2 As String
        Dim parts2() As String
        parts2 = split(s, esc)
        If UBound(parts2) > 0 Then
            i = 0
            While i <= UBound(parts2)
                parts2(i) = Replace(parts2(i), sep, "MYSEPARATORBYPHILJONES")
                ' nb : that's an unlikely string, but the function *will* fail if the argument contains it
                ' "this record can not be played on record player B"
                i = i + 2
            Wend
            s2 = Join(parts2, esc)
            parts = split(s2, "MYSEPARATORBYPHILJONES")
        Else
            parts = split(s, sep)
        End If
    Else
        parts = split(s, sep)
    End If
    mySplit = parts
End Function
