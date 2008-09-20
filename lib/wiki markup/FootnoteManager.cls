VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FootnoteManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' this object manages the extraction of footnotes embedded in pages and
' places them at the bottom of the page

Public footnotes As Collection
Private counter As Integer

Public Sub init()
    Set footnotes = New Collection
    counter = 1
End Sub

Public Sub addFootnote(index As Integer, note As String)
    Call footnotes.Add(note, CStr(index))
End Sub

Public Function extractFootnotes(line As String, native As Boolean) As String
    Dim s As String, note As String
    Dim i As Integer, v As Variant
    If InStr(line, "{{") Then
        Dim parts() As String
        parts = Split(line, "{{")
        s = parts(0)
        For i = 1 To UBound(parts)
            If InStr(parts(i), "}}") Then
                note = Left(parts(i), InStr(parts(i), "}}") - 1)
                Call addFootnote(counter, note)
                If Not native Then
                    'internal links not working in the native html
                    MsgBox (1)
                    s = s + "(<b><a href='#" & counter & "'>" & counter & "</a></b>)"
                Else
                    s = s + "(<b>" & counter & "</b>)"
                End If
                s = s + Mid(parts(i), InStr(parts(i), "}}") + 2)
                counter = counter + 1
            Else
                s = s + parts(i)
            End If
        Next i
          
        extractFootnotes = s
    Else
        extractFootnotes = line
    End If
End Function


Public Function getFootnoteCollection() As Collection
    Set getFootnoteCollection = footnotes
End Function

Public Function getFootnotesAsHtmlString() As String
    Dim s As String
    Dim i As Integer
    s = ""
    For i = 1 To counter - 1
        ' nasty, shouldn't be here.
        s = s & "* (<a name='" & i & "'>" & i & "</a>) "
        s = s + footnotes.item(i)
        s = s + vbCrLf
    Next i
    getFootnotesAsHtmlString = s
End Function

