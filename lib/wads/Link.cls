VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Link"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Object that represents a link and all it's potential attributes

Public target As String ' where the link is going
Public text As String ' what you see
Public linkType As String ' the type of this link
Public nameSpace As String ' is the link at a remote site, if so namespace
Public external As Boolean ' is this a link to the outside world?
Public interMap As Boolean ' is this an intermap link?

' used for WikiMap

Public Sub init(txt As String, targ As String, lTyp As String, ns As String, ext As Boolean, imap As Boolean)
    target = targ
    text = txt
    linkType = lTyp
    nameSpace = ns
    external = ext
    interMap = imap
End Sub

Public Function deepCopy() As Link
    Dim l As New Link
    l.external = external
    l.interMap = interMap
    l.linkType = linkType
    l.nameSpace = nameSpace
    l.target = target
    l.text = text
    Set deepCopy = l
End Function

Public Function toString() As String
    toString = "(" & target & ", " & text & ", " & linkType & ", " & nameSpace & ", " & external & ")"
End Function

Public Function isCommand() As Boolean
    If Left(target, 1) = "#" Then
        isCommand = True
    Else
        isCommand = False
    End If
End Function

Public Function matches(l2 As Link) As Boolean
    matches = True
    If l2.target <> target Then
        matches = False
        Exit Function
    End If
    
    If l2.text <> text Then
        matches = False
        Exit Function
    End If
    
    If l2.external <> external Then
        matches = False
        Exit Function
    End If
    
    If l2.interMap <> interMap Then
        matches = False
        Exit Function
    End If
    
    If l2.linkType <> linkType Then
        matches = False
        Exit Function
    End If
    
    If l2.nameSpace <> nameSpace Then
        matches = False
        Exit Function
    End If
    
End Function

Private Sub Class_Initialize()
    target = ""
    text = ""
    linkType = ""
    nameSpace = ""
    external = False
End Sub
