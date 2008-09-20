VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NativeLinkWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements LinkWrapper

Private myWads As WikiAnnotatedDataStore
Private mySysConf As SystemConfigurations
Private myMap As interWikiMap

Private Sub Class_Terminate()
    Set myWads = Nothing
    Set mySysConf = Nothing
    Set myMap = Nothing
End Sub

Public Function asLinkWrapper() As LinkWrapper
    Set asLinkWrapper = Me
End Function

Private Property Set LinkWrapper_remoteInterMap(ByVal RHS As interWikiMap)
    Set myMap = RHS
End Property

Private Property Get LinkWrapper_remoteInterMap() As interWikiMap
    Set LinkWrapper_remoteInterMap = myMap
End Property

Private Property Set LinkWrapper_remoteSysConf(ByVal RHS As SystemConfigurations)
    Set mySysConf = RHS
End Property

Private Property Get LinkWrapper_remoteSysConf() As SystemConfigurations
    Set LinkWrapper_remoteSysConf = mySysConf
End Property

Private Property Set LinkWrapper_remoteWads(ByVal RHS As WikiAnnotatedDataStore)
    Set myWads = RHS
End Property

Private Property Get LinkWrapper_remoteWads() As WikiAnnotatedDataStore
    Set LinkWrapper_remoteWads = myWads
End Property

Private Function LinkWrapper_wrap(l As link) As String
    Dim s As String, iMapUrl As String
        
    If l.interMap = True Then
        'MsgBox (l.nameSpace)
        'MsgBox (l.toString)
        If myMap.getUrl(l.nameSpace) <> "ERROR" Then
            'MsgBox (l.target)
            'MsgBox (l.nameSpace)
            l.target = myMap.getUrl(l.nameSpace) + l.target
            LinkWrapper_wrap = "<a href='" & l.target & "'>" & l.text & "</a>"
            Exit Function
        Else
            'MsgBox (l.nameSpace)
            LinkWrapper_wrap = "<font color=''>Warning. '" & l.nameSpace & "' not defined</font>                    "
            Exit Function
        End If
    End If
    
    If l.external = True Then
        s = "<a href='" & l.target & "' id='external' target='new'>" & l.text & "</a>"
    Else
        If myWads.pageExists(l.target) Or Left(l.target, 1) = "#" Then
            s = "<a href='about:blank' class='" & _
            l.linkType & "' id='" & l.target & "'><font color='" & _
            mySysConf.getTypeColour(l.linkType) & "'>" & _
            l.text & "</font></a>"
        Else
            s = "<a href='about:blank' class='newPage' id='" & l.target & _
            "'><font color='#ff6666'>" & l.text & "</font></a>"
        End If
    End If
    LinkWrapper_wrap = s
End Function

