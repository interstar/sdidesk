VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExportHtmlLinkWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements LinkWrapper

Dim myWads As WikiAnnotatedDataStore
Dim mySysConf As SystemConfigurations
Dim myMap As InterWikiMap

Public exportedPages As PageSet
Public defaultMakeLinks As Boolean

' exportedPages is a PageSet of other pages which are being exported
' along with this one. If it is defined, and contains the page
' targeted by this link, we'll make it a real link
' if it exists, but doesn't contain the page, we'll leave
' this link as plain-text
' On the other hand, if we haven't been given exportedPages
' we'll appeal to the defaultMakeLinks flag
' Why might this flag be useful? If we have already exported a
' large collection of pages, and now the user wants to make changes
' to one and re-export it, without losing the working links to
' other pages, that should be an option provided by an exporter.
' The flag in ExportHtmlLinkWrapper supports this.

Public Function asLinkWrapper() As LinkWrapper
    Set asLinkWrapper = Me
End Function

Private Sub Class_Initialize()
    Set exportedPages = Nothing
    defaultMakeLinks = False
End Sub

Private Sub Class_Terminate()
    Set exportedPages = Nothing
    Set myWads = Nothing
End Sub

Private Property Set LinkWrapper_remoteInterMap(ByVal RHS As InterWikiMap)
    Set myMap = RHS
End Property

Private Property Get LinkWrapper_remoteInterMap() As InterWikiMap
    Set LinkWrapper_remoteInterMap = myMap
End Property

Private Property Get LinkWrapper_remoteSysConf() As SystemConfigurations
    Set LinkWrapper_remoteSysConf = mySysConf
End Property

Private Property Set LinkWrapper_remoteSysConf(ByVal RHS As SystemConfigurations)
    Set mySysConf = RHS
End Property

Private Property Set LinkWrapper_remoteWads(ByVal RHS As WikiAnnotatedDataStore)
    Set myWads = RHS
End Property

Private Property Get LinkWrapper_remoteWads() As WikiAnnotatedDataStore
    Set LinkWrapper_remoteWads = myWads
End Property

Private Function LinkWrapper_wrap(l As link) As String
    Dim s As String
    If exportedPages Is Nothing Then
        If defaultMakeLinks = True Then
            s = "<a href='" & l.target & "'><div class='" & l.linkType & _
                "'>" & "<font color='#" & mySysConf.getTypeColour(l.linkType) & _
                "'>" & l.text & "</font></div></a>"
        Else
            s = l.text
        End If
    Else
        If exportedPages.hasPage(l.target) Or Left(l.target, 1) = "#" Then
            s = "<a href='" & l.target & "'><div class='" & l.linkType & _
            "'>" & "<font color='#" & mySysConf.getTypeColour(l.linkType) & _
            "'>" & l.text & "</font></div></a>"
        Else
            s = l.text
        End If
    End If
    LinkWrapper_wrap = s
End Function

