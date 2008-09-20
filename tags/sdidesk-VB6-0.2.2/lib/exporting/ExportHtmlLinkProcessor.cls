VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExportHtmlLinkProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This wraps the links for native HTML.
' Losing the NativeWikiToHtmlDecorator
' Decorator doesn't really seem to work in VB

Implements LinkProcessor

Private st As StringTool
Private myWads As WikiAnnotatedDataStore

Private wmg As WikiMarkupGopher
Private re As RegExp

Private links As OCollection
Private linkCount As Integer

Private Sub Class_Initialize()
    Set st = New StringTool
    Set wmg = New WikiMarkupGopher
    Set re = New RegExp
End Sub

Private Sub Class_Terminate()
    Set st = Nothing
    Set wmg = Nothing
    Set re = Nothing
    Set myWads = Nothing
    Set links = Nothing
End Sub


Private Function doubleBracketContent(l As String) As String
    Dim localRe As New RegExp
    Dim mc As MatchCollection, m As Match
    Dim aLink As New Link
    
    localRe.pattern = "\[\[(\w+>)?(#?(\w|\s|\-)+)\s*(\|(\w|\s)+)?\]\]"
    If localRe.Test(l) Then
        Set mc = localRe.Execute(l)
        Set m = mc.Item(0)
        
        aLink.linkType = m.SubMatches(0)
        If aLink.linkType = "" Then aLink.linkType = "normal"
        If Right(aLink.linkType, 1) = ">" Then
            aLink.linkType = st.trimRight(aLink.linkType)
            aLink.linkType = st.strip(aLink.linkType)
        End If
    
        aLink.target = m.SubMatches(1)
        aLink.target = st.strip(aLink.target)
        If Left(aLink.target, 1) = "#" Then
            aLink.target = Replace(aLink.target, " ", "+", 1)
        End If
        
        aLink.text = m.SubMatches(3)
        If aLink.text = "" Then aLink.text = aLink.target
        If Left(aLink.text, 1) = "|" Then
            aLink.text = st.trimLeft(aLink.text)
            aLink.text = st.strip(aLink.text)
        End If
        
        aLink.target = Replace(aLink.target, " ", "_")
        
        Call links.Add(aLink, CStr(linkCount))
        doubleBracketContent = "LINK" & linkCount
        linkCount = linkCount + 1

    Else
        doubleBracketContent = "Error with :: " & l
    End If
    
    Set localRe = Nothing
    Set mc = Nothing
    Set m = Nothing

End Function

Private Function singleBracketContent(s As String, protocol As String) As String
    Dim re As New RegExp
    re.pattern = "(" & protocol & "(\w|\/|\.)+)\s(.*)"
    If re.Test(s) Then
        Dim l As New Link
        Dim mc As MatchCollection
        Dim m As Match
        Set mc = re.Execute(s)
        Set m = mc.Item(0)
        l.target = m.SubMatches(0)
        l.text = m.SubMatches(2)
        Call links.Add(l, CStr(linkCount))
        singleBracketContent = "LINK" & linkCount
        linkCount = linkCount + 1
    Else
        singleBracketContent = "Error in singleBracketContent with '" & s & "'"
    End If
    Set re = Nothing
End Function

Private Function singleBrackets(s As String, protocol As String)
    Dim mc As MatchCollection, m As Match
    re.pattern = "(.*?)\[(" & protocol & "(\w|\/|\.)+\s((.)*?))\](.*)"
    If re.Test(s) Then
        Set mc = re.Execute(s)
        Set m = mc.Item(0)
        singleBrackets = m.SubMatches(0) & _
        singleBracketContent(m.SubMatches(1), protocol) & _
        singleBrackets(m.SubMatches(5), protocol)
    Else
        singleBrackets = s
    End If
    
End Function

Private Function wrapSingleBracketLinks(l As String) As String
    Dim s As String
    s = singleBrackets(l, "http://")
    s = singleBrackets(s, "https://")
    s = singleBrackets(s, "ftp://")
    s = singleBrackets(s, "ftp://")
    s = singleBrackets(s, "mailto:")
    wrapSingleBracketLinks = s
End Function


Private Function wikiWordsAmongBrackets(l As String) As String
    Dim mc As MatchCollection
    Dim m As Match, i As Integer
    
    Dim build As String
    build = ""
    
    re.pattern = "((.*?)(\[\[.*?\]\])(.*))"
    If re.Test(l) Then
        Set mc = re.Execute(l)
        Set m = mc.Item(0)
        build = build + wikiWords(m.SubMatches(1))
        build = build + doubleBracketContent(m.SubMatches(2))
        build = build + wikiWordsAmongBrackets(m.SubMatches(3))
    Else
        
        build = wikiWords(l)
    End If
    wikiWordsAmongBrackets = build
End Function

Private Function wikiWordToLink(wikiWord As String) As Link
    Dim l As New Link
    l.target = wikiWord
    l.text = wikiWord
    l.linkType = "normal"
    l.nameSpace = ""
    
    Set wikiWordToLink = l
End Function

Private Function wikiWords(l As String) As String
    ' only do this outside square brackets
    Dim build As String, i As Integer
    Dim mc As MatchCollection, m As Match
    Dim aLink As Link

    If InStr(l, "[[") > 0 Then
        build = wikiWordsAmongBrackets(l)
    Else
        re.Global = False
        
        ' (A?[A-Z]([a-z])+(A?[A-Z]([a-z]|\/)+)+)
        re.pattern = "(.*?)" & wmg.wwPattern & "(.*)"
        
        If re.Test(l) Then
            Set mc = re.Execute(l)
            Set m = mc.Item(0)
            Set aLink = wikiWordToLink(m.SubMatches(1))
            Call links.Add(aLink, CStr(linkCount))
            linkCount = linkCount + 1
            build = m.SubMatches(0) & "LINK" & (linkCount - 1) & " " & _
            wikiWords(m.SubMatches(5))
       
        Else
            build = l
        End If
    End If
    wikiWords = build
End Function

Public Function linksToString() As String
    Dim l As Link, s As String, i As Integer
    s = ""
    For Each l In links.toCollection
        s = s & l.toString() & vbCrLf
    Next l
    linksToString = s
End Function

Public Function restoreRealLinks(l As String) As String
    Dim i As Integer, aLink As Link
    Dim s As String
    s = l
    For i = 0 To linkCount - 1
        Set aLink = links.Item(CStr(i))
        s = Replace(s, ("LINK" & i), LinkProcessor_wrapLink(aLink))
    Next i
    restoreRealLinks = s
    Set aLink = Nothing
End Function

Private Property Set LinkProcessor_remoteWads(ByVal RHS As WikiAnnotatedDataStore)
    Set myWads = RHS
End Property

Private Property Get LinkProcessor_remoteWads() As WikiAnnotatedDataStore
    Set LinkProcessor_remoteWads = myWads
End Property

Private Function LinkProcessor_wrapLink(l As Link) As String
    Dim s As String
    If myWads.pageExists(l.target) Or Left(l.target, 1) = "#" Then
        s = "<a href='" & l.target & "'>" & l.text & "</a>"
    Else
        s = l.text
    End If
    LinkProcessor_wrapLink = s
End Function




Private Function LinkProcessor_getAllLinks(l2 As String) As OCollection
    Dim l As String
    l = collectAllLinks(l)
    Set LinkProcessor_getAllLinks = links
End Function

Private Function LinkProcessor_getAllLinksInBigDocument(doc As String) As OCollection
    Dim oc As New OCollection
    Dim lnk As Link, counter As Integer
    Dim lines() As String, l As String
    Dim v As Variant
    lines = Split(doc, vbCrLf)
    counter = 0
    For Each v In lines
        l = CStr(v)
        Call collectAllLinks(l)
        For Each lnk In links.toCollection
            Call oc.Add(lnk, CStr(counter))
            counter = counter + 1
        Next lnk
    Next v
    Set LinkProcessor_getAllLinksInBigDocument = oc
End Function

Private Function LinkProcessor_wrapAllLinks(l As String, lw As LinkWrapper) As String
    Dim s As String
    s = collectAllLinks(l)
    LinkProcessor_wrapAllLinks = restoreRealLinks(s, lw)
End Function



