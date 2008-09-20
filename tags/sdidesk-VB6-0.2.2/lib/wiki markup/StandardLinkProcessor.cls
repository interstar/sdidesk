VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StandardLinkProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' this class is now a generic parser of lines
' which matches any link-patterns and turns them into link-objects within
' the "links" collection

' now needs to be given a LinkWrapper to actually turn links into HTML

Implements LinkProcessor

Private st As StringTool
Private wmg As WikiMarkupGopher

Private links As OCollection
Private linkCount As Integer

Private Sub Class_Initialize()
    Set st = New StringTool
    Set links = New OCollection
    Set wmg = New WikiMarkupGopher
End Sub

Private Sub Class_Terminate()
    Set st = Nothing
    Set links = Nothing
End Sub

Public Function asLinkProcessor() As LinkProcessor
    Set asLinkProcessor = Me
End Function


Public Function doubleBracketContent(l As String) As String
    
    Dim typeArrow As Integer
    Dim altBar As Integer
    Dim rest As String
    Dim aLink As New Link
    
    rest = l
    
    altBar = InStr(rest, "|")
    If altBar > 0 Then
        aLink.text = st.strip(Right(rest, Len(l) - altBar))
        rest = Left(rest, altBar - 1)
    End If
    
    typeArrow = InStr(rest, ">")
    If typeArrow > 0 Then
        aLink.linkType = st.strip(Left(l, typeArrow - 1))
        rest = st.strip(Right(rest, Len(rest) - (typeArrow)))
    End If
    
    aLink.target = rest
    If aLink.text = "" Then
        aLink.text = aLink.target
    End If
    
    If Left(aLink.target, 1) = "#" Then
        aLink.target = Replace(aLink.target, " ", "+", 1)
    End If
        
    aLink.target = Replace(aLink.target, " ", "_")
    
    Call links.add(aLink, CStr(linkCount))
    doubleBracketContent = "LINK" & linkCount
    linkCount = linkCount + 1
End Function

Public Function singleBracketContent(s As String) As String
    Dim aLink As New Link
    Dim i As Integer
    i = InStr(s, " ")
    aLink.target = Left(s, i - 1)
    aLink.text = Right(s, Len(s) - i)
    aLink.external = True
    
    Call links.add(aLink, CStr(linkCount))
    singleBracketContent = "LINK" & linkCount
    linkCount = linkCount + 1
End Function

Public Function isUrlChar(c As String) As Boolean
    isUrlChar = True
    If c = " " Then isUrlChar = False
    If c = "" Then isUrlChar = False
    If c = "*" Then isUrlChar = False
    If c = "(" Or c = ")" Then isUrlChar = False
    If c = "[" Or c = "]" Then isUrlChar = False
    If c = "{" Or c = "}" Then isUrlChar = False
    If c = "<" Or c = ">" Then isUrlChar = False
    If c = "#" Then isUrlChar = False
End Function

Public Function untilNonUrl(s As String, start As Integer) As Integer
    Dim i As Integer, c As String
    i = start
    c = Mid(s, i, 1)
    While isUrlChar(c)
        i = i + 1
        c = Mid(s, i, 1)
    Wend
    untilNonUrl = i - 1
End Function


Public Function looseURL(l As String, protocol As String) As String
  ' changes a URL in the raw text into appropriate HTML link
  ' linkProtocol is something like http://, https://, file://, mailto: etc.
  
  Dim bef As String, rest As String, url As String
  Dim aLink As New Link
  Dim i As Integer, j As Integer
  i = InStr(l, protocol)
  If i < 1 Then
    looseURL = l
    Exit Function
  End If
  

  If i > 1 Then
      If (Mid(l, i - 1, 1) = "[") Then
          looseURL = l
          Exit Function
      End If
  End If
  ' i > 0 and not "[" at i-1
  
  bef = Left(l, i - 1)
  j = untilNonUrl(l, i)
  url = Mid(l, i, j - (i - 1))
  If Right(url, 1) = vbCrLf Then
    st.trimRight (url)
  End If
  Call aLink.init(url, url, "normal", "", True, False)
  url = "LINK" & linkCount
  Call links.add(aLink, CStr(linkCount))
  linkCount = linkCount + 1
  rest = looseURL(Right(l, Len(l) - j), protocol)
  looseURL = bef & url & rest
End Function

Public Function singleBrackets(s As String, protocol As String)
    Dim bb As Integer, be As Integer
    
    bb = InStr(s, "[" & protocol)
    be = InStr(s, "]")
    
    If bb > 0 And be > 0 And be > bb Then
        ' we found a single bracket link
        singleBrackets = Left(s, bb - 1) & _
        singleBracketContent(Mid(s, bb + 1, (be - bb) - 1)) & _
        singleBrackets(Right(s, Len(s) - be), protocol)
    Else
        singleBrackets = s
    End If
End Function

Private Function wrapSingleBracketLinks(l As String) As String
    Dim s As String
    If InStr(l, "[") > 0 Then
        s = singleBrackets(l, "http://")
        s = singleBrackets(s, "https://")
        s = singleBrackets(s, "ftp://")
        s = singleBrackets(s, "mailto:")
    Else
        s = l
    End If
    wrapSingleBracketLinks = s
End Function

Friend Function wikiWordsAmongBrackets(l As String) As String
    Dim build As String
    Dim bb As Integer, be As Integer
    bb = InStr(l, "[[")
    be = InStr(l, "]]")
    If bb > 0 And be > 0 Then
        If bb > 1 Then
            build = wikiWords(Left(l, bb - 1))
        Else
            build = ""
        End If
        'build = build + "##" + Mid(l, bb, (be - bb + 2)) + "##"
        build = build + doubleBracketContent(Mid(l, bb + 2, (be - bb) - 2))
        build = build + wikiWordsAmongBrackets(Right(l, Len(l) - (be + 1)))
        wikiWordsAmongBrackets = build
    Else
        wikiWordsAmongBrackets = wikiWords(l)
    End If
End Function

Friend Function wikiWordToLink(wikiWord As String) As Link
    Dim l As New Link
    If (InStr(wikiWord, ":") > 0) Then
        l.external = True
        l.text = wikiWord
        Dim parts() As String
        'MsgBox (wikiWord)
        parts = Split(wikiWord, ":")
        l.nameSpace = CStr(parts(0))
        l.target = CStr(parts(1))
        l.linkType = "normal"
        l.interMap = True
        'MsgBox (l.nameSpace)
    Else
        l.external = False
        l.target = wikiWord
        l.text = wikiWord
        l.linkType = "normal"
        l.nameSpace = ""
        l.interMap = False
    End If
    Set wikiWordToLink = l
End Function

Public Function findNextCapital(l As String, start As Integer) As Integer
    Dim i As Integer
    i = start
    While Not wmg.isCapital(Mid(l, i, 1)) And (i < Len(l))
        i = i + 1
    Wend
    findNextCapital = i
End Function

Public Function firstWikiWord(l As String) As String
' this function looks to see if there's a wiki word in the string,
' splits it into three parts :
' before,
' WikiWord,
' rest (which may contain further words, and is analysed recursively)

Dim i As Integer, j As Integer
Dim hasWW As Boolean
Dim startOfWW As Integer
Dim endOfWW As Integer
Dim build As String
Dim found As Boolean
Dim ww As String

startOfWW = 0
endOfWW = 0
i = Me.findNextCapital(l, 1)
build = ""
found = False
'MsgBox ("aaa : " & l)

Do While i < Len(l)
    j = wmg.measureWikiWordAtFront(Right(l, Len(l) - (i - 2)))
    If j < 0 Then
        'MsgBox ("b : *" & Right(l, Len(l) - (i - 2)) & "*")
        i = Me.findNextCapital(l, i)
        i = i + 1
    Else
        If i > 2 Then
            startOfWW = i - 1
            endOfWW = i + j - 1
        Else
            startOfWW = 1
            endOfWW = i + j
        End If
        ww = Mid(l, startOfWW, j)
        'MsgBox ("found WikiWord : *" & ww & "*")
        
        If startOfWW = 1 Then
            build = ""
        Else
            build = Left(l, startOfWW - 1)
        End If
        Dim aLink As Link
        
        Set aLink = wikiWordToLink(ww)
        Call links.add(aLink, CStr(linkCount))
        linkCount = linkCount + 1
        build = build & "LINK" & (linkCount - 1)
        build = build & firstWikiWord(Right(l, Len(l) - (endOfWW - 1)))
    
        found = True
        Exit Do
    End If
Loop

If found = True Then
    firstWikiWord = build
Else
    firstWikiWord = l
End If
End Function

Public Function wikiWords(l As String) As String
    ' only do this outside square brackets
    If InStr(l, "[[") > 0 Then
        wikiWords = wikiWordsAmongBrackets(l)
    Else
        wikiWords = firstWikiWord(l)
    End If
End Function

Public Function linksToString() As String
    Dim l As Link, s As String, i As Integer
    s = ""
    For Each l In links.toCollection
        s = s & l.toString() & vbCrLf
    Next l
    linksToString = s
    Set l = Nothing
End Function


Public Function restoreRealLinks(l As String, lw As LinkWrapper) As String
    Dim i As Integer, aLink As Link
    Dim s As String
    s = l
    For i = 0 To linkCount - 1
        Set aLink = links.Item(CStr(i))
        s = Replace(s, ("LINK" & i), lw.wrap(aLink))
    Next i
    restoreRealLinks = s
    Set aLink = Nothing
End Function


Private Function collectAllLinks(l As String) As String
    Dim s As String
    Set links = New OCollection
    linkCount = 0
    s = wrapSingleBracketLinks(l)
    s = wikiWords(s)
    s = looseURL(s, "http://")
    s = looseURL(s, "https://")
    s = looseURL(s, "mailto:")
    s = looseURL(s, "file:///")
    
    collectAllLinks = s
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
            Call oc.add(lnk, CStr(counter))
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



