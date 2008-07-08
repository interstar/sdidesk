Option Strict Off
Option Explicit On
Friend Class StandardLinkProcessor
	Implements _LinkProcessor
	
	' this class is now a generic parser of lines
	' which matches any link-patterns and turns them into link-objects within
	' the "links" collection
	
	' now needs to be given a LinkWrapper to actually turn links into HTML
	
	
	Private st As StringTool
	Private wmg As WikiMarkupGopher
	
	Private links As OCollection
	Private linkCount As Short
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		st = New StringTool
		links = New OCollection
		wmg = New WikiMarkupGopher
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object st may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		st = Nothing
		'UPGRADE_NOTE: Object links may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		links = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Public Function asLinkProcessor() As _LinkProcessor
		asLinkProcessor = Me
	End Function
	
	
	Public Function doubleBracketContent(ByRef l As String) As String
		
		Dim typeArrow As Short
		Dim altBar As Short
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
	
	Public Function singleBracketContent(ByRef s As String) As String
		Dim aLink As New Link
		Dim i As Short
		i = InStr(s, " ")
		aLink.target = Left(s, i - 1)
		aLink.text = Right(s, Len(s) - i)
		aLink.external = True
		
		Call links.add(aLink, CStr(linkCount))
		singleBracketContent = "LINK" & linkCount
		linkCount = linkCount + 1
	End Function
	
	Public Function isUrlChar(ByRef c As String) As Boolean
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
	
	Public Function untilNonUrl(ByRef s As String, ByRef start As Short) As Short
		Dim i As Short
		Dim c As String
		i = start
		c = Mid(s, i, 1)
		While isUrlChar(c)
			i = i + 1
			c = Mid(s, i, 1)
		End While
		untilNonUrl = i - 1
	End Function
	
	
	Public Function looseURL(ByRef l As String, ByRef protocol As String) As String
		' changes a URL in the raw text into appropriate HTML link
		' linkProtocol is something like http://, https://, file://, mailto: etc.
		
		Dim rest, bef, url As String
		Dim aLink As New Link
		Dim i, j As Short
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
			st.trimRight(url)
		End If
		Call aLink.init(url, url, "normal", "", True, False)
		url = "LINK" & linkCount
		Call links.add(aLink, CStr(linkCount))
		linkCount = linkCount + 1
		rest = looseURL(Right(l, Len(l) - j), protocol)
		looseURL = bef & url & rest
	End Function
	
	Public Function singleBrackets(ByRef s As String, ByRef protocol As String) As Object
		Dim bb, be As Short
		
		bb = InStr(s, "[" & protocol)
		be = InStr(s, "]")
		
		If bb > 0 And be > 0 And be > bb Then
			' we found a single bracket link
			'UPGRADE_WARNING: Couldn't resolve default property of object singleBrackets(Right(s, Len(s) - be), protocol). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object singleBrackets. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			singleBrackets = Left(s, bb - 1) & singleBracketContent(Mid(s, bb + 1, (be - bb) - 1)) & singleBrackets(Right(s, Len(s) - be), protocol)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object singleBrackets. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			singleBrackets = s
		End If
	End Function
	
	Private Function wrapSingleBracketLinks(ByRef l As String) As String
		Dim s As String
		If InStr(l, "[") > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object singleBrackets(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			s = singleBrackets(l, "http://")
			'UPGRADE_WARNING: Couldn't resolve default property of object singleBrackets(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			s = singleBrackets(s, "https://")
			'UPGRADE_WARNING: Couldn't resolve default property of object singleBrackets(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			s = singleBrackets(s, "ftp://")
			'UPGRADE_WARNING: Couldn't resolve default property of object singleBrackets(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			s = singleBrackets(s, "mailto:")
		Else
			s = l
		End If
		wrapSingleBracketLinks = s
	End Function
	
	Friend Function wikiWordsAmongBrackets(ByRef l As String) As String
		Dim build As String
		Dim bb, be As Short
		bb = InStr(l, "[[")
		be = InStr(l, "]]")
		If bb > 0 And be > 0 Then
			If bb > 1 Then
				build = wikiWords(Left(l, bb - 1))
			Else
				build = ""
			End If
			'build = build + "##" + Mid(l, bb, (be - bb + 2)) + "##"
			build = build & doubleBracketContent(Mid(l, bb + 2, (be - bb) - 2))
			build = build & wikiWordsAmongBrackets(Right(l, Len(l) - (be + 1)))
			wikiWordsAmongBrackets = build
		Else
			wikiWordsAmongBrackets = wikiWords(l)
		End If
	End Function
	
	Friend Function wikiWordToLink(ByRef wikiWord As String) As Link
		Dim l As New Link
		Dim parts() As String
		If (InStr(wikiWord, ":") > 0) Then
			l.external = True
			l.text = wikiWord
			'MsgBox (wikiWord)
			parts = Split(wikiWord, ":")
			l.nameSpace_Renamed = CStr(parts(0))
			l.target = CStr(parts(1))
			l.linkType = "normal"
			l.interMap = True
			'MsgBox (l.nameSpace)
		Else
			l.external = False
			l.target = wikiWord
			l.text = wikiWord
			l.linkType = "normal"
			l.nameSpace_Renamed = ""
			l.interMap = False
		End If
		wikiWordToLink = l
	End Function
	
	Public Function findNextCapital(ByRef l As String, ByRef start As Short) As Short
		Dim i As Short
		i = start
		While Not wmg.isCapital(Mid(l, i, 1)) And (i < Len(l))
			i = i + 1
		End While
		findNextCapital = i
	End Function
	
	Public Function firstWikiWord(ByRef l As String) As String
		' this function looks to see if there's a wiki word in the string,
		' splits it into three parts :
		' before,
		' WikiWord,
		' rest (which may contain further words, and is analysed recursively)
		
		Dim i, j As Short
		Dim hasWW As Boolean
		Dim startOfWW As Short
		Dim endOfWW As Short
		Dim build As String
		Dim found As Boolean
		Dim ww As String
		
		startOfWW = 0
		endOfWW = 0
		i = Me.findNextCapital(l, 1)
		build = ""
		found = False
		'MsgBox ("aaa : " & l)
		
		Dim aLink As Link
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
				
				aLink = wikiWordToLink(ww)
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
	
	Public Function wikiWords(ByRef l As String) As String
		' only do this outside square brackets
		If InStr(l, "[[") > 0 Then
			wikiWords = wikiWordsAmongBrackets(l)
		Else
			wikiWords = firstWikiWord(l)
		End If
	End Function
	
	Public Function linksToString() As String
		Dim l As Link
		Dim s As String
		Dim i As Short
		s = ""
		For	Each l In links.toCollection
			s = s & l.toString_Renamed() & vbCrLf
		Next l
		linksToString = s
		'UPGRADE_NOTE: Object l may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		l = Nothing
	End Function
	
	
	Public Function restoreRealLinks(ByRef l As String, ByRef lw As _LinkWrapper) As String
		Dim i As Short
		Dim aLink As Link
		Dim s As String
		s = l
		For i = 0 To linkCount - 1
			aLink = links.Item(CStr(i))
			s = Replace(s, "LINK" & i, lw.wrap(aLink))
		Next i
		restoreRealLinks = s
		'UPGRADE_NOTE: Object aLink may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		aLink = Nothing
	End Function
	
	
	Private Function collectAllLinks(ByRef l As String) As String
		Dim s As String
		links = New OCollection
		linkCount = 0
		s = wrapSingleBracketLinks(l)
		s = wikiWords(s)
		s = looseURL(s, "http://")
		s = looseURL(s, "https://")
		s = looseURL(s, "mailto:")
		s = looseURL(s, "file:///")
		
		collectAllLinks = s
	End Function
	
	
	Private Function LinkProcessor_getAllLinks(ByRef l2 As String) As OCollection Implements _LinkProcessor.getAllLinks
		Dim l As String
		l = collectAllLinks(l)
		LinkProcessor_getAllLinks = links
	End Function
	
	Private Function LinkProcessor_getAllLinksInBigDocument(ByRef doc As String) As OCollection Implements _LinkProcessor.getAllLinksInBigDocument
		Dim oc As New OCollection
		Dim lnk As Link
		Dim counter As Short
		Dim lines() As String
		Dim l As String
		Dim v As Object
		lines = Split(doc, vbCrLf)
		counter = 0
		For	Each v In lines
			'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			l = CStr(v)
			Call collectAllLinks(l)
			For	Each lnk In links.toCollection
				Call oc.add(lnk, CStr(counter))
				counter = counter + 1
			Next lnk
		Next v
		LinkProcessor_getAllLinksInBigDocument = oc
	End Function
	
	Private Function LinkProcessor_wrapAllLinks(ByRef l As String, ByRef lw As _LinkWrapper) As String Implements _LinkProcessor.wrapAllLinks
		Dim s As String
		s = collectAllLinks(l)
		LinkProcessor_wrapAllLinks = restoreRealLinks(s, lw)
	End Function
End Class