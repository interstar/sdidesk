Option Strict Off
Option Explicit On
Friend Class WikiMarkupGopher
	
	' this class does various useful mark-up functions
	' but contains no state
	
	Public Function qq(ByRef s As String) As String
		' puts string in quotes
		qq = Chr(34) & s & Chr(34)
	End Function
	
	Public Function rightFromPosition(ByRef s As String, ByRef p As Short) As Object
		' returns the right from the position p.
		' unlike right(s,10)
		' 10 is counted from the *left*
		rightFromPosition = Right(s, Len(s) - p)
	End Function
	
	Public Function isImage(ByRef url As String) As Boolean
		Dim r As Boolean
		r = False
		If Right(url, 4) = ".gif" Then r = True
		If Right(url, 4) = ".jpg" Then r = True
		If Right(url, 4) = ".bmp" Then r = True
		If Right(url, 4) = ".png" Then r = True
		If Right(url, 5) = ".jpeg" Then r = True
		isImage = r
	End Function
	
	
	Public Function rightFromChar(ByRef s As String, ByRef c As String) As String
		' returns the right part of a string from the last instance of c
		Dim p As Short
		p = InStrRev(s, c)
		'UPGRADE_WARNING: Couldn't resolve default property of object rightFromPosition(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rightFromChar = rightFromPosition(s, p)
	End Function
	
	Public Function testTwoTags(ByRef l As String, ByRef wikiPattern As String) As Boolean
		' this function returns true if there are at least two instances of wikiPattern
		' in string l
		Dim j As Short
		j = InStr(l, wikiPattern)
		If j > 0 Then
			j = InStr(j + 1, l, wikiPattern)
			If j > 0 Then
				testTwoTags = True
				Exit Function
			End If
		End If
		testTwoTags = False
	End Function
	
	
	Public Function wrapTags(ByRef l As String, ByRef wikiPattern As String, ByRef openTag As String, ByRef closeTag As String) As String
		' while there are two instances of the tag wikiPattern,
		' replace the first instance with openTag and the second with closeTag
		
		Dim l2 As String
		l2 = l
		Do While testTwoTags(l2, wikiPattern)
			l2 = Replace(l2, wikiPattern, openTag, 1, 1)
			l2 = Replace(l2, wikiPattern, closeTag, 1, 1)
		Loop 
		wrapTags = l2
	End Function
	
	Public Function isAlpha(ByRef c As String) As Boolean
		Dim i As Short
		If c = "" Then
			isAlpha = False
			Exit Function
		End If
		i = Asc(c)
		If (i >= 48 And i <= 57) Or (i >= 65 And i <= 90) Or (i >= 97 And i <= 122) Or (i >= 192 And i <= 221) Or (i >= 224 And i <= 253) Then
			isAlpha = True
		Else
			isAlpha = False
		End If
	End Function
	
	Public Function isSlash(ByRef c As String) As Boolean
		If c = "/" Then
			isSlash = True
		Else
			isSlash = False
		End If
	End Function
	
	
	Public Function isAlphaOrSlash(ByRef c As String) As Boolean
		' now we have sub-pages need to allow slashes, but don't want
		' to screw-up isAlpha
		isAlphaOrSlash = (isAlpha(c) Or isSlash(c))
	End Function
	
	Public Function isAlphaOrSlashOrColon(ByRef c As String) As Boolean
		' now we have sub-pages need to allow slashes, but don't want
		' to screw-up isAlpha
		isAlphaOrSlashOrColon = (isAlpha(c) Or isSlash(c) Or c = ":")
	End Function
	
	
	Public Function nextNonAlpha(ByRef s As String) As Short
		' returns the position of next non-alpha
		Dim i As Short
		For i = 1 To Len(s)
			If Not isAlphaOrSlashOrColon(Mid(s, i, 1)) Then
				Exit For
			End If
		Next i
		nextNonAlpha = i
	End Function
	
	
	Public Function untilNextNonAlpha(ByRef s As String) As String
		' take the left of string s until the next non-alpha character
		Dim e As Short ' end of alpha
		e = Me.nextNonAlpha(s) - 1
		untilNextNonAlpha = Left(s, e)
	End Function
	
	Public Function isCapital(ByRef c As String) As Boolean
		If c = "" Then
			isCapital = False
			Exit Function
		Else
			If (Asc(c) > 64 And Asc(c) < 91) Or (Asc(c) >= 192 And Asc(c) <= 221) Then
				isCapital = True
			Else
				isCapital = False
			End If
		End If
	End Function
	
	Public Function hasCapital(ByRef l As String) As Boolean
		Dim i As Short
		Dim c As String
		
		hasCapital = False
		For i = 1 To Len(l)
			c = Mid(l, i, 1)
			If isCapital(c) Then
				hasCapital = True
				Exit Function
			End If
		Next i
	End Function
	
	Public Function measureWikiWordAtFront(ByRef s As String) As Short
		Dim i As Short
		Dim c As String
		Dim lastCapital As Boolean
		Dim secondCapital, hasSlash As Boolean
		Dim hasColon, secondColon As Boolean
		Dim wikiWord As Boolean
		wikiWord = True
		
		If Not isCapital(Left(s, 1)) Then
			measureWikiWordAtFront = -1
			Exit Function
		End If
		
		wikiWord = True
		lastCapital = True
		secondCapital = False
		hasSlash = False
		i = 2
		
		While i < Len(s) And isAlphaOrSlashOrColon(Mid(s, i, 1))
			c = Mid(s, i, 1)
			
			If isSlash(c) And isCapital(Mid(s, i - 1, 1)) Then
				measureWikiWordAtFront = -1
				Exit Function
			End If
			
			If Not Me.isAlphaOrSlashOrColon(c) Then
				measureWikiWordAtFront = -1
				Exit Function
			End If
			If lastCapital = True Then
				
				' previous was capital
				If isCapital(c) Then
					If Not Mid(s, i - 1, 1) = "A" Then
						measureWikiWordAtFront = -1
						Exit Function
					End If
					lastCapital = True
					If i > 2 Then
						secondCapital = True
					End If
				Else
					If (isSlash(c)) Then
						If hasSlash Then ' prevent two slashes
							measureWikiWordAtFront = -1
							Exit Function
						Else
							hasSlash = True
						End If
					End If
					If (c = ":") Then
						If hasColon Then ' prevent two colons
							measureWikiWordAtFront = -1
							Exit Function
						Else
							hasColon = True
						End If
					End If
					lastCapital = False
				End If
			Else
				' previous was not capital
				If isCapital(c) Then
					lastCapital = True
					secondCapital = True
				Else
					If (isSlash(c)) Then
						If hasSlash Then ' prevent two slashes
							measureWikiWordAtFront = -1
							Exit Function
						Else
							hasSlash = True
							lastCapital = False
						End If
					End If
					If (c = ":") Then
						If hasColon Then ' prevent two colons
							measureWikiWordAtFront = -1
							Exit Function
						Else
							hasColon = True
						End If
					End If
					lastCapital = False
				End If
			End If
			i = i + 1
		End While
		
		If secondCapital = False Then
			measureWikiWordAtFront = -1
			Exit Function
		End If
		
		If s = "" Then
			measureWikiWordAtFront = -1
			Exit Function
		End If
		
		If Not isAlpha(Mid(s, i, 1)) Then
			i = i - 1
		End If
		measureWikiWordAtFront = i
	End Function
	
	
	Public Function isWikiWord(ByRef s As String) As Boolean
		Dim i As Short
		If s = "" Then
			isWikiWord = False
		Else
			i = measureWikiWordAtFront(s)
			'MsgBox (s & " : " & i & " : " & Len(s))
			If i = Len(s) And isAlpha(Right(s, 1)) Then
				isWikiWord = True
			Else
				isWikiWord = False
			End If
		End If
	End Function
	
	
	Public Function extractBracketCommand(ByRef s2 As String) As String
		Dim s As String
		s = s2
		
		Dim nc As New NavCommand
		If Left(s, 1) = "#" Then
			' this is a command
			Call nc.init(s)
			s = nc.getCommand
		Else
			s = ""
		End If
		
		extractBracketCommand = s
		' clean up
		'UPGRADE_NOTE: Object nc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		nc = Nothing
	End Function
	
	Public Function extractBracketContents(ByRef s2 As String) As String
		
		Dim s As String
		s = s2
		
		Dim nc As New NavCommand
		If Left(s, 1) = "#" Then
			' this is a command
			Call nc.init(s)
			s = nc.getPageName
		End If
		
		s = Replace(s, " ", "_")
		
		extractBracketContents = s
		
		' clean up
		'UPGRADE_NOTE: Object nc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		nc = Nothing
	End Function
	
	Public Function getAllTargets(ByRef l2 As String, ByRef wads As _WikiAnnotatedDataStore) As String
		Dim build As String
		Dim blp As _LinkProcessor
		blp = POLICY_getFactory().getStandardLinkProcessor
		
		Dim c As New OCollection
		Dim noted As New VCollection
		
		Dim lnk As Link
		
		c = blp.getAllLinksInBigDocument(l2)
		
		build = ""
		For	Each lnk In c.toCollection
			If wads.pageExists((lnk.target)) Then
				If Not noted.hasKey(CStr(lnk.target)) Then
					Call noted.add(CStr(lnk.target), CStr(lnk.target))
					build = build & lnk.target & ",, "
				End If
			End If
		Next lnk
		'UPGRADE_NOTE: Object c may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		c = Nothing
		'UPGRADE_NOTE: Object noted may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		noted = Nothing
		getAllTargets = build
	End Function
	
	
	Public Function untilNext(ByRef s2 As String, ByRef pattern As String) As Object
		MsgBox("wmg.until next deprecated")
	End Function
	
	Public Function afterNext(ByRef s2 As String, ByRef pattern As String) As Object
		Dim pos As Short
		Dim s As String
		pos = InStr(s2, pattern)
		
		If pos > 0 Then
			s = Right(s2, (Len(s2) - pos) - Len(pattern) + 1)
		Else
			s = ""
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object afterNext. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		afterNext = s
	End Function
End Class