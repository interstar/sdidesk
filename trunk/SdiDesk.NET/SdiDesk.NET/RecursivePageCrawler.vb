Option Strict Off
Option Explicit On
Friend Class RecursivePageCrawler
	Implements _PageCrawler
	
	Private myWads As _WikiAnnotatedDataStore ' where to get everything else
	
	Private myPages As PageSet ' where we keep the pages while crawling
	
	Private myName As String ' useful to know the name of the crawler
	Private myMaxDepth As Short ' how deep to crawl
	Private myExcludedPages As VCollection ' pages not to crawl to
	Private myLinkTypeBehaviours As VCollection ' linkTypes and what to do with them
	Private myDefaultLinkTypeBehaviour As String ' are we including or excluding
	
	Public Sub init(ByRef aName As String, ByRef maxDepth As Short, ByRef excludedPages As String, ByRef linkTypeBehaviours As String)
		myName = aName
		myMaxDepth = maxDepth
		Call parseExcludedPagesFromString(excludedPages)
		Call parseLinkTypeBehavioursFromString(linkTypeBehaviours)
	End Sub
	
	Public Function asPageCrawler() As _PageCrawler
		asPageCrawler = Me
	End Function
	
	Public Sub parseExcludedPagesFromString(ByRef s As String)
		' format of s is
		' PageName|PageName2|PageName3
		myExcludedPages = New VCollection
		Dim parts() As String
		Dim v As Object
		Dim v2 As String
		If s <> "" Then ' if blank argument, do nothing
			If InStr(s, "|") > 0 Then ' multiple excluded pages
				parts = Split(s, "|")
				For	Each v In parts
					'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v2 = CStr(v)
					If Not myExcludedPages.hasKey(v2) Then
						Call myExcludedPages.Add(v2, v2)
					End If
				Next v
			Else
				' one excluded page
				Call myExcludedPages.Add(s, s)
			End If
		End If
	End Sub
	
	Public Sub parseLinkTypeBehavioursFromString(ByRef s As String)
		' format of s is
		' +explanation|definition|normal|counterArg
		' which means, the default is exclude but include the following list
		' alternatively
		' -explanation|definition
		' means that the default is include but exclude the following list
		
		myLinkTypeBehaviours = New VCollection
		
		Dim parts() As String
		Dim v As Object
		Dim v2 As String
		If s = "" Or s = " " Then
			' no args, defaults to +
			myDefaultLinkTypeBehaviour = "+"
			
		Else
			If Left(s, 1) = "+" Then
				myDefaultLinkTypeBehaviour = "-"
			Else
				myDefaultLinkTypeBehaviour = "+"
			End If
			' the above might be confusing?
			' if we put a + at the front, these are things we're explicitly *including*
			' against a default of excluding
			' if we put a - at the front, these are things we're explicity *excluding*
			' against a default of including
			
			' strip off the first char
			s = Right(s, Len(s) - 1)
			
			parts = Split(s, "|")
			For	Each v In parts
				'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				v2 = CStr(v)
				If Not myLinkTypeBehaviours.hasKey(v2) Then
					Call myLinkTypeBehaviours.Add(v2, v2)
				End If
			Next v
		End If
	End Sub
	
	
	Public Sub recursiveCrawl(ByRef startPage As String, ByRef depth As Short)
		' this function starts with the name of a page
		' gathers it's out-links, then
		' follows out-links to any other pages not currently in the main set
		' so DOESN'T go circular
		' and has the possibility of a maximum depth restriction
		' At the end, we should have gathered all relevant pages into pages
		
		' let's go ...
		
		Dim p As Object
		Dim recurse As Boolean
		Dim ps As New PageSet
		Dim r As String
		If depth < myMaxDepth Or myMaxDepth < 0 Then
			' otherwise we're not going anywhere.
			' note if you set maxDepth to say, -1, then there's no maximum
			
			ps.init()
			
			' now fill it from the out-links from startPage
			' this doesn't handle extra links from including. Should we?
			r = myWads.getRawPageData(startPage)
			ps = PageCrawler_fillPageSetFromString(r)
			
			' now iterate through it
			
			
			For	Each p In ps.pages.toCollection
				
				' if this page has NOT yet been gathered into our "pages" set,
				' we will recurse into it
				'UPGRADE_WARNING: Couldn't resolve default property of object p.pageName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Not myPages.hasPage(p.pageName) Then
					' ie. we DON'T yet have this page
					' let's have it
					'UPGRADE_WARNING: Couldn't resolve default property of object p. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Call myPages.addPage(p)
					
					' and let's do the recursion
					'UPGRADE_WARNING: Couldn't resolve default property of object p.pageName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Call Me.recursiveCrawl(p.pageName, depth + 1)
				End If
				
			Next p
			
		Else 'depth >= maxDepth, just come out
		End If
		
		'UPGRADE_NOTE: Object p may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		p = Nothing
		'UPGRADE_NOTE: Object ps may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		ps = Nothing
		
	End Sub
	
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		myPages = New PageSet
		Call myPages.init()
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object myPages may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		myPages = Nothing
		'UPGRADE_NOTE: Object myWads may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		myWads = Nothing
		'UPGRADE_NOTE: Object myPages may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		myPages = Nothing
		
		'UPGRADE_NOTE: Object myExcludedPages may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		myExcludedPages = Nothing
		'UPGRADE_NOTE: Object myLinkTypeBehaviours may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		myLinkTypeBehaviours = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	
	Private Sub PageCrawler_clear() Implements _PageCrawler.clear
		' resets the PageSet
		myPages = New PageSet
		Call myPages.init()
	End Sub
	
	Private Sub PageCrawler_crawl(ByRef startPage As String) Implements _PageCrawler.crawl
		myPages.clearOut()
		Call Me.recursiveCrawl(startPage, 0)
	End Sub
	
	Private Function linkExcluded(ByRef lnk As Link) As Boolean
		If lnk.isCommand() Then
			linkExcluded = True
			Exit Function
		End If
		
		If (myDefaultLinkTypeBehaviour = "+") Then
			' exclusions are explicit,
			If myLinkTypeBehaviours.hasKey((lnk.linkType)) Then
				' it's in the excluded list
				linkExcluded = True
				Exit Function
			End If
		Else
			' inclusions are explicit
			If Not myLinkTypeBehaviours.hasKey((lnk.linkType)) Then
				' it's not in the included list
				linkExcluded = True
				Exit Function
			End If
		End If
		
		If Not myWads.pageExists((lnk.target)) Then
			linkExcluded = True
			Exit Function
		End If
		
		' but now let's see if this page itself if explicitly excluded
		linkExcluded = myExcludedPages.hasKey((lnk.target))
		
	End Function
	
	Private Function filterOutlinks(ByRef initialOuts As OCollection) As PageSet
		Dim lnk As Link
		Dim ps As New PageSet
		Dim p As _Page
		Call ps.init()
		Call ps.clearOut()
		
		For	Each lnk In initialOuts.toCollection
			If Not linkExcluded(lnk) Then
				p = myWads.store.loadRaw((lnk.target))
				Call ps.addPage(p)
			End If
		Next lnk
		
		filterOutlinks = ps
	End Function
	
	Private Function outLinksFromNetwork() As OCollection
		
	End Function
	
	Private Function PageCrawler_fillPageSetFromPage(ByRef p As _Page) As PageSet Implements _PageCrawler.fillPageSetFromPage
		If Not p.isNetwork Then
			Call p.cook(POLICY_getFactory().getPagePreparer, POLICY_getFactory().getNativePageCooker, False)
			PageCrawler_fillPageSetFromPage = PageCrawler_fillPageSetFromString((p.prepared))
		Else
			
		End If
	End Function
	
	Private Function PageCrawler_fillPageSetFromString(ByRef s As String) As PageSet Implements _PageCrawler.fillPageSetFromString
		' Creates a new PageSet and fills it
		' from the raw links defined in s
		
		Dim outlinks As OCollection
		
		Dim lp As _LinkProcessor
		lp = POLICY_getFactory().getStandardLinkProcessor
		
		outlinks = lp.getAllLinksInBigDocument(s)
		
		PageCrawler_fillPageSetFromString = filterOutlinks(outlinks)
	End Function
	
	Private Function PageCrawler_getPages() As PageSet Implements _PageCrawler.getPages
		PageCrawler_getPages = myPages
	End Function
	
	
	Private Property PageCrawler_wads() As _WikiAnnotatedDataStore Implements _PageCrawler.wads
		Get
			PageCrawler_wads = myWads
		End Get
		Set(ByVal Value As _WikiAnnotatedDataStore)
			myWads = Value
		End Set
	End Property
	
	
	Private Property PageCrawler_name() As String Implements _PageCrawler.name
		Get
			PageCrawler_name = myName
		End Get
		Set(ByVal Value As String)
			myName = Value
		End Set
	End Property
	
	
	Private Property PageCrawler_pages() As PageSet Implements _PageCrawler.pages
		Get
			PageCrawler_pages = myPages
		End Get
		Set(ByVal Value As PageSet)
			PageCrawler_pages = Value
		End Set
	End Property
	
	
	Private Function PageCrawler_toString() As String Implements _PageCrawler.toString_Renamed
		Dim s As String
		s = "BOX<" & vbCrLf & "<p>'''" & myName & "''' is an example of a Recursive''''''Page''''''Crawler</p>" & "<p>It follows links to a maximum depth of " & myMaxDepth & " from the start-page</p>" & "<p>It excludes these pages : " & vbCrLf & myExcludedPages.toString_Renamed() & "</p>" & vbCrLf
		If myDefaultLinkTypeBehaviour = "-" Then
			s = s & "<p>It also ignores all links except those of these types : " & vbCrLf & myLinkTypeBehaviours.toString_Renamed() & vbCrLf & "</P>" & vbCrLf
		Else
			s = s & "<p>It also ignores all links of these types : " & myLinkTypeBehaviours.toString_Renamed() & vbCrLf & "</p>" & vbCrLf
		End If
		s = s & ">BOX" & vbCrLf & "<p>To edit or add a new crawler, please go to the CrawlerDefinitions page.</p>"
		PageCrawler_toString = s
	End Function
End Class