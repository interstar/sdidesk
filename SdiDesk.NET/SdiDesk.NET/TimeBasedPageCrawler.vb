Option Strict Off
Option Explicit On
Friend Class TimeBasedPageCrawler
	Implements _PageCrawler
	
	' page crawler for finding recent pages / monthly pages
	' at the moment, only gets recent
	' as we need to work out how to give it all the parameters it
	' needs ... the use of a single crawler definition table is breaking
	' down!!1
	
	Private myWads As _WikiAnnotatedDataStore '
	Private myStore As _PageStore
	Private myPages As PageSet ' where we keep the pages while crawling
	
	Private myName As String ' useful to know the name of the crawler
	Private myExcludedPages As VCollection ' pages not to crawl to
	Private myDefaultLinkTypeBehaviour As String ' undefined linkTypes
	Private myLinkTypeBehaviours As VCollection ' linkTypes and what to do with them
	
	
	Public Sub init(ByRef aName As String, ByRef ep As String, ByRef ltb As String)
		myName = aName
		Call parseExcludedPagesFromString(ep)
		Call parseLinkTypeBehavioursFromString(ltb)
	End Sub
	
	Public Sub parseExcludedPagesFromString(ByRef s As String)
		' format of s is
		' PageName1|PageName2|PageName3
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
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		myPages = New PageSet
		Call myPages.init()
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	Private Sub PageCrawler_clear() Implements _PageCrawler.clear
		' resets the PageSet
		myPages = New PageSet
		Call myPages.init()
	End Sub
	
	Private Sub PageCrawler_crawl(ByRef startPage As String) Implements _PageCrawler.crawl
		Dim se As New ScriptEngine_Renamed
		Dim recentString As String
		myPages.clearOut()
		'UPGRADE_WARNING: Couldn't resolve default property of object myWads. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		recentString = se.recentChanges(myWads)
		myPages = PageCrawler_fillPageSetFromString(recentString)
	End Sub
	
	Private Function PageCrawler_fillPageSetFromPage(ByRef p As _Page) As PageSet Implements _PageCrawler.fillPageSetFromPage
		' does nothing for this crawler
	End Function
	
	Private Function PageCrawler_fillPageSetFromString(ByRef s As String) As PageSet Implements _PageCrawler.fillPageSetFromString
		' does nothing for this crawler
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
	
	
	Private Property PageCrawler_store() As _PageStore
		Get
			PageCrawler_store = myStore
		End Get
		Set(ByVal Value As _PageStore)
			myStore = Value
		End Set
	End Property
	
	Private Function PageCrawler_toString() As String Implements _PageCrawler.toString_Renamed
		Dim s As String
		s = "<p>'''" & myName & "''' is an example of a Time''''''Based''''''Page''''''Crawler </p>" & "<p>It excludes these pages : " & myExcludedPages.toString_Renamed() & "</p>"
		PageCrawler_toString = s
	End Function
End Class