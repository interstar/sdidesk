Option Strict Off
Option Explicit On
Friend Class AllPagesPageCrawler
	Implements _PageCrawler
	
	' A PageCrawler which returns all pages
	
	
	Private myWads As _WikiAnnotatedDataStore
	
	Private myName As String
	Private myPages As PageSet
	Private myStore As _PageStore
	
	Public Sub init(ByRef aName As String)
		myName = aName
	End Sub
	
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object myWads may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		myWads = Nothing
		'UPGRADE_NOTE: Object myPages may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		myPages = Nothing
		'UPGRADE_NOTE: Object myStore may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		myStore = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	
	Private Sub PageCrawler_clear() Implements _PageCrawler.clear
		myPages = New PageSet
	End Sub
	
	Private Sub PageCrawler_crawl(ByRef startPage As String) Implements _PageCrawler.crawl
		myPages = myStore.AllPages
	End Sub
	
	Private Function PageCrawler_fillPageSetFromPage(ByRef p As _Page) As PageSet Implements _PageCrawler.fillPageSetFromPage
		Call PageCrawler_crawl("")
        'PageCrawler_fillPageSetFromPage = PageCrawler.getPages
        MessageBox.Show("reached at error posiont 2")
	End Function
	
	Private Function PageCrawler_fillPageSetFromString(ByRef s As String) As PageSet Implements _PageCrawler.fillPageSetFromString
		Call PageCrawler_crawl("")
        'PageCrawler_fillPageSetFromString = PageCrawler.getPages
        MessageBox.Show("reached at error posiont 2")
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
			myPages = Value
		End Set
	End Property
	
	
	Private Function PageCrawler_toString() As String Implements _PageCrawler.toString_Renamed
		PageCrawler_toString = "'''" & myName & "'''" & " is an example of an ''All''''''Pages'' crawler. It picks up all the pages in the wiki. (The list of pages you see on the AllPages page.)"
	End Function
End Class