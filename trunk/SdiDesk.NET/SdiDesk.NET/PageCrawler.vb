Option Strict Off
Option Explicit On
Interface _PageCrawler
	 Property wads As _WikiAnnotatedDataStore
	 Property pages As PageSet
	 Property name As String
	Sub clear()
	Function getPages() As PageSet
	Function fillPageSetFromString(ByRef s As String) As PageSet
	Function fillPageSetFromPage(ByRef p As _Page) As PageSet
	Function toString_Renamed() As String
	Sub crawl(ByRef startPage As String)
End Interface
Friend Class PageCrawler
	Implements _PageCrawler
	
	' The purpose of the PageCrawler is to go around gathering pages
	' into a PageSet according to a strategy
	' now an interface
	
	
	Dim wads_MemberVariable As WikiAnnotatedDataStore
    Public Property wads() As _WikiAnnotatedDataStore Implements _PageCrawler.wads
        Get
            wads = wads_MemberVariable
        End Get
        Set(ByVal Value As _WikiAnnotatedDataStore)
            wads_MemberVariable = Value
        End Set
    End Property ' where to get pages etc.
	
	Dim pages_MemberVariable As PageSet
	Public Property pages() As PageSet Implements _PageCrawler.pages
		Get
			pages = pages_MemberVariable
		End Get
		Set(ByVal Value As PageSet)
			pages_MemberVariable = Value
		End Set
	End Property ' where we keep the pages while crawling
	
	Dim name_MemberVariable As String
	Public Property name() As String Implements _PageCrawler.name
		Get
			name = name_MemberVariable
		End Get
		Set(ByVal Value As String)
			name_MemberVariable = Value
		End Set
	End Property ' useful to know the name of the crawler
	
	Public Sub clear() Implements _PageCrawler.clear
		' clear the number of pages crawled
	End Sub
	
	Public Function getPages() As PageSet Implements _PageCrawler.getPages
		' return the current pages crawled in a pageset
	End Function
	
	Public Function fillPageSetFromString(ByRef s As String) As PageSet Implements _PageCrawler.fillPageSetFromString
		' start with a string representation of the body of a page,
		' and get all pages linked from it
	End Function
	
	Public Function fillPageSetFromPage(ByRef p As _Page) As PageSet Implements _PageCrawler.fillPageSetFromPage
		' this starts with a page and fills the pages from it.
		' handles networks differently etc.
	End Function
	
	'UPGRADE_NOTE: toString was upgraded to toString_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function toString_Renamed() As String Implements _PageCrawler.toString_Renamed
		' return a string representation of the pages crawled
	End Function
	
	Public Sub crawl(ByRef startPage As String) Implements _PageCrawler.crawl
		' do the crawl
	End Sub
End Class