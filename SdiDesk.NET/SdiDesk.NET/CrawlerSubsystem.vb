Option Strict Off
Option Explicit On
Interface _CrawlerSubsystem
	 Property crawlerManager As CrawlerDefinitionTable
    Function makeCrawlersPage() As _Page
    Function makeCrawlResultsPage(ByRef crawlerName As String, ByRef startPage As String) As _Page
End Interface
Friend Class CrawlerSubsystem
	Implements _CrawlerSubsystem
	
	' this interface for the subsystem that handles crawlers
	
	Dim crawlerManager_MemberVariable As CrawlerDefinitionTable
	Public Property crawlerManager() As CrawlerDefinitionTable Implements _CrawlerSubsystem.crawlerManager
		Get
			crawlerManager = crawlerManager_MemberVariable
		End Get
		Set(ByVal Value As CrawlerDefinitionTable)
			crawlerManager_MemberVariable = Value
		End Set
	End Property
	
	
	Public Function makeCrawlersPage() As _Page Implements _CrawlerSubsystem.makeCrawlersPage
	End Function
	
	Public Function makeCrawlResultsPage(ByRef crawlerName As String, ByRef startPage As String) As _Page Implements _CrawlerSubsystem.makeCrawlResultsPage
	End Function
End Class