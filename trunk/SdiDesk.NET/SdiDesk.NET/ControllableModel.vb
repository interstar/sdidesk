Option Strict Off
Option Explicit On
Interface _ControllableModel
    Function getSingleUserState() As _SingleUserState
    Function getSystemConfigurations() As _SystemConfigurations
    Function getWikiAnnotatedDataStore() As _WikiAnnotatedDataStore
    Function getLocalFileSystem() As _LocalFileSystem
    Function getExportSubsystem() As _ExportSubsystem
    Function getCrawlerSubsystem() As _CrawlerSubsystem
    Function getPageCooker() As _PageCooker
    Function getPagePreparer() As PagePreparer
    Function loadNewPage(ByRef pageName As String) As String
    Function loadRawPage(ByRef pageName As String) As String
    Function getCurrentPage() As _Page
    Sub setCurrentPage(ByRef p As _Page)
    Sub savePage()
    Function newPage() As String
    Function newNetworkPage() As String
    Sub deletePage(ByRef pageName As String)
    Function wordCount(ByRef pageName As String) As Short
    Function makeSearchResultsPage(ByRef searchTerm As String) As _Page
    Function makeHistoryPage(ByRef pageName As String) As _Page
End Interface
Friend Class ControllableModel
	Implements _ControllableModel
	' ControllableModel is the interface that the ControlLevel talks to
	
	
	Public Function getSingleUserState() As _SingleUserState Implements _ControllableModel.getSingleUserState
	End Function
	
	Public Function getSystemConfigurations() As _SystemConfigurations Implements _ControllableModel.getSystemConfigurations
	End Function
	
	Public Function getWikiAnnotatedDataStore() As _WikiAnnotatedDataStore Implements _ControllableModel.getWikiAnnotatedDataStore
	End Function
	
	Public Function getLocalFileSystem() As _LocalFileSystem Implements _ControllableModel.getLocalFileSystem
	End Function
	
	Public Function getExportSubsystem() As _ExportSubsystem Implements _ControllableModel.getExportSubsystem
	End Function
	
	Public Function getCrawlerSubsystem() As _CrawlerSubsystem Implements _ControllableModel.getCrawlerSubsystem
	End Function
	
	Public Function getPageCooker() As _PageCooker Implements _ControllableModel.getPageCooker
	End Function
	
	Public Function getPagePreparer() As PagePreparer Implements _ControllableModel.getPagePreparer
	End Function
	
	Public Function loadNewPage(ByRef pageName As String) As String Implements _ControllableModel.loadNewPage
	End Function
	
	Public Function loadRawPage(ByRef pageName As String) As String Implements _ControllableModel.loadRawPage
	End Function
	
	Public Function getCurrentPage() As _Page Implements _ControllableModel.getCurrentPage
	End Function
	
	Public Sub setCurrentPage(ByRef p As _Page) Implements _ControllableModel.setCurrentPage
	End Sub
	
	Public Sub savePage() Implements _ControllableModel.savePage
	End Sub
	
	Public Function newPage() As String Implements _ControllableModel.newPage
	End Function
	
	Public Function newNetworkPage() As String Implements _ControllableModel.newNetworkPage
	End Function
	
	Public Sub deletePage(ByRef pageName As String) Implements _ControllableModel.deletePage
	End Sub
	
	Public Function wordCount(ByRef pageName As String) As Short Implements _ControllableModel.wordCount
	End Function
	
	Public Function makeSearchResultsPage(ByRef searchTerm As String) As _Page Implements _ControllableModel.makeSearchResultsPage
	End Function
	
	Public Function makeHistoryPage(ByRef pageName As String) As _Page Implements _ControllableModel.makeHistoryPage
	End Function
End Class