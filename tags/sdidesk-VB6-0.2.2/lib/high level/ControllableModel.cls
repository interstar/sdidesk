VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControllableModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ControllableModel is the interface that the ControlLevel talks to


Public Function getSingleUserState() As SingleUserState
End Function

Public Function getSystemConfigurations() As SystemConfigurations
End Function

Public Function getWikiAnnotatedDataStore() As WikiAnnotatedDataStore
End Function

Public Function getLocalFileSystem() As LocalFileSystem
End Function

Public Function getExportSubsystem() As ExportSubsystem
End Function

Public Function getCrawlerSubsystem() As CrawlerSubsystem
End Function

Public Function getPageCooker() As PageCooker
End Function

Public Function getPagePreparer() As PagePreparer
End Function

Public Function loadNewPage(pageName As String) As String
End Function

Public Function loadRawPage(pageName As String) As String
End Function

Public Function getCurrentPage() As Page
End Function

Public Sub setCurrentPage(p As Page)
End Sub

Public Sub savePage()
End Sub

Public Function newPage() As String
End Function

Public Function newNetworkPage() As String
End Function

Public Sub deletePage(pageName As String)
End Sub

Public Function wordCount(pageName As String) As Integer
End Function

Public Function makeSearchResultsPage(searchTerm As String) As Page
End Function

Public Function makeHistoryPage(pageName As String) As Page
End Function

