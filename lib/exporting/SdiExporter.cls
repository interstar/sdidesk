VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SdiExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Exporter

' This is another Exporter, designed to allow you to export a set of pages
' in raw page format, for use in another sdi

Private myWads As WikiAnnotatedDataStore ' link to WADS

Private selfName As String ' my name
Private selfVersion As String ' my version
Private selfPath As String ' my path
Private selfStartPage As String ' my copy of the start page

Private selfCrawlers As Collection ' the collection of crawlers that define this export
Private selfPages As PageSet ' the collection of pages to be exportet



Public Function saveStringInFile(s As String, fileName As String)
   Open fileName For Output As 1
   Print #1, s
   Close #1
End Function


Private Sub Exporter_addCrawler(c As PageCrawler)
   selfCrawlers.Add c
End Sub

Private Function Exporter_canExportOnePage() As Boolean
    Exporter_canExportOnePage = False
End Function

Private Sub Exporter_export()
   Dim o As Page
   Dim fs As New FileSystemPageStore
   
   Dim extra As Page
   Set extra = myWads.store.loadRaw("CrawlerDefinitions")
   Call selfPages.addPage(extra)
   Set extra = myWads.store.loadRaw("ExportDefinitions")
   Call selfPages.addPage(extra)
   Set extra = myWads.store.loadRaw("LinkTypeDefinitions")
   Call selfPages.addPage(extra)
      
      
   If Not selfPages.hasPage("StartPage") Then
     Set extra = POLICY_getFactory().getNewPageInstance
     extra.createdDate = Date
     extra.pageName = "StartPage"
     extra.raw = "Welcome to your new SdiDesk. Happy Hypertexting!"
     Call selfPages.addPage(extra)
   End If
     
   For Each o In selfPages.pages.toCollection
     Dim p2 As Page
     Set p2 = o.spawnCopy()
    
     Dim pathName As String, fileName As String
     pathName = selfPath & "pages\"
     Call fs.ensureFullNameDirectory(pathName)
    
     pathName = pathName & Left(o.pageName, 1)
     pathName = fs.ensureTrailingSlash(pathName)
     Call fs.ensureFullNameDirectory(pathName)
    
     fileName = pathName + p2.pageName + ".mnp"
     Call fs.savePageToFile(p2, fileName)
     
   Next o
   Set fs = Nothing
End Sub

Private Sub Exporter_exportThisPage(p As Page)
    MsgBox ("not fatal error : SdiExporter can't export single pages")
End Sub

Private Function Exporter_getPages() As PageSet
  If selfPages Is Nothing Then
    Set selfPages = New PageSet
  End If
  If selfPages.size = 0 Then
    Dim o As Object, cr As PageCrawler
    For Each o In selfCrawlers
      Set cr = o
      Call cr.crawl(selfStartPage)
      selfPages.merge (cr.getPages)
    Next o
  End If
  Set Exporter_getPages = selfPages
  Set cr = Nothing
  Set o = Nothing
End Function



Private Property Let Exporter_name(ByVal RHS As String)
  selfName = RHS
End Property

Private Property Get Exporter_name() As String
  Exporter_name = selfName
End Property

Private Property Set Exporter_pages(ByVal RHS As PageSet)
  Set selfPages = RHS
End Property

Private Property Get Exporter_pages() As PageSet
  Set Exporter_pages = selfPages
End Property

Private Property Let Exporter_path(ByVal RHS As String)
  selfPath = RHS
End Property

Private Property Get Exporter_path() As String
  Exporter_path = selfPath
End Property

Private Property Let Exporter_startPage(ByVal RHS As String)
  selfStartPage = RHS
End Property

Private Property Get Exporter_startPage() As String
  Exporter_startPage = selfStartPage
End Property

Private Property Let Exporter_version(ByVal RHS As String)
  selfVersion = RHS
End Property

Private Property Get Exporter_version() As String
  Exporter_version = selfVersion
End Property

Private Property Set Exporter_wads(ByVal RHS As WikiAnnotatedDataStore)
    Set myWads = RHS
End Property

Private Property Get Exporter_wads() As WikiAnnotatedDataStore
    Set Exporter_wads = myWads
End Property
