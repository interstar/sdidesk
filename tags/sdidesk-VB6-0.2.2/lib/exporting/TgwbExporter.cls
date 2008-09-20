VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TgwbExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Exporter

' TgwbExporter
' Exports a file suitable for TouchGraph Wiki Browser

Private myWads As WikiAnnotatedDataStore

Private selfName As String ' my name
Private selfVersion As String ' my version
Private selfPath As String ' my path
Private selfStartPage As String ' my copy of the start page

Private crawlers As Collection ' the collection of crawlers that define this export
Private selfPages As PageSet ' the collection of pages to be exportet

Private myCrawler As RecursivePageCrawler

Public Function saveStringInFile(s As String, fileName As String)
   Open fileName For Output As 1
   Print #1, s
   Close #1
End Function

Public Function pageToLine(pageName As String) As String
  ' takes the name of a page, and returns a line, listing all out-links from the page
  Dim ps As PageSet
  
  Dim build As String
  build = pageName & " "
  
  Set ps = myCrawler.asPageCrawler.fillPageSetFromString(myWads.getRawPageData(pageName))
  
  Dim o As Page
  For Each o In ps.pages.toCollection
    build = build + o.pageName + " "
  Next o
  pageToLine = build
  
  Set ps = Nothing
End Function


Private Sub Class_Initialize()
   Set crawlers = New Collection
   Set myWads = Nothing
   Set myCrawler = POLICY_getFactory().getNewPageCrawlerInstance("recursive", "tgwbCrawler", 1, "", "")
End Sub

Private Sub Class_Terminate()
   Set crawlers = Nothing
   Set selfPages = Nothing
   Set myCrawler = Nothing
End Sub

Private Sub Exporter_addCrawler(c As PageCrawler)
   crawlers.Add c
End Sub

Private Function Exporter_canExportOnePage() As Boolean
    Exporter_canExportOnePage = False
End Function

Private Sub Exporter_export()
   Dim o As Object
   Dim build As String
   build = ""
   Call Exporter_getPages ' ensure the PageSet is filled
   For Each o In selfPages.pages.toCollection
     build = build + pageToLine(o.pageName) + vbCrLf
   Next o

   Dim fileName As String
   
   fileName = selfPath + "\sdiWiki.txt"
   Open fileName For Output As 1
   Print #1, build
   Close #1
   
   Set o = Nothing
End Sub

Private Sub Exporter_exportThisPage(p As Page)
    MsgBox ("Non terminal error : attempt to export one page via the TouchGraph Wiki Browser exporter. This option is not available")
End Sub

Private Function Exporter_getPages() As PageSet
  If selfPages Is Nothing Then
    Set selfPages = New PageSet
  End If
  If selfPages.size = 0 Then
    Dim o As Object, cr As PageCrawler
    For Each o In crawlers
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
