VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RSSExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Exporter

' This is another Exporter, designed to allow you to export a set of pages
' as an RSS 2.0 file,
' mindlessly simple, but at least it lets you get the stuff out in
' RSS / XML which is all anyone cares about these days.

Private selfModel As ModelLevel ' link to the model level
Private selfName As String ' my name
Private selfVersion As String ' my version
Private selfPath As String ' my path
Private selfStartPage As String ' my copy of the start page

Private selfCrawlers As Collection ' the collection of crawlers that define this export
Private selfPages As PageSet ' the collection of pages to be exported

Private selfCutter As RssCutter ' the template for the Rss

Private mg As WikiMarkupGopher

Public Sub init(m As ModelLevel, pageName As String)
  Set selfModel = m
  Set selfCutter = New RssCutter
  Call selfCutter.init(pageName, selfModel.store, selfModel.prep)
  Set mg = New WikiMarkupGopher
End Sub


Public Function saveStringInFile(s As String, fileName As String)
   Dim st As New StringTool
   On Error GoTo couldNotOpen
       Open fileName For Output As 1
       Print #1, s
       Close #1
       Set st = Nothing
       Exit Function

couldNotOpen:
   MsgBox ("sorry, couldn't open " + fileName + " to write to, this page wasn't exported")
   Set st = Nothing

End Function


Public Function rssHeader(title As String, link As String, description As String)
   Dim rss As String
   rss = "<rss version=" + mg.qq("2.0") + ">" + vbCrLf
   rss = rss + "<channel>" + vbCrLf
   rss = rss + "<title>" + selfCutter.channelTitle + "</title>" + vbCrLf
   rss = rss + "<link>" + selfCutter.channelLink + "</link>" + vbCrLf
   rss = rss + "<description>" + selfCutter.channelDescription + "</description>"
   rssHeader = rss
End Function

Public Function rssItem(p As Page)
    Call selfModel.prep.prepare(p, False)
    Dim item As String
    item = "<item>" + vbCrLf
    item = item + "<title>" + p.pageName + "</title>" + vbCrLf
    item = item + "<category>" + p.categories + "</category>" + vbCrLf
    item = item + "<pubdate>" + CStr(p.lastEdited) + "</pubdate>" + vbCrLf
    item = item + "<description>" + p.prepared + "</description>" + vbCrLf
    item = item + "</item>" + vbCrLf
    rssItem = item
End Function

Public Sub export(ps As PageSet, fileName As String)
   
End Sub


Private Sub Class_Initialize()
   Set selfCrawlers = New Collection
End Sub

Private Sub Class_Terminate()
   Set selfCrawlers = Nothing
   Set selfPages = Nothing
   Set selfModel = Nothing
   Set selfCutter = Nothing
End Sub

Private Sub Exporter_addCrawler(c As PageCrawler)
   selfCrawlers.Add c
End Sub

Private Sub Exporter_export()
   Dim o As Object, p As Page
   Dim rss As String, build As String
   
   build = ""
   Call Exporter_getPages ' ensure the PageSet is filled
   
   build = build + Me.rssHeader(selfCutter.channelTitle, selfCutter.channelLink, selfCutter.channelDescription)
   For Each o In selfPages.pages
     build = build + rssItem(o) + vbCrLf
   Next o
   
   build = build + "</channel></rss>" + vbCrLf

   Dim fileName As String
   fileName = selfPath + "\sdiDesk.rss"
   
   Open fileName For Output As 1
   Print #1, build
   Close #1
   
   Set o = Nothing
   
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

Private Property Set Exporter_model(ByVal RHS As ModelLevel)
   Set selfModel = RHS
End Property

Private Property Get Exporter_model() As ModelLevel
  Set Exporter_model = selfModel
End Property

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

