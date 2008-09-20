VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AllPagesPageCrawler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A PageCrawler which returns all pages

Implements PageCrawler

Private myWads As WikiAnnotatedDataStore

Private myName As String
Private myPages As PageSet
Private myStore As PageStore

Public Sub init(aName As String)
    myName = aName
End Sub


Private Sub Class_Terminate()
    Set myWads = Nothing
    Set myPages = Nothing
    Set myStore = Nothing
End Sub


Private Sub PageCrawler_clear()
    Set myPages = New PageSet
End Sub

Private Sub PageCrawler_crawl(startPage As String)
    Set myPages = myStore.AllPages
End Sub

Private Function PageCrawler_fillPageSetFromPage(p As Page) As PageSet
    Call PageCrawler_crawl("")
    Set PageCrawler_fillPageSetFromPage = PageCrawler.getPages
End Function

Private Function PageCrawler_fillPageSetFromString(s As String) As PageSet
    Call PageCrawler_crawl("")
    Set PageCrawler_fillPageSetFromString = PageCrawler.getPages
End Function

Private Function PageCrawler_getPages() As PageSet
    Set PageCrawler_getPages = myPages
End Function

Private Property Set PageCrawler_wads(ByVal RHS As WikiAnnotatedDataStore)
    Set myWads = RHS
End Property


Private Property Get PageCrawler_wads() As WikiAnnotatedDataStore
    Set PageCrawler_wads = myWads
End Property

Private Property Let PageCrawler_name(ByVal RHS As String)
    myName = RHS
End Property

Private Property Get PageCrawler_name() As String
    PageCrawler_name = myName
End Property

Private Property Set PageCrawler_pages(ByVal RHS As PageSet)
    Set myPages = RHS
End Property

Private Property Get PageCrawler_pages() As PageSet
    Set PageCrawler_pages = myPages
End Property


Private Function PageCrawler_toString() As String
    PageCrawler_toString = "'''" & myName & "'''" & _
    " is an example of an ''All''''''Pages'' crawler. It picks up all the pages in the wiki. (The list of pages you see on the AllPages page.)"
End Function
