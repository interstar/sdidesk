VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PageCrawler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' The purpose of the PageCrawler is to go around gathering pages
' into a PageSet according to a strategy
' now an interface


Public wads As WikiAnnotatedDataStore ' where to get pages etc.

Public pages As PageSet ' where we keep the pages while crawling

Public name As String ' useful to know the name of the crawler

Public Sub clear()
' clear the number of pages crawled
End Sub

Public Function getPages() As PageSet
' return the current pages crawled in a pageset
End Function

Public Function fillPageSetFromString(s As String) As PageSet
' start with a string representation of the body of a page,
' and get all pages linked from it
End Function

Public Function fillPageSetFromPage(p As Page) As PageSet
' this starts with a page and fills the pages from it.
' handles networks differently etc.
End Function

Public Function toString() As String
' return a string representation of the pages crawled
End Function

Public Sub crawl(startPage As String)
' do the crawl
End Sub
