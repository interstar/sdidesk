VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PageStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' The PageStore is an interface to an object which hides the storage and
' searching of pages.

' It knows how to save a page somewhere and fetch it

' Also to bring list of all pages or search for pages with some criterion

Public pictureLocality As String  ' the protocol and location for pictures

Public Function getPageStoreIdentifier() As String
' string which represents the "address" of this page-store
' eg. a directory or URL
End Function

Public Function pageExists(pageName As String) As Boolean
End Function

Public Function safeDate(s As String) As Date
End Function

Public Function loadRaw(pageName As String) As Page
' load a page but don't prepare or cook it, hence page is still raw
End Function

Public Function loadUntilNotRedirectRaw(pageName As String) As Page
End Function

Public Function loadOldPage(pageName As String, version As Integer) As Page
End Function

Public Sub savePage(p As Page)
End Sub

Public Function deletePage(pageName As String)
End Function

Public Function loadMonth(month As Integer, year As Integer) As String
' loads a page content containing time data
End Function

Public Sub saveMonth(month As Integer, year As Integer, body As String)
End Sub

Public Function timeIndexAsWikiFormat(month As Integer, year As Integer, order As Boolean)
End Function

Public Function pageContains(pageName As String, searchText As String) As Boolean
End Function

Public Function getPageSetOfAllPagesStartingWith(s As String) As PageSet
End Function

Public Function AllPages() As PageSet
End Function

Public Function getPageSetContaining(searchText) As PageSet
End Function
