VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Export"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Export is a an interface for "exports"
' NOT the same as "Exporter" which is the interface for exporters.

' An Export is an application of an exporter program to a particular
' page set. It's defined by an ExportRecord containing three fields :

' A name
' An Exporter (the external program which will do the exporting)
' A parameter page, which contains further parameters for this export

' For example, you may have a set of pages in SdiDesk which represent
' your personal web-site. An HtmlExporter can take pages and render
' them as flat HTML. However you want to export the entire site in
' one go. (It's quicker, and all the links *between* pages can be
' preserved)

Public Sub readParameters(paramPage As String)
' this takes the name of a page containing parameters for the export
' after this sub is run, the export should know it's parameters
End Sub

Public Sub addCrawler(pc As PageCrawler)
' add another crawler to the Export
' NB : this means we expect Exports to be able to be defined by
' a number of PageCrawlers
End Sub

Public Function getPageSet()
' get the set of pages which are cought by this Export
' normally it follows it's crawlers
' this function may be the function which triggers the crawling,
' or may simply return something that was crawled as and when
' the crawlers were added.
End Function

Public Sub doExport()
' go do it
End Sub



