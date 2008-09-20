VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Exporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' The new export model goes like this :
' different types of exporting are handled by separate custom programs.
' These are called "Exporters" (as opposed to "Exports")

' The "Exporter" is an abstract class / interface for the object
' which is the core of an Exporter.

' Exporter is no longer used in the main SdiDesk program.

Public Function readCommand() As VCollection
' get the command line parameters into the exporter
' returns a VCollection of keys, values

' expected
' For export :
' name - of the export,
' param - the page containing parameters
' psi - page-store identifier (how to find the PageStore)

' For instant export
' page - the pageName to be exported
' psi - page-store identifier (how to find the PageStore)

End Function

Public Function getVersion() As String
' get the version number of this program
End Function

Public Function canInstant() As Boolean
' does it make sense to export a single page (by hitting the "export"
' button) with this Exporter.
' Eg. HtmlExporter? Yes, sometimes we want just one HTML page
' TgwbExporter? probably not, a TouchGraph map only makes sense for a PageSet
End Function

Public Function canMulti() As Boolean
' opposite of canInstant
' Does it make sense to try to call this Exporter from an Export
' ie. group of pages
' eg. HtmlExporter, Yes,
'
End Function

Public Sub doExport(exp As Export)
' actually do the export. By the end of this, the relevant pages will have
' been exported in another format and will be sitting on the file system
' at the end of path

End Sub

Public Sub doInstantExport(page As page)
' do an export of a single page (if it's possible)
' we have both options so that exporters might be able to export
' dynamically created pages, (not in the PageStore)
End Sub


Public Sub doInstantExportPageName(pageName As String)
' do an export of a single page (if it's possible)
End Sub


