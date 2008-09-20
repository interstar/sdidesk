VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExportSubsystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' interface for the export subsystem
Public pageStoreIdentifier As String
' this is the string the ExportManager will pass to any
' export programs so they can find the PageStore
' Currently the main data directory
' though may later be a URL

Public Sub refreshExportManager(wads As WikiAnnotatedDataStore)
' reload the details from the PageStore
' if the definitions of exports have been updated.
End Sub

Public Function makeExportsPage() As Page
' create a page that's a list of currently available exports
End Function

Public Function makeExportersPage() As Page
' create a page that details the export programs available
End Function

Public Function makeChooseExporterPage(currentPageName As String) As Page
' when the user wants to export the current page, a list of
' exporter programs to choose from
End Function

Public Sub scanForExports()
' scans the local drive for export plug-ins
End Sub

Public Sub doExport(name As String)
' fire off the export
End Sub

Public Sub doInstantExport(exporterName As String, pageName As String)
' fire off an instant export
End Sub
