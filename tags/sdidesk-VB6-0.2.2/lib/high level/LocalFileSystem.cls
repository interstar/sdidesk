VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LocalFileSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' an interface for the model level to deal with the local file
' system, eg. make directory page

'still a bit of a mess because ...

' makeDirectoryPage is used for looking at a local directory
' whereas changeDirectory is really about where the PageStore is going
' to put things, doesn't make sense if there's a remote PageStore
' (as I hope there will be one day)

Public Function hasLocalFileSystem() As Boolean
End Function

Public Function makeDirectoryPage(path As String) As Page
End Function

Public Sub changeDirectory(path As String)
End Sub

Public Function getMainDataDirectory() As String
End Function

Public Function getExporterDirectory() As String
' gets the exporter directory
End Function

Public Function getDirectoryListingAsVCollection(d As String) As VCollection
' returns a vcollection of directory d
End Function
