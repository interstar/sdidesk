VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WikiAnnotatedDataStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Conceptually the WADS encapsulates all the responsibilities
' for managing a PageStore, producing native and export versions of
' wiki pages

' this is an interface which encapsulates the management
' of a PageStore, processing pages for native display
' or export etc.

' we are refactoring to this interface gradually

Public store As PageStore

Public Function pageExists(pName As String) As Boolean

End Function

Public Function getPageVar(pageName As String, varName As String) As String

End Function

Public Function getPageSetContaining(s As String) As PageSet

End Function

Public Function getRawPageData(pName As String) As String

End Function
