VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DummyWads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Dummy Wads, used for writing tests

Implements WikiAnnotatedDataStore

Private Function WikiAnnotatedDataStore_getPageSetContaining(s As String) As PageSet
    MsgBox ("Error, no page-set in the DummyWads")
    End
End Function

Private Function WikiAnnotatedDataStore_getPageVar(pageName As String, varName As String) As String
    WikiAnnotatedDataStore_getPageVar = "dummyVal"
End Function

Private Function WikiAnnotatedDataStore_getRawPageData(pName As String) As String
    If pName = "HelloWorld" Then
        WikiAnnotatedDataStore_getRawPageData = "Another green world"
    Else
        WikiAnnotatedDataStore_getRawPageData = "blah blah blah"
    End If
End Function

Private Function WikiAnnotatedDataStore_pageExists(pName As String) As Boolean
    If pName = "NonExistant" Then
        WikiAnnotatedDataStore_pageExists = False
    Else
        WikiAnnotatedDataStore_pageExists = True
    End If
End Function

Private Property Set WikiAnnotatedDataStore_store(ByVal RHS As PageStore)
End Property

Private Property Get WikiAnnotatedDataStore_store() As PageStore
    MsgBox ("Error, no page-store yet in the DummyWads")
    End
End Property

