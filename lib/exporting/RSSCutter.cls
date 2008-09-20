VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RssCutter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This is an object which gets information for an RSS file
' from a page

' It also knows how to parse a string from a page

Public channelTitle As String
Public channelName As String

Public channelDescription As String
Public channelLink As String
Public channelAuthor As String

Public Function getVar(p As Page, varName As String, defaultVal As String)
   If p.hasVar(varName) Then
      getVar = p.getVal(varName)
   Else
      getVar = defaultVal
   End If
End Function

Public Sub init(pageName As String, store As PageStore, prep As PagePreparer)
  ' expects all data to be set on the page
  Dim p As Page
  Set p = store.loadRaw(pageName)
  Call prep.prepare(p, False)
  channelName = getVar(p, "channelName", "SdiDesk")
  channelTitle = getVar(p, "channelTitle", "SdiDesk")
  
  channelDescription = getVar(p, "channelDescription", "Pages from my wiki")
  channelLink = getVar(p, "channelLink", "http://www.nooranch.com/")
  channelAuthor = getVar(p, "channelAuthor", "")
  ' etc.
End Sub



