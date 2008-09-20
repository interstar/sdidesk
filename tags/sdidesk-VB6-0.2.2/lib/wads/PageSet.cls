VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PageSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A PageSet is a collection of pages
' You can fill it with pages and search them

Public pages As OCollection

Public Sub init()
  Set pages = New OCollection
End Sub

Public Sub addPage(p As Page)
  If pages.hasKey(p.pageName) Then
    pages.Remove (p.pageName)
    Call pages.Add(p, p.pageName)
  Else
    Call pages.Add(p, p.pageName)
  End If
End Sub

Public Sub addPageFromName(pageName As String, store As PageStore)
   Dim p As Page
   Set p = store.loadRaw(pageName)
   Call Me.addPage(p)
End Sub

Public Function hasPage(pName As String) As Boolean
  If pages.hasKey(pName) Then
    hasPage = True
  Else
    hasPage = False
  End If
End Function

Public Sub removePage(pName As String)
  If hasPage(pName) Then
    Call pages.Remove(pName)
  End If
End Sub

Public Sub clearOut()
  Set pages = New OCollection
End Sub

Public Sub merge(ps2 As PageSet)
  Dim size As Integer
  Dim o As Object
  For Each o In ps2.pages.toCollection
    Call Me.addPage(o)
  Next o
End Sub

Public Function toWikiMarkup() As String
  Dim s As String
  s = ""
  Dim i As Object
  For Each i In pages.toCollection
    s = s & "* [[" & i.pageName & "]]" & vbCrLf
  Next i
  toWikiMarkup = s
End Function

Public Sub saveAll(store As PageStore)
  Dim i As Object
  For Each i In pages.toCollection
    Call store.savePage(i)
  Next i
End Sub

Public Function size() As Integer
  size = pages.count
End Function

