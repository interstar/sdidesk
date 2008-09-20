VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StandardPagesManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' this object just encapsulates the responsibility of ensuring
' a set of default pages to kick off your new SdiDesk
' it makes RecentChanges, AllPages, some LinkTypeDefinitions,
' CrawlerDefinitions, ExportDefinitions and a default BasicHtmlTemplate

Public Function ensurePage(store As PageStore, pageName As String, defaultRaw As String)
    Dim p As Page
    Set p = POLICY_getFactory().getNewPageInstance
    
    If Not store.pageExists(pageName) Then
      p.pageName = pageName
      p.raw = defaultRaw
  
      Call store.savePage(p)
    End If
    Set p = Nothing
End Function

Public Function ensureStandardPages(store As PageStore)
    Dim s As String
    Call ensurePage(store, "RecentChanges", "##RecentChanges")
    Call ensurePage(store, "AllPages", "##AllPages")
    Call ensurePage(store, "LinkTypeDefinitions", "Type,, Colour" + vbCrLf + "____" + _
    vbCrLf + "example,, #0066aa" + vbCrLf + "explanation,, #ff9900" + _
    vbCrLf + "definition,, #339966" + vbCrLf + "counter,, #aa6633" + _
    vbCrLf + "normal,, #000099")
    
    s = "name,, type,, maxDepth,, excluded pages,, excluded link types" + _
    vbCrLf + "____" + _
    vbCrLf + "simple_recursive ,, recursive,, -1,, ,, " + _
    vbCrLf + "depth_one,, recursive,, 1,, ,, ,, " + _
    vbCrLf + "recent_changes,, recent,, ,, ,, ,,"
    
    Call ensurePage(store, "CrawlerDefinitions", s)
    
    s = "Name,, Program,, Parameters" + _
    vbCrLf + "____" + vbCrLf + _
    "Dummy,, DummyExporter.exe,, Exports/MySite" + vbCrLf
    
    Call ensurePage(store, "ExportDefinitions", s)
    
    s = "<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN' " + _
    vbCrLf + "'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'>" + _
    vbCrLf + "<html xmlns='http://www.w3.org/1999/xhtml'>" + _
    vbCrLf + "<head>" + vbCrLf + _
    vbCrLf + "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1' />" + _
    vbCrLf + "<meta name='generator' content='SdiDesk 0.2.0' />" + _
    vbCrLf + "</head>" + vbCrLf + "----" + _
    vbCrLf + "<body bgcolor='#ffffff' text='#000000'>" + _
    vbCrLf + "----" + vbCrLf + "</body> </html>"
    
    Call ensurePage(store, "BasicHtmlTemplate", s)
    
End Function


