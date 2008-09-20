VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExportHtmlPageCooker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements PageCooker

' This class is the concrete implementation of the abstract / interface
' class : PageCooker

' This is the page cooker which turns the pages into HTML for export

Private myWikiToHtml As WikiToHtml ' to do all the HTML work,
Private myLinkProcessor As StandardLinkProcessor ' to parse links
Private myLinkWrapper As ExportHtmlLinkWrapper  ' to wrap links

Public Sub setPageSet(ps As PageSet)
    Set myLinkWrapper.exportedPages = ps  ' we need a pageset of all pages we're
    ' exporting, so we know which links are real or not
End Sub

Public Function asPageCooker() As PageCooker
    Set asPageCooker = Me
End Function



Private Function PageCooker_cook(aPage As page) As String
   Dim build As String, intermediate As String
   build = ""
   build = build + myWikiToHtml.mainTransform(aPage.prepared, myLinkProcessor, myLinkWrapper)
   PageCooker_cook = build
End Function

Private Function PageCooker_cookObject(aPage As page) As Object
    Dim s As New StringTool ' dummy object
    Set PageCooker_cookObject = s
End Function


Private Property Set PageCooker_LinkProcessor(ByVal RHS As LinkProcessor)
    Set myLinkProcessor = RHS
End Property

Private Property Get PageCooker_LinkProcessor() As LinkProcessor
    Set PageCooker_LinkProcessor = myLinkProcessor
End Property

Private Property Set PageCooker_LinkWrapper(ByVal RHS As LinkWrapper)
    Set myLinkWrapper = RHS
End Property

Private Property Get PageCooker_LinkWrapper() As LinkWrapper
    Set PageCooker_LinkWrapper = myLinkWrapper
End Property

Private Property Set PageCooker_wads(ByVal m As WikiAnnotatedDataStore)
   Set myWads = m
End Property

Private Property Get PageCooker_wads() As WikiAnnotatedDataStore
   Set PageCooker_wads = myWads
End Property

