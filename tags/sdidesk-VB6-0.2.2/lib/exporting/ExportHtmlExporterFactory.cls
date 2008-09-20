VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExportHtmlExporterFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Concrete implementation of the ExporterFactory interface
' for HtmlExporter

Private myWads As WikiAnnotatedDataStore
Private mySysConf As SystemConfigurations
Private myInterMap As interWikiMap

Private myLinkProcessor As StandardLinkProcessor
Private myLinkWrapper As ExportHtmlLinkWrapper
Private myPageCooker As ExportHtmlPageCooker
Private myPagePreparer As PagePreparer

Implements ExporterFactory

Public Function asExporterFactory() As ExporterFactory
    Set asExporterFactory = Me
End Function

Private Function ExporterFactory_getLinkProcessor() As LinkProcessor
    Set ExporterFactory_getLinkProcessor = myLinkProcessor
End Function

Private Function ExporterFactory_getLinkWrapper() As LinkWrapper
    Set ExporterFactory_getLinkWrapper = myLinkWrapper
End Function

Private Function ExporterFactory_getPageCooker() As PageCooker
    Set ExporterFactory_getPageCooker = myPageCooker
End Function

Private Function ExporterFactory_getPagePreparer() As PagePreparer
    Set ExporterFactory_getPagePreparer = myPagePreparer
End Function

Private Sub ExporterFactory_init(wads As WikiAnnotatedDataStore, sysConf As SystemConfigurations, iMap As interWikiMap)
    
    Set myPageCooker = New ExportHtmlPageCooker
    Set myLinkWrapper = New ExportHtmlLinkWrapper
    Set myPagePreparer = New PagePreparer
    Set myLinkProcessor = New StandardLinkProcessor
    
    Set myPagePreparer.wads = wads
    
    Set myLinkWrapper.asLinkWrapper.remoteInterMap = iMap
    Set myLinkWrapper.asLinkWrapper.remoteSysConf = sysConf
    Set myLinkWrapper.asLinkWrapper.remoteWads = wads

    Set myPageCooker.asPageCooker.LinkProcessor = myLinkProcessor
    Set myPageCooker.asPageCooker.LinkWrapper = myLinkWrapper
    
End Sub
