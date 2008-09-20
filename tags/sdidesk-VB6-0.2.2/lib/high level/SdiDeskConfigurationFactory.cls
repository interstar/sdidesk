VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SdiDeskConfigurationFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Dependency Injection

' essentially, this factory creates the appropriate objects for
' use within SdiDesk

' ideally, no other component of the system should make decisions about
' what concrete class ever implements a particular abstract class

Private myPagePreparer As PagePreparer
Private myNativePageCooker As NativePageCooker

Private myPageStore As FileSystemPageStore

Private myLinkTypeManager As LinkTypeManager
Private myCrawlerDefinitionTable As CrawlerDefinitionTable
Private myExportManager As ExportManager

Private myModelLevel As ModelImplementation

Private myNativeLinkWrapper As NativeLinkWrapper
Private myStandardLinkProcessor As StandardLinkProcessor

Public Function getNativePageCooker() As PageCooker
    Set getNativePageCooker = myNativePageCooker
End Function

Public Function getPagePreparer() As PagePreparer
    Set getPagePreparer = myPagePreparer
End Function

Public Function getNewPageInstance() As Page
    Dim p As New MemoryResidentPage
    Set getNewPageInstance = p
End Function

Public Function getNewPageCrawlerInstance(crawlerType As String, name As String, maxDepth As Integer, exPag As String, exTyp As String) As PageCrawler
    Dim base As PageCrawler
    Select Case crawlerType:
    Case "recursive"
        Dim c As New RecursivePageCrawler
        Call c.init(name, maxDepth, exPag, exTyp)
        Set base = c
    Case "recent"
        Dim c2 As New TimeBasedPageCrawler
        Call c2.init(name, exPag, exTyp)
        Set base = c2
    Case "all"
        Dim c3 As New AllPagesPageCrawler
        Call c3.init(name)
        Set base = c3
    Case Else:
        MsgBox ("SdiDeskConfigurationFactory didn't recognise a crawler type called '" + crawlerType + "'")
    End Select
    
    Set base.wads = myModelLevel
    Set getNewPageCrawlerInstance = base
    Set c = Nothing
    Set c2 = Nothing
    Set c3 = Nothing
End Function

Public Function getModelLevel() As ModelLevel
    Set getModelLevel = myModelLevel
End Function

Public Function getNativeLinkWrapper() As LinkWrapper
    Set getNativeLinkWrapper = myNativeLinkWrapper
End Function

Public Function getStandardLinkProcessor() As LinkProcessor
    Set getStandardLinkProcessor = myStandardLinkProcessor
End Function


Public Function getNewPageStore(psi As String) As PageStore
    Set myPageStore = New FileSystemPageStore
    Call myPageStore.setDataDirectory(psi)
    Set getNewPageStore = myPageStore
End Function

Private Sub Class_Initialize()
    Set myPagePreparer = New PagePreparer
    Set myNativePageCooker = New NativePageCooker
    Set myPageStore = New FileSystemPageStore
    ' default data directory for myPageStore is App.path
    ' if we want to change that, uncomment the next line and change
    ' argument to the desired path
    ' Call myPageStore.setDataDirectory(altPath)
    
    ' default pages
    Dim spm As New StandardPagesManager
    Call spm.ensureStandardPages(myPageStore)
    Set spm = Nothing
    
    ' make sure the model has the page-store
    ' doing this via the WADS interface of the model
    Set myModelLevel = New ModelImplementation
    Dim model As ModelLevel
    Dim localWads As WikiAnnotatedDataStore
    Set model = myModelLevel
    Set localWads = model.getWikiAnnotatedDataStore
    Set localWads.store = myPageStore
    
    ' set up the model's page-preparer and native-page-cooker (chef)
    ' and their backlinks to the wads.
    Call model.setPagePreparer(myPagePreparer)
    Call model.setPageCooker(myNativePageCooker)
         
    Set myPagePreparer.wads = myModelLevel
  
    Dim sysConf As SystemConfigurations
    Set sysConf = Me.getModelLevel.getSystemConfigurations
    
    sysConf.configPage = "ConfigPage"
    sysConf.startPage = "StartPage"
    sysConf.helpIndexPage = "HelpIndex"
    sysConf.allPage = "AllPages"
    sysConf.recentChangesPage = "RecentChanges"
    
    Call linkDefs
    Call crawlerDefs
    Call interWikiMap
    
    Call nativeLinkWrapperAndStandardLinkProcessor
    
    Set myNativeLinkWrapper.asLinkWrapper.remoteInterMap = Me.getModelLevel.getSystemConfigurations.interMap

End Sub

Public Sub nativeLinkWrapperAndStandardLinkProcessor()
    Set myNativeLinkWrapper = New NativeLinkWrapper
    Set Me.getNativeLinkWrapper.remoteSysConf = Me.getModelLevel.getSystemConfigurations
    Set Me.getNativeLinkWrapper.remoteWads = Me.getModelLevel.getWikiAnnotatedDataStore
    
    Set myStandardLinkProcessor = New StandardLinkProcessor
    Set myNativePageCooker.myLinkProcessor = myStandardLinkProcessor
    Set myNativePageCooker.myLinkWrapper = myNativeLinkWrapper
    
End Sub

Private Sub linkDefs()
    ' load the link definitions from page and set up the table
    Dim linkDefs As String
    linkDefs = Me.getModelLevel.getWikiAnnotatedDataStore.getRawPageData("LinkTypeDefinitions")
    
    Set myLinkTypeManager = New LinkTypeManager
    Call myLinkTypeManager.setupLinkTypes(linkDefs)

    ' give it to the model
    Dim localSysConf As SystemConfigurations
    Set localSysConf = Me.getModelLevel.getSystemConfigurations
    
    Call localSysConf.setLinkTypeManager(myLinkTypeManager)
End Sub

Private Sub interWikiMap()
    Dim interMap As String
    interMap = Me.getModelLevel.getWikiAnnotatedDataStore.getRawPageData("InterMap")
    Dim lines() As String
    Dim parts() As String
    Dim v As Variant
    lines = Split(interMap, vbCrLf)
    For Each v In lines
        parts = Split(CStr(v), " ")
        If UBound(parts) > 0 Then
            Call Me.getModelLevel.getSystemConfigurations.interMap.add(parts(1), parts(0))
        End If
    Next v
    
End Sub

Private Sub crawlerDefs()
    ' load crawler definitions from page and set up table
    Dim crawlDefs As String
    crawlDefs = Me.getModelLevel.getWikiAnnotatedDataStore.getRawPageData("CrawlerDefinitions")
    
    Set myCrawlerDefinitionTable = New CrawlerDefinitionTable
    Call myCrawlerDefinitionTable.parseFromTableString(crawlDefs, Me.getModelLevel, myPageStore, myNativePageCooker)
    
    ' give it to model
    Dim localCrawlerSubsystem As CrawlerSubsystem
    Set localCrawlerSubsystem = Me.getModelLevel.getCrawlerSubsystem
    
    Set localCrawlerSubsystem.crawlerManager = myCrawlerDefinitionTable
       
End Sub

Public Function wrapPageInNetwork(p As Page) As Network
    If typeName(p) = "Network" Then
        Set wrapPageInNetwork = p
    Else
        Dim n As New Network
        Set n.innerPage = p
        Call n.init(1, 200, 0.75)
        n.parseFromPrettyPersist (p.raw)
        p.cooked = "A network"
        p.pageType = "network"
        Set wrapPageInNetwork = n
    End If
End Function
