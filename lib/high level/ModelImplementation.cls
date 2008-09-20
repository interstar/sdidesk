VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ModelImplementation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' The "Model" part of an MVC conception of SdiDesk

' It has (almost) no responsibility for call-back to rest of system
' or the ''dynamics'' of control as the user flows through the
' system

' now (as of March 2005) the model level implements various
' interfaces

' Currently

Implements ModelLevel

Implements WikiAnnotatedDataStore
Implements SingleUserState
Implements SystemConfigurations
Implements CrawlerSubsystem
Implements ExportSubsystem
Implements ControllableModel
Implements LocalFileSystem


Public mainForm As Form ' this is the WADSMainForm which calls it
' and is needed for the occasional callback

Public prep As PagePreparer ' what transforms raw into prepared
Public chef As PageCooker ' what transforms prepared into cooked

' support for WikiAnnotatedDataStore interface
Public myStore As PageStore ' stores the pages

' support for SingleUserState interface
Private myBacklinks As Boolean
Private myCurrentPageName As String
Private myOldPageName As String
Private myPageEditState As PageEditState
Private myIsLoading As Boolean
Private myHistory As NavigationHistory
Private myChangesSaved As Boolean
Private myCurrentPage As page

' support from SystemConfiguration interface
Private linkTypeMan As LinkTypeManager ' stores info. about link types and colours
Private myConfigPage As String
Private myStartPage As String
Private myHelpIndexPage As String
Private myAllPage As String
Private myRecentChangesPage  As String

Private myInterWikiMap As interWikiMap

' support for CrawlerSubsystem interface
Private myCrawlerMan As CrawlerDefinitionTable ' creates and stores the crawlers

' support for ExportSubsystem interface
Private myExportMan As ExportManager ' stores info. about export scripts
Private myPageStoreIdentifier As String ' where is the PageStore?

' useful stuff
Private st As StringTool

Public Sub loadConfigs()
  Dim confPage As page
  confPage = WikiAnnotatedDataStore_getRawPageData("RootConfig")
End Sub


Public Function qq(s As String) As String
  qq = Chr(34) + s + Chr(34)
End Function


Public Function asModelLevel() As ModelLevel
    Set asModelLevel = Me
End Function





Public Function loadPage(pageName As String) As String
   ' This is a lower level function than loadNewPage,
   ' backPage and forwardPage
   ' It just manages the loading
   ' Use loadNewPage and backPage and forwardPage
   ' when you want the package deal
   
   Set myCurrentPage = Me.myStore.loadRaw(pageName)
   Call myCurrentPage.cook(prep, chef, myBacklinks)
   myCurrentPageName = myCurrentPage.pageName
   loadPage = "normal"
   If myCurrentPage.isNetwork Then loadPage = "network"
   If myCurrentPage.isTable Then loadPage = "table"
   If myCurrentPage.isRedirect Then loadPage = "redirect"
   If myCurrentPage.isNew Then loadPage = "new page"
End Function

Public Function pageExists(pageName As String) As Boolean
   pageExists = myStore.pageExists(pageName)
End Function

Public Function pageContains(pageName As String, s As String) As Boolean
  Dim r As String
  r = WikiAnnotatedDataStore_getRawPageData(pageName)
  If InStr(r, s) > 0 Then
    pageContains = True
  Else
    pageContains = False
  End If
End Function



Public Function loadRawPage(pageName As String) As String
   Set myCurrentPage = myStore.loadRaw(pageName)
   myCurrentPageName = myCurrentPage.pageName
   loadRawPage = myCurrentPage.getMyType
   
'   Call loadPage(pageName)
     
End Function






Public Sub savePage()
   Call myStore.savePage(myCurrentPage)
   
' now special behaviours
   If myCurrentPage.pageName = "LinkTypeDefinitions" Then
     ' we've saved new link types, so reload them
     Call linkTypeMan.setupLinkTypes(myCurrentPage.raw)
   End If
      
   If myCurrentPage.pageName = "CrawlerDefinitions" Then
     ' we've saved new crawler defs, so reload them
     Call myCrawlerMan.parseFromTableString(myCurrentPage.raw, Me, myStore, chef)
   End If
   
   If myCurrentPage.pageName = "ExportDefinitions" Then
     ' we've saved new export defs, so reload them
     Call myExportMan.parseFromRawString(myCurrentPage.raw)
   End If
   
End Sub

Public Function getCurrentPage() As page
  Set getCurrentPage = myCurrentPage
End Function

Public Sub setCurrentPage(p As page)
  Set myCurrentPage = p
End Sub




Public Function find(searchString As String) As String
  Dim ps As PageSet
  Set ps = myStore.getPageSetContaining(searchString)
  find = ps.toWikiMarkup
End Function





Private Sub Class_Initialize()
    Set st = New StringTool
    Set myHistory = New NavigationHistory
    Set myExportMan = New ExportManager
    Set myInterWikiMap = New interWikiMap
    myBacklinks = False
End Sub

Private Sub Class_Terminate()
    Set st = Nothing
    Set myExportMan = Nothing
    Set myInterWikiMap = Nothing
End Sub

Private Sub ControllableModel_deletePage(pageName As String)
   Call myStore.deletePage(pageName)
End Sub

Private Function ControllableModel_getCrawlerSubsystem() As CrawlerSubsystem
    Set ControllableModel_getCrawlerSubsystem = Me
End Function

Private Function ControllableModel_getCurrentPage() As page
    If myCurrentPage.isNetwork And typeName(myCurrentPage) <> "Network" Then
        Dim net As New Network
        Set net = POLICY_getFactory().wrapPageInNetwork(myCurrentPage)
        Set myCurrentPage = net
    End If
    Set ControllableModel_getCurrentPage = myCurrentPage
End Function

Private Function ControllableModel_getExportSubsystem() As ExportSubsystem
    Set ControllableModel_getExportSubsystem = Me
End Function

Private Function ControllableModel_getLocalFileSystem() As LocalFileSystem
    Set ControllableModel_getLocalFileSystem = Me
End Function

Private Function ControllableModel_getPageCooker() As PageCooker
    Set ControllableModel_getPageCooker = Me.chef
End Function

Private Function ControllableModel_getPagePreparer() As PagePreparer
    Set ControllableModel_getPagePreparer = prep
End Function

Private Function ControllableModel_getSingleUserState() As SingleUserState
    Set ControllableModel_getSingleUserState = Me
End Function

Private Function ControllableModel_getSystemConfigurations() As SystemConfigurations
    Set ControllableModel_getSystemConfigurations = Me
End Function

Private Function ControllableModel_getWikiAnnotatedDataStore() As WikiAnnotatedDataStore
    Set ControllableModel_getWikiAnnotatedDataStore = Me
End Function

Private Function ControllableModel_loadNewPage(pageName As String) As String
   ' this is the higher level way of getting a page
   ' that updates the history
   ' and handles redirects
   Dim pageType As String
   Dim flag As Boolean
   flag = False
   Do
      pageType = loadPage(pageName)
      If myCurrentPage.isRedirect() Then
        pageName = myCurrentPage.getRedirectPage()
      Else
        flag = True
      End If
   Loop While flag = False
   myCurrentPageName = pageName
End Function

Private Function ControllableModel_loadRawPage(pageName As String) As String
    ControllableModel_loadRawPage = loadRawPage(pageName)
End Function

Private Function ControllableModel_makeHistoryPage(pageName As String) As page
   Dim f1 As String
   Dim s As String, s2 As String, colHex As String
   Dim p As page, p1 As page
   Dim i As Integer
   s = "== Current ==" + vbCrLf
   s = s + WikiAnnotatedDataStore_getRawPageData(pageName) + vbCrLf + "----" + vbCrLf
   For i = 1 To 5
     Set p1 = myStore.loadOldPage(pageName, i)
     colHex = Hex(16 - (2 * i))
     colHex = colHex & colHex & colHex
     colHex = colHex & colHex
     s2 = "#NoWiki" & vbCrLf
     s2 = s2 & "<table width=100% bgcolor='#" & colHex & "'>"
     s2 = s2 & vbCrLf & "#Wiki" & vbCrLf
     s2 = s2 & vbCrLf & "<tr><td>" & vbCrLf
     s2 = s2 & "== Version -" & i & " (" & p1.lastEdited & ") ==" & vbCrLf
     s2 = s2 & "[[#revert " & p1.pageName & " " & i & "]]" & vbCrLf
     s2 = s2 & p1.raw & vbCrLf
     s2 = s2 & "</td></tr></table>" & vbCrLf & vbCrLf
     
     s = s + s2
   Next i
   Set p = POLICY_getFactory().getNewPageInstance
   p.pageName = "#history " + pageName
   p.raw = s
   Call p.cook(prep, chef, myBacklinks)
   Set ControllableModel_makeHistoryPage = p
End Function

Private Function ControllableModel_makeSearchResultsPage(searchTerm As String) As page
   Dim s As String
   s = "== Search Results ==" + vbCrLf + "Your search for ''" + searchTerm + "''" + vbCrLf
   
   Dim ps As PageSet
   Set ps = myStore.getPageSetContaining(searchTerm)
   s = s + " produced " + CStr(ps.size()) + " results " + vbCrLf + vbCrLf
   s = s + "BOX<" + vbCrLf
   s = s + ps.toWikiMarkup
   s = s + ">BOX" + vbCrLf
   
   Dim p As page
   Set p = POLICY_getFactory().getNewPageInstance()
   
   p.pageName = "#find " + searchTerm
   p.raw = s
   Call p.cook(prep, chef, False)
   Set ControllableModel_makeSearchResultsPage = p
End Function

Private Function ControllableModel_newNetworkPage() As String
  Dim p As page
  Set p = POLICY_getFactory().getNewPageInstance
  p.createdDate = Date
  p.raw = "#Network,, 1" + vbCrLf + "----" + vbCrLf + "----"
  
  Call p.cook(prep, chef, myBacklinks)
  Call setCurrentPage(p)
  myCurrentPageName = ""
  ControllableModel_newNetworkPage = "network"
End Function

Private Function ControllableModel_newPage() As String
  Dim p As page
  Set p = POLICY_getFactory().getNewPageInstance
  p.createdDate = Date
  Call setCurrentPage(p)
  myCurrentPageName = ""
  ControllableModel_newPage = "new page"
End Function

Private Sub ControllableModel_savePage()
    Call savePage
End Sub

Private Sub ControllableModel_setCurrentPage(p As page)
    Set myCurrentPage = p
End Sub

Private Function ControllableModel_wordCount(pageName As String) As Integer
    Dim p As page
    Set p = myStore.loadRaw(pageName)
    ControllableModel_wordCount = p.wordCount()
    Set p = Nothing
End Function

Private Property Set CrawlerSubsystem_crawlerManager(ByVal RHS As CrawlerDefinitionTable)
    Set myCrawlerMan = RHS
End Property

Private Property Get CrawlerSubsystem_crawlerManager() As CrawlerDefinitionTable
    Set CrawlerSubsystem_crawlerManager = myCrawlerMan
End Property

Private Function CrawlerSubsystem_makeCrawlersPage() As page
    Dim s As String, s2 As String
    Dim v As Variant, pc As PageCrawler
    Dim p As page
    Set p = POLICY_getFactory().getNewPageInstance
     
    s = "== Currently defined Crawlers ==" & vbCrLf
    s = s & "<table border='0' bgcolor='#ffeffe'>" + vbCrLf
    s = s & "Change CrawlerDefinitions <p />" & vbCrLf
    For Each v In myCrawlerMan.crawlerNames.toCollection
        Set pc = myCrawlerMan.crawlers.Item(CStr(v))
        s = s + "<tr><th valign='top'>"
        s = s & st.strip(CStr(v)) & "</th><td valign='center'>" & _
        vbCrLf & "##Button #showCrawler " & pc.name & ",, " & "Show" & _
        vbCrLf & "</td><td valign='center'>" & vbCrLf & _
        "##Button #crawl " & pc.name & ",, Example" & vbCrLf & _
        "</td></tr>"
    Next v
    s = s & "</table>" & vbCrLf
   
    p.pageName = "#crawlers"
    p.raw = s
    Call p.cook(prep, chef, False)
    
    Set pc = Nothing
    Set CrawlerSubsystem_makeCrawlersPage = p
End Function

Private Function CrawlerSubsystem_makeCrawlResultsPage(crawlerName As String, startPage As String) As page
   Dim s As String
   s = "== Crawler Results ==" + vbCrLf + "Your crawler ''" + crawlerName + "'' gathered : " + vbCrLf
   
   Dim ps As PageSet, crawl As PageCrawler
   
   Set crawl = myCrawlerMan.getCrawler(crawlerName)
   If Not crawl Is Nothing Then
     Call crawl.crawl(startPage)
     Set ps = crawl.getPages
   
     s = s + " these " + CStr(ps.size()) + " pages " + vbCrLf + vbCrLf
     s = s + "BOX<" + vbCrLf
     s = s + ps.toWikiMarkup
     s = s + ">BOX" + vbCrLf + vbCrLf
     s = s + "CrawlerDefinitions" + vbCrLf + vbCrLf
   Else
     s = "<font color=#990000>Error : probably couldn't find a crawler called " + crawlerName
     s = s + "</font> (Try CrawlerDefinitions)"
   
   End If
   Dim p As page
   Set p = POLICY_getFactory().getNewPageInstance
   p.pageName = "#crawl " + crawlerName + " " + startPage
   p.raw = s
   Call p.cook(prep, chef, False)
   
   Set CrawlerSubsystem_makeCrawlResultsPage = p
End Function

Private Sub ExportSubsystem_doExport(name As String)
    Call myExportMan.callExport(name, myStore.getPageStoreIdentifier)
End Sub

Private Sub ExportSubsystem_doInstantExport(exporterName As String, pageName As String)
    Call myExportMan.callInstantExport(exporterName, Me.asModelLevel.getWikiAnnotatedDataStore.store.getPageStoreIdentifier, pageName)
End Sub

Private Function ExportSubsystem_makeChooseExporterPage(pageName As String) As page
    Dim s As String, cnd As String
    Dim p As page
    Dim v As Variant
    
    Set p = POLICY_getFactory.getNewPageInstance
    
    p.pageName = "#chooseExporter pageName"
    s = "== Choose an Exporter ==" & vbCrLf
    s = s & "You want to export the page " & pageName & " but I need to know which of these currently known exporters you want to export with." & vbCrLf & vbCrLf
    s = s & "<table border='0' cellspacing='2' cellpadding='4' bgcolor='#edfdee'>" & _
    vbCrLf & "<tr><th>Exporter</th><th></th>"
     
    For Each v In myExportMan.exportPrograms.toCollection
        cnd = "#exportOne " & v & " " & pageName
        s = s & "<tr><th valign='top'>" & v & "</td>" & _
        "<td valign='top'>" & vbCrLf & "##Button " & _
        cnd & ",, " & "Do it" & vbCrLf & "</td>" & _
        "</tr>" & vbCrLf
    Next v
    s = s & "</table>" & vbCrLf
    s = s & vbCrLf & "Or do you really want to run one of the [[#exports|predefined exports]]? "
    p.raw = s
    Call p.cook(prep, chef, False)
    
    Set ExportSubsystem_makeChooseExporterPage = p
    
End Function

Private Function ExportSubsystem_makeExportersPage() As page
    Dim s As String
    Dim p As page
    Set p = POLICY_getFactory.getNewPageInstance
    p.pageName = "#exporters"
    s = "== Exporters I know about ==" & vbCrLf
    
    s = s & "BOX<" & vbCrLf & "#NoWiki" & vbCrLf
    s = s & myExportMan.exportersToString & vbCrLf
    s = s & "#Wiki" & vbCrLf & ">BOX" & vbCrLf
    s = s & "<p>Exporters are separate programs which can read pages from SdiDesk and export them in other file-formats. " & _
    vbCrLf & "If you want to install a new Exporter, simply drop the program into your ''exporters'' directory : " & _
    "<br /><br />" & Me.asModelLevel.getLocalFileSystem.getExporterDirectory & "<br /><br /> and restart SdiDesk</p>"
    p.raw = s
    Call p.cook(prep, chef, False)
    
    Set ExportSubsystem_makeExportersPage = p
End Function

Private Function ExportSubsystem_makeExportsPage() As page
    Dim s As String, er As ExportRecord
    Dim p As page
    Set p = POLICY_getFactory.getNewPageInstance
    
    p.pageName = "#exports"
    s = "== Currently defined Exports ==" & vbCrLf
    s = s & "Change ExportDefinitions <p />" & vbCrLf
    s = s & "<table border='1' cellspacing='2' cellpadding='4' bgcolor='#effeff'>" & _
    vbCrLf & "<tr><th>Name</th><th></th><th>Program</th><th>Defining Parameters</th>"
    
    For Each er In myExportMan.exportTable.toCollection
        s = s & "<tr><th valign='top'>" & er.name & "</td>" & _
        "<td valign='top'>" & vbCrLf & "##Button " & "#export " & er.name & ",, " & _
        "Do it" & vbCrLf & "</td>"
        s = s & "<td valign='top'>" & er.program & "</td>" & _
        "<td valign='top'>" & er.paramPage & "</td>" & _
        "</tr>" & vbCrLf
    Next er
    s = s & "</table>" & vbCrLf
   
    p.raw = s
    Call p.cook(prep, chef, False)
    
    Set ExportSubsystem_makeExportsPage = p
End Function



Private Property Let ExportSubsystem_pageStoreIdentifier(ByVal RHS As String)
    myPageStoreIdentifier = RHS
End Property

Private Property Get ExportSubsystem_pageStoreIdentifier() As String
    ExportSubsystem_pageStoreIdentifier = myPageStoreIdentifier
End Property

Private Sub ExportSubsystem_refreshExportManager(wads As WikiAnnotatedDataStore)
    Set myExportMan = New ExportManager
    Call myExportMan.scanForExporters(Me.asModelLevel.getLocalFileSystem)
    Call myExportMan.parseFromRawString(wads.getRawPageData("ExportDefinitions"))
End Sub


Private Sub ExportSubsystem_scanForExports()
    Call myExportMan.scanForExporters(Me)
End Sub

Private Sub LocalFileSystem_changeDirectory(path As String)
  Dim fs As New FileSystemPageStore
  
  Dim fName As String
  fName = path
    
  ' seems we can only get a path + fileName from the chooser,
  ' but don't need the filename so strip it
   fName = fs.pathFromFileName(fName)
   fName = fs.ensureTrailingSlash(fName)
   
   ' ensure the directory
   Call fs.ensureFullNameDirectory(fName)
    
   ' we are going to copy essential pages, so let's
   ' have them in a PageSet
   Dim ps As New PageSet
   Call ps.init
   Call ps.addPageFromName("AllPages", Me.myStore)
   Call ps.addPageFromName("RecentChanges", Me.myStore)
   
   ' now change to use the new directory
   Set Me.myStore = fs
   fs.setDataDirectory (fName)
      
   ' and save the essential pages there
   Call ps.saveAll(Me.myStore)
   
   ' call back
   Call mainForm.controller.actionLoad("StartPage", True)
  
   ' wipe NavigationHistory because it can confuse the user
   ' (So at the moment we can't "back" and "forward" through
   ' different directories. That *seems* right to me, because it
   ' keeps the directories really separate.
   ' OTOH, users may want them more closely integrated)
   
   Call myHistory.setComboBox(Me.mainForm.HistoryList)
End Sub


Private Function LocalFileSystem_getDirectoryListingAsVCollection(d As String) As VCollection
    Dim fs As FileSystemPageStore
    Set fs = myStore
    Set LocalFileSystem_getDirectoryListingAsVCollection = fs.dirAsVCollection(d)
End Function

Private Function LocalFileSystem_getExporterDirectory() As String
    Dim fs As FileSystemPageStore
    Set fs = myStore
    LocalFileSystem_getExporterDirectory = fs.exporterDirectory
End Function

Private Function LocalFileSystem_getMainDataDirectory() As String
    Dim fs As FileSystemPageStore
    Set fs = myStore
    LocalFileSystem_getMainDataDirectory = fs.mainDataDirectory
End Function

Private Function LocalFileSystem_hasLocalFileSystem() As Boolean
    LocalFileSystem_hasLocalFileSystem = True
End Function

Private Function LocalFileSystem_makeDirectoryPage(path As String) As page
   Dim fs As New FileSystemPageStore
   
   Dim s As String
   s = "=== Listing of " + path + " ===" + vbCrLf
   s = s + "BOX<" + vbCrLf
   s = s + fs.dirAsPage(path, mainForm.DirListBox)
   s = s + ">BOX" + vbCrLf + vbCrLf

   Dim p As page
   Set p = POLICY_getFactory().getNewPageInstance
   p.pageName = "#dir " + path
   p.raw = s
   Call p.cook(prep, chef, False)
   
   Set fs = Nothing
   Set LocalFileSystem_makeDirectoryPage = p
End Function

Private Function ModelLevel_getControllableModel() As ControllableModel
    Set ModelLevel_getControllableModel = Me
End Function

Private Function ModelLevel_getCrawlerSubsystem() As CrawlerSubsystem
    Set ModelLevel_getCrawlerSubsystem = Me
End Function

Private Function ModelLevel_getExportSubsystem() As ExportSubsystem
    Set ModelLevel_getExportSubsystem = Me
End Function

Private Function ModelLevel_getLocalFileSystem() As LocalFileSystem
    Set ModelLevel_getLocalFileSystem = Me
End Function

Private Function ModelLevel_getSingleUserState() As SingleUserState
    Set ModelLevel_getSingleUserState = Me
End Function

Private Function ModelLevel_getSystemConfigurations() As SystemConfigurations
    Set ModelLevel_getSystemConfigurations = Me
End Function

Private Function ModelLevel_getWikiAnnotatedDataStore() As WikiAnnotatedDataStore
    Set ModelLevel_getWikiAnnotatedDataStore = Me
End Function

Private Sub ModelLevel_setCallBackForm(f As Form)
    Set mainForm = f
    Dim sus As SingleUserState
    Set sus = ModelLevel_getSingleUserState()
    sus.history.setComboBox (f.HistoryList)
    Set sus = Nothing
End Sub

Private Sub ModelLevel_setForm(f As Form)
    Set mainForm = f
End Sub

Private Sub ModelLevel_setPageCooker(pc As PageCooker)
    Set chef = pc
End Sub


Private Sub ModelLevel_setPagePreparer(pp As PagePreparer)
    Set prep = pp
End Sub

Private Property Let SingleUserState_backlinks(ByVal RHS As Boolean)
    myBacklinks = RHS
End Property

Private Property Get SingleUserState_backlinks() As Boolean
    SingleUserState_backlinks = myBacklinks
End Property

Private Property Let SingleUserState_changesSaved(ByVal RHS As Boolean)
    myChangesSaved = RHS
End Property

Private Property Get SingleUserState_changesSaved() As Boolean
    SingleUserState_changesSaved = myChangesSaved
End Property

Private Property Set SingleUserState_currentPage(ByVal RHS As page)
    Set myCurrentPage = RHS
End Property

Private Property Get SingleUserState_currentPage() As page
    Set SingleUserState_currentPage = myCurrentPage
End Property

Private Property Let SingleUserState_currentPageName(ByVal RHS As String)
    myCurrentPageName = RHS
End Property

Private Property Get SingleUserState_currentPageName() As String
    SingleUserState_currentPageName = myCurrentPageName
End Property

Private Property Let SingleUserState_editState(ByVal RHS As PageEditState)
    myPageEditState = RHS
End Property

Private Property Get SingleUserState_editState() As PageEditState
    SingleUserState_editState = myPageEditState
End Property

Private Property Set SingleUserState_history(ByVal RHS As NavigationHistory)
    Set myHistory = RHS
End Property

Private Property Get SingleUserState_history() As NavigationHistory
    Set SingleUserState_history = myHistory
End Property

Private Property Let SingleUserState_isLoading(ByVal RHS As Boolean)
    myIsLoading = RHS
End Property

Private Property Get SingleUserState_isLoading() As Boolean
    SingleUserState_isLoading = myIsLoading
End Property



Private Property Let SingleUserState_oldPageName(ByVal RHS As String)
    myOldPageName = RHS
End Property

Private Property Get SingleUserState_oldPageName() As String
    SingleUserState_oldPageName = myOldPageName
End Property

Private Property Let SystemConfigurations_allPage(ByVal RHS As String)
    myAllPage = RHS
End Property

Private Property Get SystemConfigurations_allPage() As String
    SystemConfigurations_allPage = myAllPage
End Property

Private Property Let SystemConfigurations_configPage(ByVal RHS As String)
    myConfigPage = RHS
End Property

Private Property Get SystemConfigurations_configPage() As String
    SystemConfigurations_configPage = myConfigPage
End Property

Private Function SystemConfigurations_getTypeColour(typeName As String) As String
    SystemConfigurations_getTypeColour = linkTypeMan.getColour(typeName)
End Function

Private Property Let SystemConfigurations_helpIndexPage(ByVal RHS As String)
    myHelpIndexPage = RHS
End Property

Private Property Get SystemConfigurations_helpIndexPage() As String
    SystemConfigurations_helpIndexPage = myHelpIndexPage
End Property

Private Property Set SystemConfigurations_interMap(ByVal RHS As interWikiMap)
    Set myInterWikiMap = RHS
End Property

Private Property Get SystemConfigurations_interMap() As interWikiMap
    Set SystemConfigurations_interMap = myInterWikiMap
End Property

Private Property Let SystemConfigurations_recentChangesPage(ByVal RHS As String)
    myRecentChangesPage = RHS
End Property

Private Property Get SystemConfigurations_recentChangesPage() As String
    SystemConfigurations_recentChangesPage = myRecentChangesPage
End Property

Private Sub SystemConfigurations_setLinkTypeManager(l As LinkTypeManager)
    Set linkTypeMan = l
End Sub

Private Property Let SystemConfigurations_startPage(ByVal RHS As String)
    myStartPage = RHS
End Property

Private Property Get SystemConfigurations_startPage() As String
    SystemConfigurations_startPage = myStartPage
End Property

Private Function WikiAnnotatedDataStore_getPageSetContaining(s As String) As PageSet
    Set WikiAnnotatedDataStore_getPageSetContaining = myStore.getPageSetContaining(s)
End Function


Private Function WikiAnnotatedDataStore_getPageVar(pageName As String, varName As String) As String
    Dim p As page
    Set p = myStore.loadUntilNotRedirectRaw(pageName)
    
    Call p.cook(Me.prep, Me.chef, False)
    WikiAnnotatedDataStore_getPageVar = p.getVal(varName)
    Set p = Nothing
End Function

Private Function WikiAnnotatedDataStore_getRawPageData(pageName As String) As String
    Dim p As page
    Set p = myStore.loadRaw(pageName)
    WikiAnnotatedDataStore_getRawPageData = p.raw
    Set p = Nothing
End Function

Private Function WikiAnnotatedDataStore_pageExists(pName As String) As Boolean
    WikiAnnotatedDataStore_pageExists = myStore.pageExists(pName)
End Function


Private Property Set WikiAnnotatedDataStore_store(ByVal RHS As PageStore)
    Set myStore = RHS
End Property

Private Property Get WikiAnnotatedDataStore_store() As PageStore
    Set WikiAnnotatedDataStore_store = myStore
End Property
