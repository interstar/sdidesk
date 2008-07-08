Option Strict Off
Option Explicit On
Friend Class ModelImplementation
	Implements _ModelLevel
	Implements _WikiAnnotatedDataStore
	Implements _SingleUserState
	Implements _SystemConfigurations
	Implements _CrawlerSubsystem
	Implements _ExportSubsystem
	Implements _ControllableModel
	Implements _LocalFileSystem
	
	' The "Model" part of an MVC conception of SdiDesk
	
	' It has (almost) no responsibility for call-back to rest of system
	' or the ''dynamics'' of control as the user flows through the
	' system
	
	' now (as of March 2005) the model level implements various
	' interfaces
	
	' Currently
	
	
	
	
    Public mainForm As WADSMainForm ' System.Windows.Forms.Form ' this is the WADSMainForm which calls it
	' and is needed for the occasional callback
	
	Public prep As PagePreparer ' what transforms raw into prepared
	Public chef As _PageCooker ' what transforms prepared into cooked
	
	' support for WikiAnnotatedDataStore interface
	Public myStore As _PageStore ' stores the pages
	
	' support for SingleUserState interface
	Private myBacklinks As Boolean
	Private myCurrentPageName As String
	Private myOldPageName As String
    Private myPageEditState As PageEditState
	Private myIsLoading As Boolean
	Private myHistory As NavigationHistory
	Private myChangesSaved As Boolean
	Private myCurrentPage As _page
	
	' support from SystemConfiguration interface
	Private linkTypeMan As LinkTypeManager ' stores info. about link types and colours
	Private myConfigPage As String
	Private myStartPage As String
	Private myHelpIndexPage As String
	Private myAllPage As String
	Private myRecentChangesPage As String
	
	Private myInterWikiMap As InterWikiMap
	
	' support for CrawlerSubsystem interface
	Private myCrawlerMan As CrawlerDefinitionTable ' creates and stores the crawlers
	
	' support for ExportSubsystem interface
	Private myExportMan As ExportManager ' stores info. about export scripts
	Private myPageStoreIdentifier As String ' where is the PageStore?
	
	' useful stuff
	Private st As StringTool
	
	Public Sub loadConfigs()
		Dim confPage As _page
		'UPGRADE_WARNING: Couldn't resolve default property of object confPage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		confPage = WikiAnnotatedDataStore_getRawPageData("RootConfig")
	End Sub
	
	
	Public Function qq(ByRef s As String) As String
		qq = Chr(34) & s & Chr(34)
	End Function
	
	
	Public Function asModelLevel() As _ModelLevel
		asModelLevel = Me
	End Function
	
	
	
	
	
	Public Function loadPage(ByRef pageName As String) As String
		' This is a lower level function than loadNewPage,
		' backPage and forwardPage
		' It just manages the loading
		' Use loadNewPage and backPage and forwardPage
		' when you want the package deal
		
		myCurrentPage = Me.myStore.loadRaw(pageName)
		Call myCurrentPage.cook(prep, chef, myBacklinks)
		myCurrentPageName = myCurrentPage.pageName
		loadPage = "normal"
		If myCurrentPage.isNetwork Then loadPage = "network"
		If myCurrentPage.isTable Then loadPage = "table"
		If myCurrentPage.isRedirect Then loadPage = "redirect"
		If myCurrentPage.isNew Then loadPage = "new page"
	End Function
	
	Public Function pageExists(ByRef pageName As String) As Boolean
		pageExists = myStore.pageExists(pageName)
	End Function
	
	Public Function pageContains(ByRef pageName As String, ByRef s As String) As Boolean
		Dim r As String
		r = WikiAnnotatedDataStore_getRawPageData(pageName)
		If InStr(r, s) > 0 Then
			pageContains = True
		Else
			pageContains = False
		End If
	End Function
	
	
	
	Public Function loadRawPage(ByRef pageName As String) As String
		myCurrentPage = myStore.loadRaw(pageName)
		myCurrentPageName = myCurrentPage.pageName
		loadRawPage = myCurrentPage.getMyType
		
		'   Call loadPage(pageName)
		
	End Function
	
	
	
	
	
	
	Public Sub savePage()
		Call myStore.savePage(myCurrentPage)
		
		' now special behaviours
		If myCurrentPage.pageName = "LinkTypeDefinitions" Then
			' we've saved new link types, so reload them
			Call linkTypeMan.setupLinkTypes((myCurrentPage.raw))
		End If
		
		If myCurrentPage.pageName = "CrawlerDefinitions" Then
			' we've saved new crawler defs, so reload them
			'UPGRADE_WARNING: Couldn't resolve default property of object Me. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call myCrawlerMan.parseFromTableString((myCurrentPage.raw), Me, myStore, chef)
		End If
		
		If myCurrentPage.pageName = "ExportDefinitions" Then
			' we've saved new export defs, so reload them
			Call myExportMan.parseFromRawString((myCurrentPage.raw))
		End If
		
	End Sub
	
	Public Function getCurrentPage() As _page
		getCurrentPage = myCurrentPage
	End Function
	
	Public Sub setCurrentPage(ByRef p As _page)
		myCurrentPage = p
	End Sub
	
	
	
	
	Public Function find(ByRef searchString As String) As String
		Dim ps As PageSet
		ps = myStore.getPageSetContaining(searchString)
		find = ps.toWikiMarkup
	End Function
	
	
	
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		st = New StringTool
		myHistory = New NavigationHistory
		myExportMan = New ExportManager
		myInterWikiMap = New InterWikiMap
		myBacklinks = False
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object st may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		st = Nothing
		'UPGRADE_NOTE: Object myExportMan may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		myExportMan = Nothing
		'UPGRADE_NOTE: Object myInterWikiMap may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		myInterWikiMap = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Private Sub ControllableModel_deletePage(ByRef pageName As String) Implements _ControllableModel.deletePage
		Call myStore.deletePage(pageName)
	End Sub
	
	Private Function ControllableModel_getCrawlerSubsystem() As _CrawlerSubsystem Implements _ControllableModel.getCrawlerSubsystem
		ControllableModel_getCrawlerSubsystem = Me
	End Function
	
	Private Function ControllableModel_getCurrentPage() As _page Implements _ControllableModel.getCurrentPage
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Dim net As New Network
		If myCurrentPage.isNetwork And TypeName(myCurrentPage) <> "Network" Then
			net = POLICY_getFactory().wrapPageInNetwork(myCurrentPage)
			myCurrentPage = net
		End If
		ControllableModel_getCurrentPage = myCurrentPage
	End Function
	
	Private Function ControllableModel_getExportSubsystem() As _ExportSubsystem Implements _ControllableModel.getExportSubsystem
		ControllableModel_getExportSubsystem = Me
	End Function
	
	Private Function ControllableModel_getLocalFileSystem() As _LocalFileSystem Implements _ControllableModel.getLocalFileSystem
		ControllableModel_getLocalFileSystem = Me
	End Function
	
	Private Function ControllableModel_getPageCooker() As _PageCooker Implements _ControllableModel.getPageCooker
		ControllableModel_getPageCooker = Me.chef
	End Function
	
	Private Function ControllableModel_getPagePreparer() As PagePreparer Implements _ControllableModel.getPagePreparer
		ControllableModel_getPagePreparer = prep
	End Function
	
	Private Function ControllableModel_getSingleUserState() As _SingleUserState Implements _ControllableModel.getSingleUserState
		ControllableModel_getSingleUserState = Me
	End Function
	
	Private Function ControllableModel_getSystemConfigurations() As _SystemConfigurations Implements _ControllableModel.getSystemConfigurations
		ControllableModel_getSystemConfigurations = Me
	End Function
	
	Private Function ControllableModel_getWikiAnnotatedDataStore() As _WikiAnnotatedDataStore Implements _ControllableModel.getWikiAnnotatedDataStore
		ControllableModel_getWikiAnnotatedDataStore = Me
	End Function
	
	Private Function ControllableModel_loadNewPage(ByRef pageName As String) As String Implements _ControllableModel.loadNewPage
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
	
	Private Function ControllableModel_loadRawPage(ByRef pageName As String) As String Implements _ControllableModel.loadRawPage
		ControllableModel_loadRawPage = loadRawPage(pageName)
	End Function
	
	Private Function ControllableModel_makeHistoryPage(ByRef pageName As String) As _page Implements _ControllableModel.makeHistoryPage
		Dim f1 As String
		Dim s2, s, colHex As String
		Dim p, p1 As _page
		Dim i As Short
		s = "== Current ==" & vbCrLf
		s = s & WikiAnnotatedDataStore_getRawPageData(pageName) & vbCrLf & "----" & vbCrLf
		For i = 1 To 5
			p1 = myStore.loadOldPage(pageName, i)
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
			
			s = s & s2
		Next i
		p = POLICY_getFactory().getNewPageInstance
		p.pageName = "#history " & pageName
		p.raw = s
		Call p.cook(prep, chef, myBacklinks)
		ControllableModel_makeHistoryPage = p
	End Function
	
	Private Function ControllableModel_makeSearchResultsPage(ByRef searchTerm As String) As _page Implements _ControllableModel.makeSearchResultsPage
		Dim s As String
		s = "== Search Results ==" & vbCrLf & "Your search for ''" & searchTerm & "''" & vbCrLf
		
		Dim ps As PageSet
		ps = myStore.getPageSetContaining(searchTerm)
		s = s & " produced " & CStr(ps.size()) & " results " & vbCrLf & vbCrLf
		s = s & "BOX<" & vbCrLf
		s = s & ps.toWikiMarkup
		s = s & ">BOX" & vbCrLf
		
		Dim p As _page
		p = POLICY_getFactory().getNewPageInstance()
		
		p.pageName = "#find " & searchTerm
		p.raw = s
		Call p.cook(prep, chef, False)
		ControllableModel_makeSearchResultsPage = p
	End Function
	
	Private Function ControllableModel_newNetworkPage() As String Implements _ControllableModel.newNetworkPage
		Dim p As _page
		p = POLICY_getFactory().getNewPageInstance
		p.createdDate = Today
		p.raw = "#Network,, 1" & vbCrLf & "----" & vbCrLf & "----"
		
		Call p.cook(prep, chef, myBacklinks)
		Call setCurrentPage(p)
		myCurrentPageName = ""
		ControllableModel_newNetworkPage = "network"
	End Function
	
	Private Function ControllableModel_newPage() As String Implements _ControllableModel.newPage
		Dim p As _page
		p = POLICY_getFactory().getNewPageInstance
		p.createdDate = Today
		Call setCurrentPage(p)
		myCurrentPageName = ""
		ControllableModel_newPage = "new page"
	End Function
	
	Private Sub ControllableModel_savePage() Implements _ControllableModel.savePage
		Call savePage()
	End Sub
	
	Private Sub ControllableModel_setCurrentPage(ByRef p As _page) Implements _ControllableModel.setCurrentPage
		myCurrentPage = p
	End Sub
	
	Private Function ControllableModel_wordCount(ByRef pageName As String) As Short Implements _ControllableModel.wordCount
		Dim p As _page
		p = myStore.loadRaw(pageName)
		ControllableModel_wordCount = p.wordCount()
		'UPGRADE_NOTE: Object p may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		p = Nothing
	End Function
	
	
	Private Property CrawlerSubsystem_crawlerManager() As CrawlerDefinitionTable Implements _CrawlerSubsystem.crawlerManager
		Get
			CrawlerSubsystem_crawlerManager = myCrawlerMan
		End Get
		Set(ByVal Value As CrawlerDefinitionTable)
			myCrawlerMan = Value
		End Set
	End Property
	
	
	
	
	Private Property ExportSubsystem_pageStoreIdentifier() As String Implements _ExportSubsystem.pageStoreIdentifier
		Get
			ExportSubsystem_pageStoreIdentifier = myPageStoreIdentifier
		End Get
		Set(ByVal Value As String)
			myPageStoreIdentifier = Value
		End Set
	End Property
	
	
	Private Property SingleUserState_backlinks() As Boolean Implements _SingleUserState.backlinks
		Get
			SingleUserState_backlinks = myBacklinks
		End Get
		Set(ByVal Value As Boolean)
			myBacklinks = Value
		End Set
	End Property
	
	
	Private Property SingleUserState_changesSaved() As Boolean Implements _SingleUserState.changesSaved
		Get
			SingleUserState_changesSaved = myChangesSaved
		End Get
		Set(ByVal Value As Boolean)
			myChangesSaved = Value
		End Set
	End Property
	
	
	Private Property SingleUserState_currentPage() As _page Implements _SingleUserState.currentPage
		Get
			SingleUserState_currentPage = myCurrentPage
		End Get
		Set(ByVal Value As _page)
			myCurrentPage = Value
		End Set
	End Property
	
	
	Private Property SingleUserState_currentPageName() As String Implements _SingleUserState.currentPageName
		Get
			SingleUserState_currentPageName = myCurrentPageName
		End Get
		Set(ByVal Value As String)
			myCurrentPageName = Value
		End Set
	End Property
	
	
    Private Property SingleUserState_editState() As PageEditState Implements _SingleUserState.editState
        Get
            SingleUserState_editState = myPageEditState
        End Get
        Set(ByVal Value As PageEditState)
            myPageEditState = Value
        End Set
    End Property
	
	
	Private Property SingleUserState_history() As NavigationHistory Implements _SingleUserState.history
		Get
			SingleUserState_history = myHistory
		End Get
		Set(ByVal Value As NavigationHistory)
			myHistory = Value
		End Set
	End Property
	
	
	Private Property SingleUserState_isLoading() As Boolean Implements _SingleUserState.isLoading
		Get
			SingleUserState_isLoading = myIsLoading
		End Get
		Set(ByVal Value As Boolean)
			myIsLoading = Value
		End Set
	End Property
	
	
	
	
	Private Property SingleUserState_oldPageName() As String Implements _SingleUserState.oldPageName
		Get
			SingleUserState_oldPageName = myOldPageName
		End Get
		Set(ByVal Value As String)
			myOldPageName = Value
		End Set
	End Property
	
	
	Private Property SystemConfigurations_allPage() As String Implements _SystemConfigurations.allPage
		Get
			SystemConfigurations_allPage = myAllPage
		End Get
		Set(ByVal Value As String)
			myAllPage = Value
		End Set
	End Property
	
	
	Private Property SystemConfigurations_configPage() As String Implements _SystemConfigurations.configPage
		Get
			SystemConfigurations_configPage = myConfigPage
		End Get
		Set(ByVal Value As String)
			myConfigPage = Value
		End Set
	End Property
	
	
	Private Property SystemConfigurations_helpIndexPage() As String Implements _SystemConfigurations.helpIndexPage
		Get
			SystemConfigurations_helpIndexPage = myHelpIndexPage
		End Get
		Set(ByVal Value As String)
			myHelpIndexPage = Value
		End Set
	End Property
	
	
	Private Property SystemConfigurations_interMap() As InterWikiMap Implements _SystemConfigurations.interMap
		Get
			SystemConfigurations_interMap = myInterWikiMap
		End Get
		Set(ByVal Value As InterWikiMap)
			myInterWikiMap = Value
		End Set
	End Property
	
	
	Private Property SystemConfigurations_recentChangesPage() As String Implements _SystemConfigurations.recentChangesPage
		Get
			SystemConfigurations_recentChangesPage = myRecentChangesPage
		End Get
		Set(ByVal Value As String)
			myRecentChangesPage = Value
		End Set
	End Property
	
	
	Private Property SystemConfigurations_startPage() As String Implements _SystemConfigurations.startPage
		Get
			SystemConfigurations_startPage = myStartPage
		End Get
		Set(ByVal Value As String)
			myStartPage = Value
		End Set
	End Property
	
	
	
	Private Property WikiAnnotatedDataStore_store() As _PageStore Implements _WikiAnnotatedDataStore.store
		Get
			WikiAnnotatedDataStore_store = myStore
		End Get
		Set(ByVal Value As _PageStore)
			myStore = Value
		End Set
	End Property
	
	Private Function CrawlerSubsystem_makeCrawlersPage() As _page Implements _CrawlerSubsystem.makeCrawlersPage
		Dim s, s2 As String
		Dim v As Object
		Dim pc As _PageCrawler
		Dim p As _page
		p = POLICY_getFactory().getNewPageInstance
		
		s = "== Currently defined Crawlers ==" & vbCrLf
		s = s & "<table border='0' bgcolor='#ffeffe'>" & vbCrLf
		s = s & "Change CrawlerDefinitions <p />" & vbCrLf
		For	Each v In myCrawlerMan.crawlerNames.toCollection
			'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pc = myCrawlerMan.crawlers.Item(CStr(v))
			s = s & "<tr><th valign='top'>"
			'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			s = s & st.strip(CStr(v)) & "</th><td valign='center'>" & vbCrLf & "##Button #showCrawler " & pc.name & ",, " & "Show" & vbCrLf & "</td><td valign='center'>" & vbCrLf & "##Button #crawl " & pc.name & ",, Example" & vbCrLf & "</td></tr>"
		Next v
		s = s & "</table>" & vbCrLf
		
		p.pageName = "#crawlers"
		p.raw = s
		Call p.cook(prep, chef, False)
		
		'UPGRADE_NOTE: Object pc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pc = Nothing
		CrawlerSubsystem_makeCrawlersPage = p
	End Function
	
	Private Function CrawlerSubsystem_makeCrawlResultsPage(ByRef crawlerName As String, ByRef startPage As String) As _page Implements _CrawlerSubsystem.makeCrawlResultsPage
		Dim s As String
		s = "== Crawler Results ==" & vbCrLf & "Your crawler ''" & crawlerName & "'' gathered : " & vbCrLf
		
		Dim ps As PageSet
		Dim crawl As _PageCrawler
		
		crawl = myCrawlerMan.getCrawler(crawlerName)
		If Not crawl Is Nothing Then
			Call crawl.crawl(startPage)
			ps = crawl.getPages
			
			s = s & " these " & CStr(ps.size()) & " pages " & vbCrLf & vbCrLf
			s = s & "BOX<" & vbCrLf
			s = s & ps.toWikiMarkup
			s = s & ">BOX" & vbCrLf & vbCrLf
			s = s & "CrawlerDefinitions" & vbCrLf & vbCrLf
		Else
			s = "<font color=#990000>Error : probably couldn't find a crawler called " & crawlerName
			s = s & "</font> (Try CrawlerDefinitions)"
			
		End If
		Dim p As _page
		p = POLICY_getFactory().getNewPageInstance
		p.pageName = "#crawl " & crawlerName & " " & startPage
		p.raw = s
		Call p.cook(prep, chef, False)
		
		CrawlerSubsystem_makeCrawlResultsPage = p
	End Function
	
	Private Sub ExportSubsystem_doExport(ByRef name As String) Implements _ExportSubsystem.doExport
		Call myExportMan.callExport(name, myStore.getPageStoreIdentifier)
	End Sub
	
	Private Sub ExportSubsystem_doInstantExport(ByRef exporterName As String, ByRef pageName As String) Implements _ExportSubsystem.doInstantExport
		Call myExportMan.callInstantExport(exporterName, Me.asModelLevel.getWikiAnnotatedDataStore.store.getPageStoreIdentifier, pageName)
	End Sub
	
	Private Function ExportSubsystem_makeChooseExporterPage(ByRef pageName As String) As _page Implements _ExportSubsystem.makeChooseExporterPage
		Dim s, cnd As String
		Dim p As _page
		Dim v As Object
		
		p = POLICY_getFactory.getNewPageInstance
		
		p.pageName = "#chooseExporter pageName"
		s = "== Choose an Exporter ==" & vbCrLf
		s = s & "You want to export the page " & pageName & " but I need to know which of these currently known exporters you want to export with." & vbCrLf & vbCrLf
		s = s & "<table border='0' cellspacing='2' cellpadding='4' bgcolor='#edfdee'>" & vbCrLf & "<tr><th>Exporter</th><th></th>"
		
		For	Each v In myExportMan.exportPrograms.toCollection
			'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			cnd = "#exportOne " & v & " " & pageName
			'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			s = s & "<tr><th valign='top'>" & v & "</td>" & "<td valign='top'>" & vbCrLf & "##Button " & cnd & ",, " & "Do it" & vbCrLf & "</td>" & "</tr>" & vbCrLf
		Next v
		s = s & "</table>" & vbCrLf
		s = s & vbCrLf & "Or do you really want to run one of the [[#exports|predefined exports]]? "
		p.raw = s
		Call p.cook(prep, chef, False)
		
		ExportSubsystem_makeChooseExporterPage = p
		
	End Function
	
	Private Function ExportSubsystem_makeExportersPage() As _page Implements _ExportSubsystem.makeExportersPage
		Dim s As String
		Dim p As _page
		p = POLICY_getFactory.getNewPageInstance
		p.pageName = "#exporters"
		s = "== Exporters I know about ==" & vbCrLf
		
		s = s & "BOX<" & vbCrLf & "#NoWiki" & vbCrLf
		s = s & myExportMan.exportersToString & vbCrLf
		s = s & "#Wiki" & vbCrLf & ">BOX" & vbCrLf
		s = s & "<p>Exporters are separate programs which can read pages from SdiDesk and export them in other file-formats. " & vbCrLf & "If you want to install a new Exporter, simply drop the program into your ''exporters'' directory : " & "<br /><br />" & Me.asModelLevel.getLocalFileSystem.getExporterDirectory & "<br /><br /> and restart SdiDesk</p>"
		p.raw = s
		Call p.cook(prep, chef, False)
		
		ExportSubsystem_makeExportersPage = p
	End Function
	
	Private Function ExportSubsystem_makeExportsPage() As _page Implements _ExportSubsystem.makeExportsPage
		Dim s As String
		Dim er As ExportRecord
		Dim p As _page
		p = POLICY_getFactory.getNewPageInstance
		
		p.pageName = "#exports"
		s = "== Currently defined Exports ==" & vbCrLf
		s = s & "Change ExportDefinitions <p />" & vbCrLf
		s = s & "<table border='1' cellspacing='2' cellpadding='4' bgcolor='#effeff'>" & vbCrLf & "<tr><th>Name</th><th></th><th>Program</th><th>Defining Parameters</th>"
		
		For	Each er In myExportMan.exportTable.toCollection
			s = s & "<tr><th valign='top'>" & er.name & "</td>" & "<td valign='top'>" & vbCrLf & "##Button " & "#export " & er.name & ",, " & "Do it" & vbCrLf & "</td>"
			'UPGRADE_WARNING: Couldn't resolve default property of object er.paramPage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object er.program. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			s = s & "<td valign='top'>" & er.program & "</td>" & "<td valign='top'>" & er.paramPage & "</td>" & "</tr>" & vbCrLf
		Next er
		s = s & "</table>" & vbCrLf
		
		p.raw = s
		Call p.cook(prep, chef, False)
		
		ExportSubsystem_makeExportsPage = p
	End Function
	
	Private Sub ExportSubsystem_refreshExportManager(ByRef wads As _WikiAnnotatedDataStore) Implements _ExportSubsystem.refreshExportManager
		myExportMan = New ExportManager
		Call myExportMan.scanForExporters(Me.asModelLevel.getLocalFileSystem)
		Call myExportMan.parseFromRawString(wads.getRawPageData("ExportDefinitions"))
	End Sub
	
	
	Private Sub ExportSubsystem_scanForExports() Implements _ExportSubsystem.scanForExports
		'UPGRADE_WARNING: Couldn't resolve default property of object Me. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call myExportMan.scanForExporters(Me)
	End Sub
	
	Private Sub LocalFileSystem_changeDirectory(ByRef path As String) Implements _LocalFileSystem.changeDirectory
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
		Call ps.init()
		Call ps.addPageFromName("AllPages", (Me.myStore))
		Call ps.addPageFromName("RecentChanges", (Me.myStore))
		
		' now change to use the new directory
		Me.myStore = fs
		fs.setDataDirectory((fName))
		
		' and save the essential pages there
		Call ps.saveAll((Me.myStore))
		
		' call back
		'UPGRADE_ISSUE: Control controller could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call mainForm.controller.actionLoad("StartPage", True)
		
		' wipe NavigationHistory because it can confuse the user
		' (So at the moment we can't "back" and "forward" through
		' different directories. That *seems* right to me, because it
		' keeps the directories really separate.
		' OTOH, users may want them more closely integrated)
		
		'UPGRADE_ISSUE: Control HistoryList could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call myHistory.setComboBox(Me.mainForm.HistoryList)
	End Sub
	
	
	Private Function LocalFileSystem_getDirectoryListingAsVCollection(ByRef d As String) As VCollection Implements _LocalFileSystem.getDirectoryListingAsVCollection
		Dim fs As FileSystemPageStore
		fs = myStore
		LocalFileSystem_getDirectoryListingAsVCollection = fs.dirAsVCollection(d)
	End Function
	
	Private Function LocalFileSystem_getExporterDirectory() As String Implements _LocalFileSystem.getExporterDirectory
		Dim fs As FileSystemPageStore
		fs = myStore
		LocalFileSystem_getExporterDirectory = fs.exporterDirectory
	End Function
	
	Private Function LocalFileSystem_getMainDataDirectory() As String Implements _LocalFileSystem.getMainDataDirectory
		Dim fs As FileSystemPageStore
		fs = myStore
		LocalFileSystem_getMainDataDirectory = fs.mainDataDirectory
	End Function
	
	Private Function LocalFileSystem_hasLocalFileSystem() As Boolean Implements _LocalFileSystem.hasLocalFileSystem
		LocalFileSystem_hasLocalFileSystem = True
	End Function
	
	Private Function LocalFileSystem_makeDirectoryPage(ByRef path As String) As _page Implements _LocalFileSystem.makeDirectoryPage
		Dim fs As New FileSystemPageStore
		
		Dim s As String
		s = "=== Listing of " & path & " ===" & vbCrLf
		s = s & "BOX<" & vbCrLf
		'UPGRADE_ISSUE: Control DirListBox could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		'UPGRADE_WARNING: Couldn't resolve default property of object fs.dirAsPage(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ''Zebo''		s = s + fs.dirAsPage(path, mainForm.DirListBox)
		s = s & ">BOX" & vbCrLf & vbCrLf
		
		Dim p As _page
		p = POLICY_getFactory().getNewPageInstance
		p.pageName = "#dir " & path
		p.raw = s
		Call p.cook(prep, chef, False)
		
		'UPGRADE_NOTE: Object fs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		fs = Nothing
		LocalFileSystem_makeDirectoryPage = p
	End Function
	
	Private Function ModelLevel_getControllableModel() As _ControllableModel Implements _ModelLevel.getControllableModel
		ModelLevel_getControllableModel = Me
	End Function
	
	Private Function ModelLevel_getCrawlerSubsystem() As _CrawlerSubsystem Implements _ModelLevel.getCrawlerSubsystem
		ModelLevel_getCrawlerSubsystem = Me
	End Function
	
	Private Function ModelLevel_getExportSubsystem() As _ExportSubsystem Implements _ModelLevel.getExportSubsystem
		ModelLevel_getExportSubsystem = Me
	End Function
	
	Private Function ModelLevel_getLocalFileSystem() As _LocalFileSystem Implements _ModelLevel.getLocalFileSystem
		ModelLevel_getLocalFileSystem = Me
	End Function
	
	Private Function ModelLevel_getSingleUserState() As _SingleUserState Implements _ModelLevel.getSingleUserState
		ModelLevel_getSingleUserState = Me
	End Function
	
	Private Function ModelLevel_getSystemConfigurations() As _SystemConfigurations Implements _ModelLevel.getSystemConfigurations
		ModelLevel_getSystemConfigurations = Me
	End Function
	
	Private Function ModelLevel_getWikiAnnotatedDataStore() As _WikiAnnotatedDataStore Implements _ModelLevel.getWikiAnnotatedDataStore
		ModelLevel_getWikiAnnotatedDataStore = Me
	End Function
	
	Private Sub ModelLevel_setCallBackForm(ByRef f As System.Windows.Forms.Form) Implements _ModelLevel.setCallBackForm
		mainForm = f
		Dim sus As _SingleUserState
		sus = ModelLevel_getSingleUserState()
		'UPGRADE_ISSUE: Control HistoryList could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        sus.history.setComboBox((mainForm.HistoryList))
		'UPGRADE_NOTE: Object sus may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		sus = Nothing
	End Sub
	
	Private Sub ModelLevel_setForm(ByRef f As System.Windows.Forms.Form) Implements _ModelLevel.setForm
		mainForm = f
	End Sub
	
	Private Sub ModelLevel_setPageCooker(ByRef pc As _PageCooker) Implements _ModelLevel.setPageCooker
		chef = pc
	End Sub
	
	
	Private Sub ModelLevel_setPagePreparer(ByRef pp As PagePreparer) Implements _ModelLevel.setPagePreparer
		prep = pp
	End Sub
	
	'UPGRADE_NOTE: typeName was upgraded to typeName_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Function SystemConfigurations_getTypeColour(ByRef typeName_Renamed As String) As String Implements _SystemConfigurations.getTypeColour
		SystemConfigurations_getTypeColour = linkTypeMan.getColour(typeName_Renamed)
	End Function
	
	Private Sub SystemConfigurations_setLinkTypeManager(ByRef l As LinkTypeManager) Implements _SystemConfigurations.setLinkTypeManager
		linkTypeMan = l
	End Sub
	
	Private Function WikiAnnotatedDataStore_getPageSetContaining(ByRef s As String) As PageSet Implements _WikiAnnotatedDataStore.getPageSetContaining
		WikiAnnotatedDataStore_getPageSetContaining = myStore.getPageSetContaining(s)
	End Function
	
	
	Private Function WikiAnnotatedDataStore_getPageVar(ByRef pageName As String, ByRef varName As String) As String Implements _WikiAnnotatedDataStore.getPageVar
		Dim p As _page
		p = myStore.loadUntilNotRedirectRaw(pageName)
		
		Call p.cook((Me.prep), (Me.chef), False)
		WikiAnnotatedDataStore_getPageVar = p.getVal(varName)
		'UPGRADE_NOTE: Object p may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		p = Nothing
	End Function
	
	Private Function WikiAnnotatedDataStore_getRawPageData(ByRef pageName As String) As String Implements _WikiAnnotatedDataStore.getRawPageData
		Dim p As _page
		p = myStore.loadRaw(pageName)
		WikiAnnotatedDataStore_getRawPageData = p.raw
		'UPGRADE_NOTE: Object p may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		p = Nothing
	End Function
	
	Private Function WikiAnnotatedDataStore_pageExists(ByRef pName As String) As Boolean Implements _WikiAnnotatedDataStore.pageExists
		WikiAnnotatedDataStore_pageExists = myStore.pageExists(pName)
	End Function
End Class