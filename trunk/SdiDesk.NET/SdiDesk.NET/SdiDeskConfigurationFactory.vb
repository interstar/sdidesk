Option Strict Off
Option Explicit On
Friend Class SdiDeskConfigurationFactory
	
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
	
	Public Function getNativePageCooker() As _PageCooker
		getNativePageCooker = myNativePageCooker
	End Function
	
	Public Function getPagePreparer() As PagePreparer
		getPagePreparer = myPagePreparer
	End Function
	
	Public Function getNewPageInstance() As _Page
		Dim p As New MemoryResidentPage
		getNewPageInstance = p
	End Function
	
	Public Function getNewPageCrawlerInstance(ByRef crawlerType As String, ByRef name As String, ByRef maxDepth As Short, ByRef exPag As String, ByRef exTyp As String) As _PageCrawler
		Dim base As _PageCrawler
		Dim c As New RecursivePageCrawler
		Dim c2 As New TimeBasedPageCrawler
		Dim c3 As New AllPagesPageCrawler
		Select Case crawlerType
			Case "recursive"
				Call c.init(name, maxDepth, exPag, exTyp)
				base = c
			Case "recent"
				Call c2.init(name, exPag, exTyp)
				base = c2
			Case "all"
				Call c3.init(name)
				base = c3
			Case Else
				MsgBox("SdiDeskConfigurationFactory didn't recognise a crawler type called '" & crawlerType & "'")
		End Select
		
		base.wads = myModelLevel
		getNewPageCrawlerInstance = base
		'UPGRADE_NOTE: Object c may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		c = Nothing
		'UPGRADE_NOTE: Object c2 may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		c2 = Nothing
		'UPGRADE_NOTE: Object c3 may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		c3 = Nothing
	End Function
	
	Public Function getModelLevel() As _ModelLevel
		getModelLevel = myModelLevel
	End Function
	
	Public Function getNativeLinkWrapper() As _LinkWrapper
		getNativeLinkWrapper = myNativeLinkWrapper
	End Function
	
	Public Function getStandardLinkProcessor() As _LinkProcessor
		getStandardLinkProcessor = myStandardLinkProcessor
	End Function
	
	
	Public Function getNewPageStore(ByRef psi As String) As _PageStore
		myPageStore = New FileSystemPageStore
		Call myPageStore.setDataDirectory(psi)
		getNewPageStore = myPageStore
	End Function
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		myPagePreparer = New PagePreparer
		myNativePageCooker = New NativePageCooker
		myPageStore = New FileSystemPageStore
		' default data directory for myPageStore is App.path
		' if we want to change that, uncomment the next line and change
		' argument to the desired path
		' Call myPageStore.setDataDirectory(altPath)
		
		' default pages
		Dim spm As New StandardPagesManager
		'UPGRADE_WARNING: Couldn't resolve default property of object myPageStore. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call spm.ensureStandardPages(myPageStore)
		'UPGRADE_NOTE: Object spm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		spm = Nothing
		
		' make sure the model has the page-store
		' doing this via the WADS interface of the model
		myModelLevel = New ModelImplementation
		Dim model As _ModelLevel
		Dim localWads As _WikiAnnotatedDataStore
		model = myModelLevel
		localWads = model.getWikiAnnotatedDataStore
		localWads.store = myPageStore
		
		' set up the model's page-preparer and native-page-cooker (chef)
		' and their backlinks to the wads.
		Call model.setPagePreparer(myPagePreparer)
		'UPGRADE_WARNING: Couldn't resolve default property of object myNativePageCooker. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call model.setPageCooker(myNativePageCooker)
		
		myPagePreparer.wads = myModelLevel
		
		Dim sysConf As _SystemConfigurations
		sysConf = Me.getModelLevel.getSystemConfigurations
		
		sysConf.configPage = "ConfigPage"
		sysConf.startPage = "StartPage"
		sysConf.helpIndexPage = "HelpIndex"
		sysConf.allPage = "AllPages"
		sysConf.recentChangesPage = "RecentChanges"
		
		Call linkDefs()
		Call crawlerDefs()
		Call interWikiMap()
		
		Call nativeLinkWrapperAndStandardLinkProcessor()
		
		myNativeLinkWrapper.asLinkWrapper.remoteInterMap = Me.getModelLevel.getSystemConfigurations.interMap
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	Public Sub nativeLinkWrapperAndStandardLinkProcessor()
		myNativeLinkWrapper = New NativeLinkWrapper
		Me.getNativeLinkWrapper.remoteSysConf = Me.getModelLevel.getSystemConfigurations
		Me.getNativeLinkWrapper.remoteWads = Me.getModelLevel.getWikiAnnotatedDataStore
		
		myStandardLinkProcessor = New StandardLinkProcessor
		myNativePageCooker.myLinkProcessor = myStandardLinkProcessor
		myNativePageCooker.myLinkWrapper = myNativeLinkWrapper
		
	End Sub
	
	Private Sub linkDefs()
		' load the link definitions from page and set up the table
		Dim linkDefs As String
		linkDefs = Me.getModelLevel.getWikiAnnotatedDataStore.getRawPageData("LinkTypeDefinitions")
		
		myLinkTypeManager = New LinkTypeManager
		Call myLinkTypeManager.setupLinkTypes(linkDefs)
		
		' give it to the model
		Dim localSysConf As _SystemConfigurations
		localSysConf = Me.getModelLevel.getSystemConfigurations
		
		Call localSysConf.setLinkTypeManager(myLinkTypeManager)
	End Sub
	
	Private Sub interWikiMap()
		Dim interMap As String
		interMap = Me.getModelLevel.getWikiAnnotatedDataStore.getRawPageData("InterMap")
		Dim lines() As String
		Dim parts() As String
		Dim v As Object
		lines = Split(interMap, vbCrLf)
		For	Each v In lines
			'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
		
		myCrawlerDefinitionTable = New CrawlerDefinitionTable
		'UPGRADE_WARNING: Couldn't resolve default property of object myPageStore. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call myCrawlerDefinitionTable.parseFromTableString(crawlDefs, Me.getModelLevel, myPageStore, myNativePageCooker)
		
		' give it to model
		Dim localCrawlerSubsystem As _CrawlerSubsystem
		localCrawlerSubsystem = Me.getModelLevel.getCrawlerSubsystem
		
		localCrawlerSubsystem.crawlerManager = myCrawlerDefinitionTable
		
	End Sub
	
	Public Function wrapPageInNetwork(ByRef p As _Page) As Network
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Dim n As New Network
		If TypeName(p) = "Network" Then
			wrapPageInNetwork = p
		Else
			n.innerPage = p
			Call n.init(1, 200, 0.75)
			n.parseFromPrettyPersist((p.raw))
			p.cooked = "A network"
			p.pageType = "network"
			wrapPageInNetwork = n
		End If
	End Function
End Class