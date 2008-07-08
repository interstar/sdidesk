Option Strict Off
Option Explicit On
Friend Class ControlLevel
	
	' this is the umberella "Control" level
	' (as opposed to Model and View levels
	
	' The purpose of this object is to provide an API or
	' set of "Actions" which can be called on the Wiki.
	
	' These include :
	
	' View / Page Navigation Actions
	' =============
	
	' "Load" : load a new page and display the appropriate view
	' "Raw" : display the raw view of the page
	' "Preview" : disply (but not save) the cooked view of the page
	' "Edit" : invoke the custom editor of the page
	' "Save" : save the page and display the cooked version
	' "PageHistory" : display a view of a page's history
	' "Dir" : display a dynamically created page showing a directory on machine
	
	' "Back" : return to the previous state of the user history
	' "Forward" : go forward to the next state of the user history
	
	' Page Manipulation Actions
	' =========================
	
	' "New" : create a new blank page
	' "NewNet" : create a new, empty network
	' "Delete" : remove a page from the system
	' "Rename" : change the name of a page (to be implemented)
	
	' Page processing actions
	' =======================
	
	' "WordCount" : load a page and count the words in it.
	
	' Fixed Page Actions
	' ====================
	' "Start" : Go to the StartPage
	' "Help" : Go to the HelpIndex
	' "Config" : Go to the ConfigIndex
	' "All" : Go to the page which shows all pages
	' "Recent" : Go to the RecentChanges pages (to be implemented)
	' "About" : A page about SdiDesk and NooRanch
	
	' Export Actions
	' ================
	' "Export" : Run an export defined in ExportDefinitions
	
	' Search Actions
	' ================
	' "Find" : return a list of pages containing the search string
	' "Crawl" : return a list of pages from a named crawler
	
	' And later on : search and replace
	
	' These actions can be invoked a number of ways
	' * buttons on the form,
	' * commands in the NavBox,
	' * user history
	' etc
	
	' ==========================================
	' Instance vars ...
	' ===========================================
	
    Public mainForm As WADSMainForm ' System.Windows.Forms.Form
	Public model As _ControllableModel
	
	' ========================
	' Methods
	' ========================
	
	
	' == Outer Methods ==
	
	Public Sub init(ByRef f As System.Windows.Forms.Form, ByRef m As _ControllableModel)
		' f is the main form
		mainForm = f
		model = m
	End Sub
	
	
	Public Sub processCommand(ByRef s As String, ByRef fromHistory As Boolean)
		' fromHistory is true if this came from a history action
		
		'UPGRADE_NOTE: command was upgraded to command_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim pageName, command_Renamed, tail As String
		
		' split command into command and argument with a NavCommand
		Dim nc As New NavCommand
		nc.init((s))
		command_Renamed = nc.getCommand
		pageName = nc.getPageName
		tail = nc.tail
		
		If fromHistory = False Then
			' wipe futurity
			Call model.getSingleUserState.history.wipeFuture()
		End If
		
		' now decide what to do with it
		Select Case command_Renamed
			
			' View / Page Navigation Actions
			Case "#load"
				Call actionLoad(pageName, fromHistory)
				
			Case "#raw"
				Call actionRaw(pageName, fromHistory)
				
			Case "#preview"
				Call actionPreview(fromHistory)
				' NB : doesn't make sense to preview a pageName
				
			Case "#edit"
				Call actionEdit(pageName, fromHistory)
				
			Case "#save"
				Call actionSave(pageName, fromHistory)
				
			Case "#history"
				Call actionPageHistory(pageName, fromHistory)
				
			Case "#revert"
				Call actionRevert(tail, fromHistory)
				
			Case "#dir"
				Call actionDir(tail, fromHistory)
				
			Case "#back"
				Call actionBack()
				
			Case "#forward"
				Call actionForward()
				
				' Page Manipulation Actions
				
			Case "#new"
				Call actionNew(fromHistory)
				
			Case "#newNet"
				Call actionNewNetwork(fromHistory)
				
			Case "#delete"
				Call actionDelete(pageName)
				
				' Page processing
			Case "#wordCount"
				Call actionWordCount(pageName)
				
				' special pages
			Case "#start"
				Call actionStart(fromHistory)
				
			Case "#config"
				Call actionConfig(fromHistory)
				
			Case "#help"
				Call actionHelp(fromHistory)
				
			Case "#all"
				Call actionAll(fromHistory)
				
			Case "#recent"
				Call actionRecent(fromHistory)
				
			Case "#about"
				Call actionAbout(fromHistory)
				
				' export actions
				
			Case "#exporters" ' list of exporters (programs)
				Call actionExporters(fromHistory)
				
			Case "#exports" ' list of exports (preset, combination of program + pageset etc)
				Call actionExports(fromHistory)
				
			Case "#export" ' do an export
				Call actionExport(tail)
				
			Case "#exportOne" ' export a single page directly
				Call actionExportOne(tail)
				
			Case "#instantExport" ' makes a page of instantExport options
				Call actionInstantExport(pageName)
				
				' search actions
			Case "#find"
				Call actionFind(tail, fromHistory)
				
			Case "#crawl"
				Call actionCrawl(tail, fromHistory)
				
				' program state
			Case "#backLinksOn"
				Call actionBackLinksOn(fromHistory)
				
			Case "#backLinksOff"
				Call actionBackLinksOff(fromHistory)
				
				' other useful info
			Case "#showCrawler"
				Call actionShowCrawler(pageName, fromHistory)
				
			Case "#crawlers"
				Call actionCrawlers(fromHistory)
				
			Case "#shell"
				Call actionShell(tail, fromHistory)
				
			Case Else
				MsgBox("ControlLevel:processCommand. Sorry, don't know how to " & command_Renamed & " when given " & s)
				
		End Select
		
	End Sub
	
	' == Action methods ===========
	' View / Page Navigation Actions
	
	Public Sub actionLoad(ByRef pName As String, ByRef fromHistory As Boolean)
        Call saveGuard(PageEditState.LoadedState)
		'UPGRADE_ISSUE: Control vm could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call mainForm.vm.hideAll()
		If model.getWikiAnnotatedDataStore.pageExists(pName) Then
			model.loadNewPage(pName)
			showCooked()
		Else
			model.loadNewPage(pName)
			showRaw()
		End If
		
        model.getSingleUserState.editState = PageEditState.LoadedState
		If fromHistory = False Then
			model.getSingleUserState.history.append(("#load " & pName))
		End If
	End Sub
	
	Public Sub actionRaw(ByRef pName As String, ByRef fromHistory As Boolean)
        Call saveGuard(PageEditState.RawState)
		'UPGRADE_ISSUE: Control vm could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call mainForm.vm.hideAll()
		Call model.loadRawPage(pName)
		Call showRaw()
		If fromHistory = False Then
			model.getSingleUserState.history.append(("#raw " & pName))
		End If
	End Sub
	
	Public Sub actionPreview(ByRef fromHistory As Boolean)
        Call saveGuard(PageEditState.PreviewState)
		Dim p As _Page
		p = formToPage()
		Call p.cook(model.getPagePreparer, model.getPageCooker, model.getSingleUserState.backlinks)
		Call model.setCurrentPage(p)
		If p.isNetwork Then
			'UPGRADE_ISSUE: Control vse could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			mainForm.vse.setMode(VseCanvas.vseMode.View)
		End If
		Call showCooked()
		If fromHistory = False Then
			model.getSingleUserState.history.append(("#preview " & p.pageName))
		End If
	End Sub
	
	Public Sub actionEdit(ByRef pName As String, ByRef fromHistory As Boolean)
        Call saveGuard(PageEditState.EditedState)
		Call model.loadRawPage(pName)
		Call model.getCurrentPage().cook(model.getPagePreparer, model.getPageCooker, model.getSingleUserState.backlinks)
		Dim theTable As New Table
		If model.getCurrentPage().isNetwork Then
			'UPGRADE_ISSUE: Control vse could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			mainForm.vse.setMode(VseCanvas.vseMode.Edit)
			'UPGRADE_ISSUE: Control vm could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
            Call mainForm.vm.showVse()
            Application.DoEvents()

			'UPGRADE_ISSUE: Control vse could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			Call mainForm.vse.draw(model.getCurrentPage(), mainForm.vse.mode)
		Else
			If model.getCurrentPage.isTable Then
				theTable.parseFromDoubleCommaString(model.getCurrentPage.raw)
				'UPGRADE_ISSUE: Control td could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				Call mainForm.td.fillFromTable(theTable)
				'UPGRADE_ISSUE: Control vm could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				Call mainForm.vm.showTable()
				'UPGRADE_NOTE: Object theTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				theTable = Nothing
			Else
				Call showRaw()
			End If
		End If
		If fromHistory = False Then
			model.getSingleUserState.history.append(("#edit " & pName))
		End If
	End Sub
	
	Public Sub actionSave(ByRef pName As String, ByRef fromHistory As Boolean)
		Dim n As String
		'UPGRADE_ISSUE: Control PageNameText could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		n = mainForm.PageNameText.text
        Dim p As _Page
        p = formToPage()
		
		Call p.cook(model.getPagePreparer, model.getPageCooker, model.getSingleUserState.backlinks)
		Call model.setCurrentPage(p)
		If p.pageName = "" Then
			MsgBox("Can't save a page with no name. NOT SAVED!!!")
			
		Else
			Call model.savePage()
			model.getSingleUserState.changesSaved = True
            model.getSingleUserState.editState = PageEditState.SavedState
			
			' special actions for page types
			If p.isNetwork Then
				'UPGRADE_ISSUE: Control vse could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				mainForm.vse.setMode(VseCanvas.vseMode.View)
			End If
			
			Call showCooked()
			
			If fromHistory = False Then
				model.getSingleUserState.history.append(("#load " & pName))
			End If
			
		End If
		
	End Sub
	
	Public Sub actionPageHistory(ByRef pName As String, ByRef fromHistory As Boolean)
        Call saveGuard(PageEditState.LoadedState)
		Dim p As _Page
		p = model.makeHistoryPage(pName)
		'UPGRADE_ISSUE: Control showCookedPage could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call mainForm.showCookedPage(p)
		If fromHistory = False Then
			model.getSingleUserState.history.append(("#history " & pName))
		End If
	End Sub
	
	Public Sub actionRevert(ByRef tail As String, ByRef fromHistory As Boolean)
        Call saveGuard(PageEditState.LoadedState)
		Dim parts() As String
		Dim s As String
		s = Replace(tail, "+", " ")
		parts = Split(s, " ")
		Dim p As _Page
		p = model.getWikiAnnotatedDataStore.store.loadOldPage(CStr(parts(0)), CShort(parts(1)))
		'UPGRADE_ISSUE: Control showRawPage could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call mainForm.showRawPage(p)
		If fromHistory = False Then
			model.getSingleUserState.history.append(("#revert " & tail))
		End If
	End Sub
	
	Public Sub actionDir(ByRef path As String, ByRef fromHistory As Boolean)
        Call saveGuard(PageEditState.LoadedState)
		Dim p As _Page
		p = model.getLocalFileSystem.makeDirectoryPage(path)
		'UPGRADE_ISSUE: Control showCookedPage could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call mainForm.showCookedPage(p)
		If fromHistory = False Then
			model.getSingleUserState.history.append(("#dir " & path))
		End If
	End Sub
	
	Public Sub actionBack()
		Call saveGuard(True)
		'UPGRADE_ISSUE: Control vm could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		mainForm.vm.hideAll()
		Call model.getSingleUserState.history.back()
		Dim cs As String
		cs = model.getSingleUserState.history.getAtIndex()
		Call Me.processCommand(cs, True)
	End Sub
	
	
	Public Sub actionForward()
		Call saveGuard(True)
		'UPGRADE_ISSUE: Control vm could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		mainForm.vm.hideAll()
		Dim cs As String
		model.getSingleUserState.history.forward()
		cs = model.getSingleUserState.history.getAtIndex()
		Call Me.processCommand(cs, True)
	End Sub
	
	' Page Manipulation Actions
	
	Public Sub actionNew(ByRef fromHistory As Boolean)
        Call saveGuard(PageEditState.LoadedState)
		Call model.newPage()
		Call showRaw()
		If fromHistory = False Then
			model.getSingleUserState.history.append(("#new"))
		End If
	End Sub
	
	Public Sub actionNewNetwork(ByRef fromHistory As Boolean)
        Call saveGuard(PageEditState.LoadedState)
		Call model.newNetworkPage()
		'Call showRaw
		'UPGRADE_ISSUE: Control PageNameText could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		mainForm.PageNameText.text = ""
		'UPGRADE_ISSUE: Control vse could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		mainForm.vse.setMode(VseCanvas.vseMode.Edit)
		'UPGRADE_ISSUE: Control vm could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        mainForm.vm.showVse()
        Application.DoEvents()

		'UPGRADE_ISSUE: Control vse could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call mainForm.vse.draw(model.getCurrentPage, mainForm.vse.mode)
		
		If fromHistory = False Then
			model.getSingleUserState.history.append(("#newNet"))
		End If
	End Sub
	
	Public Sub actionDelete(ByRef pName As String)
		Call model.deletePage(pName)
		Call actionLoad(pName, True)
	End Sub
	
	' == Special pages ==================
	
	Public Sub actionStart(ByRef fromHistory As Boolean)
		Call actionLoad(model.getSystemConfigurations.startPage, fromHistory)
	End Sub
	
	Public Sub actionHelp(ByRef fromHistory As Boolean)
		Call actionLoad(model.getSystemConfigurations.helpIndexPage, fromHistory)
	End Sub
	
	Public Sub actionConfig(ByRef fromHistory As Boolean)
		Call actionLoad(model.getSystemConfigurations.configPage, fromHistory)
	End Sub
	
	Public Sub actionAll(ByRef fromHistory As Boolean)
		Call actionLoad(model.getSystemConfigurations.allPage, fromHistory)
	End Sub
	
	Public Sub actionRecent(ByRef fromHistory As Boolean)
		Call actionLoad(model.getSystemConfigurations.recentChangesPage, fromHistory)
	End Sub
	
	Public Sub actionAbout(ByRef fromHistory As Boolean)
        Call saveGuard(PageEditState.LoadedState)
		Call model.newPage()
		Dim s As String
		s = "== About SdiDesk ==" & vbCrLf
		s = s & "=== SdiDesk version 0.2.2 alpha === " & vbCrLf
		s = s & "SdiDesk is copyright : Phil Jones ( http://www.synaesmedia.net ), 2004-2005 "
		s = s & " and released under the Gnu General Public Licence ( http://www.gnu.org/licenses/gpl.html )" & vbCrLf & vbCrLf
		s = s & "<p>" & vbCrLf & "BOX<" & vbCrLf
		s = s & "=== Author's note ===" & vbCrLf
		s = s & "It's not a license condition, but I'd like it if you preserve an active and obvious link to http://www.nooranch.com/sdidesk/ "
		s = s & " or to any alternative I later suggest / designate."
		s = s & " Please report any bugs or suggestions to the author via the NooRanch site. "
		s = s & " And talk to me about sponsoring bug-fixes and modifications. :-)" & vbCrLf & vbCrLf & "-- PhilJones" & vbCrLf
		s = s & ">BOX"
		model.getCurrentPage.raw = s
		model.getCurrentPage.pageName = "#about"
		Call model.getCurrentPage.cook(model.getPagePreparer, model.getPageCooker, model.getSingleUserState.backlinks)
		Call showCooked()
		If fromHistory = False Then
			model.getSingleUserState.history.append(("#about"))
		End If
	End Sub
	
	
	' == Exports =========================
	
	Public Sub actionExport(ByRef exportName As String)
		' this action runs an export
		Call model.getExportSubsystem.doExport(exportName)
	End Sub
	
	Public Sub actionExports(ByRef fromHistory As Boolean)
		'UPGRADE_ISSUE: Control showCookedPage could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call mainForm.showCookedPage(model.getExportSubsystem.makeExportsPage())
		If fromHistory = False Then
			model.getSingleUserState.history.append(("#exports"))
		End If
	End Sub
	
	Public Sub actionExporters(ByRef fromHistory As Boolean)
		'UPGRADE_ISSUE: Control showCookedPage could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call mainForm.showCookedPage(model.getExportSubsystem.makeExportersPage())
		If fromHistory = False Then
			Call model.getSingleUserState.history.append("#exports")
		End If
	End Sub
	
	Public Sub actionInstantExport(ByRef pageName As String)
		'UPGRADE_ISSUE: Control showCookedPage could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call mainForm.showCookedPage(model.getExportSubsystem.makeChooseExporterPage(model.getSingleUserState.currentPageName))
	End Sub
	
	Public Sub actionExportOne(ByRef tail As String)
		Dim parts() As String
		parts = Split(tail, " ")
		If UBound(parts) > 0 Then
			Call model.getExportSubsystem.doInstantExport(parts(0), parts(1))
		Else
			MsgBox("bad args to exportOne :: " & tail)
		End If
	End Sub
	
	' == Searching ========================
	Public Sub actionFind(ByRef searchTerm As String, ByRef fromHistory As Boolean)
		' we send the full string here, because the search string should be
		' able to include spaces
		
        Call saveGuard(PageEditState.LoadedState)
		
		Dim p As _Page
		
		p = model.makeSearchResultsPage(searchTerm)
		'UPGRADE_ISSUE: Control showCookedPage could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call mainForm.showCookedPage(p)
		If fromHistory = False Then
			model.getSingleUserState.history.append(("#find " & searchTerm))
		End If
	End Sub
	
	Public Sub actionCrawl(ByRef full As String, ByRef fromHistory As Boolean)
		' expected form
		' #crawl crawlName startPage
		
        Call saveGuard(PageEditState.LoadedState)
		
		Dim parts() As String
		Dim crawlName, startPage As String
		
		parts = Split(full, " ")
		
		crawlName = parts(0)
		
		If UBound(parts) > 0 Then
			startPage = parts(1)
		Else
			startPage = "StartPage"
		End If
		
		Dim p As _Page
		
		p = model.getCrawlerSubsystem.makeCrawlResultsPage(crawlName, startPage)
		
		'UPGRADE_ISSUE: Control showCookedPage could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call mainForm.showCookedPage(p)
		If fromHistory = False Then
			model.getSingleUserState.history.append(("#crawl " & full))
		End If
	End Sub
	
	Public Sub actionCrawlers(ByRef fromHistory As Boolean)
		'UPGRADE_ISSUE: Control showCookedPage could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call mainForm.showCookedPage(model.getCrawlerSubsystem.makeCrawlersPage())
		If fromHistory = False Then
			model.getSingleUserState.history.append(("#crawlers"))
		End If
	End Sub
	
	Public Sub actionShowCrawler(ByRef crawlerName As String, ByRef fromHistory As Boolean)
		Dim p As _Page
		p = POLICY_getFactory().getNewPageInstance()
		Dim pc As _PageCrawler
		pc = model.getCrawlerSubsystem.crawlerManager.crawlers.Item(crawlerName)
		p.raw = pc.toString_Renamed()
		Call p.cook(model.getPagePreparer, model.getPageCooker, False)
		p.pageName = "#showCrawler " & crawlerName
		'UPGRADE_ISSUE: Control showCookedPage could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call mainForm.showCookedPage(p)
		If fromHistory = False Then
			model.getSingleUserState.history.append(("#showCrawler " & crawlerName))
		End If
	End Sub
	
	' == Word Count ===================
	Public Function actionWordCount(ByRef pageName As String) As Object
		MsgBox("Number of words in " & pageName & "=" & model.wordCount(pageName))
	End Function
	
	' == Shell to other programs ===========
	Public Function actionShell(ByRef tail As String, ByRef fromHistory As Boolean) As Object
		' first arg is the focus, eg, VbNormalFocus
		' second is the command string
		Dim parts() As String
		Dim s As String
		parts = Split(tail, "+")
		Dim wmg As New WikiMarkupGopher
		If UBound(parts) > 0 Then
			parts(1) = wmg.qq(parts(1))
			parts(0) = parts(0) & " " & parts(1)
		End If
		Call Shell(parts(0), AppWinStyle.NormalFocus)
	End Function
	
	' == Program State ================
	Public Sub actionBackLinksOn(ByRef fromHistory As Boolean)
		model.getSingleUserState.backlinks = True
	End Sub
	
	Public Sub actionBackLinksOff(ByRef fromHistory As Boolean)
		model.getSingleUserState.backlinks = False
	End Sub
	
	' == Exit =============================
	Public Sub actionExit()
        Call saveGuard(PageEditState.LoadedState)
		End
	End Sub
	
	' == Supporting methods ===============
	
    Public Sub saveGuard(ByRef state As PageEditState)
        Dim s As String
        If model.getSingleUserState.changesSaved = False Then
            s = CStr(MsgBox("There are unsaved changes. Do you want to save?", 4, "Save?"))
            If s = "6" Then ' 6 is "yes"
                Call actionSave((model.getCurrentPage().pageName), True)
            End If
            model.getSingleUserState.changesSaved = True
        End If
    End Sub
	
	Private Sub showRaw()
		'UPGRADE_ISSUE: Control showRawPage could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call mainForm.showRawPage(model.getSingleUserState.currentPage)
	End Sub
	
	Public Sub showCooked()
		'UPGRADE_ISSUE: Control showCookedPage could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		Call mainForm.showCookedPage(model.getSingleUserState.currentPage)
	End Sub
	
    Private Function formToPage() As _Page
        Dim p As _Page
        p = model.getSingleUserState.currentPage

        'UPGRADE_ISSUE: Control PageNameText could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        If Left(mainForm.PageNameText.Text, 1) = "#" Then
            ' this is actually a typed command
            ' so use saved page name
        Else
            'UPGRADE_ISSUE: Control PageNameText could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
            p.pageName = mainForm.PageNameText.Text
        End If

        'UPGRADE_ISSUE: Control vm could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        If mainForm.vm.mode = ViewerManager.ViewerManagerMode.vmmTable Then
            'UPGRADE_ISSUE: Control td could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
            Call mainForm.td.updatePage(p)
        Else
            'UPGRADE_ISSUE: Control RawText could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
            p.raw = mainForm.RawText.Text
        End If

        formToPage = p
    End Function
End Class