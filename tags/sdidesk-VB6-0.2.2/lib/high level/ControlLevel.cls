VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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

Public mainForm As Form
Public model As ControllableModel

' ========================
' Methods
' ========================


' == Outer Methods ==

Public Sub init(f As Form, m As ControllableModel)
   ' f is the main form
   Set mainForm = f
   Set model = m
End Sub


Public Sub processCommand(s As String, fromHistory As Boolean)
  ' fromHistory is true if this came from a history action
  
  Dim command As String, pageName As String, tail As String
  
  ' split command into command and argument with a NavCommand
  Dim nc As New NavCommand
  nc.init (s)
  command = nc.getCommand
  pageName = nc.getPageName
  tail = nc.tail
  
  If fromHistory = False Then
    ' wipe futurity
    Call model.getSingleUserState.history.wipeFuture
  End If
  
  ' now decide what to do with it
  Select Case command
    
   ' View / Page Navigation Actions
    Case "#load":
      Call actionLoad(pageName, fromHistory)
    
    Case "#raw":
      Call actionRaw(pageName, fromHistory)
    
    Case "#preview":
      Call actionPreview(fromHistory)
      ' NB : doesn't make sense to preview a pageName
      
    Case "#edit":
      Call actionEdit(pageName, fromHistory)
      
    Case "#save":
      Call actionSave(pageName, fromHistory)
      
    Case "#history":
      Call actionPageHistory(pageName, fromHistory)
      
    Case "#revert":
      Call actionRevert(tail, fromHistory)
      
    Case "#dir":
      Call actionDir(tail, fromHistory)
      
    Case "#back":
      Call actionBack
      
    Case "#forward":
      Call actionForward
      
    ' Page Manipulation Actions
    
    Case "#new":
      Call actionNew(fromHistory)
      
    Case "#newNet":
      Call actionNewNetwork(fromHistory)
    
    Case "#delete":
      Call actionDelete(pageName)
      
    ' Page processing
    Case "#wordCount":
      Call actionWordCount(pageName)
    
    ' special pages
    Case "#start":
      Call actionStart(fromHistory)
      
    Case "#config":
      Call actionConfig(fromHistory)
    
    Case "#help":
      Call actionHelp(fromHistory)
    
    Case "#all":
      Call actionAll(fromHistory)
      
    Case "#recent":
      Call actionRecent(fromHistory)
    
    Case "#about":
      Call actionAbout(fromHistory)
      
    ' export actions
    
    Case "#exporters": ' list of exporters (programs)
      Call actionExporters(fromHistory)
    
    Case "#exports": ' list of exports (preset, combination of program + pageset etc)
      Call actionExports(fromHistory)
      
    Case "#export": ' do an export
      Call actionExport(tail)
    
    Case "#exportOne": ' export a single page directly
      Call actionExportOne(tail)
      
    Case "#instantExport": ' makes a page of instantExport options
      Call actionInstantExport(pageName)
      
    ' search actions
    Case "#find":
      Call actionFind(tail, fromHistory)
      
    Case "#crawl":
      Call actionCrawl(tail, fromHistory)
      
    ' program state
    Case "#backLinksOn":
      Call actionBackLinksOn(fromHistory)
      
    Case "#backLinksOff":
      Call actionBackLinksOff(fromHistory)
      
    ' other useful info
    Case "#showCrawler":
      Call actionShowCrawler(pageName, fromHistory)
    
    Case "#crawlers":
      Call actionCrawlers(fromHistory)
    
    Case "#shell":
      Call actionShell(tail, fromHistory)
    
    Case Else
      MsgBox ("ControlLevel:processCommand. Sorry, don't know how to " & command & " when given " & s)
      
  End Select
    
End Sub

' == Action methods ===========
' View / Page Navigation Actions

Public Sub actionLoad(pName As String, fromHistory As Boolean)
    Call saveGuard(LoadedState)
    Call mainForm.vm.hideAll
    If model.getWikiAnnotatedDataStore.pageExists(pName) Then
        model.loadNewPage (pName)
        showCooked
    Else
        model.loadNewPage (pName)
        showRaw
    End If
    
    model.getSingleUserState.editState = LoadedState
    If fromHistory = False Then
        model.getSingleUserState.history.append ("#load " + pName)
    End If
End Sub

Public Sub actionRaw(pName As String, fromHistory As Boolean)
  Call saveGuard(RawState)
  Call mainForm.vm.hideAll
  Call model.loadRawPage(pName)
  Call showRaw
  If fromHistory = False Then
    model.getSingleUserState.history.append ("#raw " + pName)
  End If
End Sub

Public Sub actionPreview(fromHistory As Boolean)
  Call saveGuard(PreviewState)
  Dim p As Page
  Set p = formToPage()
  Call p.cook(model.getPagePreparer, model.getPageCooker, model.getSingleUserState.backlinks)
  Call model.setCurrentPage(p)
  If p.isNetwork Then
     mainForm.vse.setMode (View)
  End If
  Call showCooked
  If fromHistory = False Then
    model.getSingleUserState.history.append ("#preview " + p.pageName)
  End If
End Sub

Public Sub actionEdit(pName As String, fromHistory As Boolean)
  Call saveGuard(EditedState)
  Call model.loadRawPage(pName)
  Call model.getCurrentPage().cook(model.getPagePreparer, model.getPageCooker, model.getSingleUserState.backlinks)
  If model.getCurrentPage().isNetwork Then
    mainForm.vse.setMode (Edit)
    Call mainForm.vm.showVse
    Call mainForm.vse.draw(model.getCurrentPage(), mainForm.vse.mode)
  Else
    If model.getCurrentPage.isTable Then
        Dim theTable As New table
        theTable.parseFromDoubleCommaString (model.getCurrentPage.raw)
        Call mainForm.td.fillFromTable(theTable)
        Call mainForm.vm.showTable
        Set theTable = Nothing
    Else
        Call showRaw
    End If
  End If
  If fromHistory = False Then
    model.getSingleUserState.history.append ("#edit " + pName)
  End If
End Sub

Public Sub actionSave(pName As String, fromHistory As Boolean)
  Dim n As String
  n = mainForm.PageNameText.text
  Dim p As Page
  Set p = formToPage()
  
  Call p.cook(model.getPagePreparer, model.getPageCooker, model.getSingleUserState.backlinks)
  Call model.setCurrentPage(p)
  If p.pageName = "" Then
    MsgBox ("Can't save a page with no name. NOT SAVED!!!")
  
  Else
    Call model.savePage
    model.getSingleUserState.changesSaved = True
    model.getSingleUserState.editState = SavedState
  
    ' special actions for page types
    If p.isNetwork Then
        mainForm.vse.setMode (View)
    End If
    
    Call showCooked

    If fromHistory = False Then
        model.getSingleUserState.history.append ("#load " + pName)
    End If
  
  End If
  
End Sub

Public Sub actionPageHistory(pName As String, fromHistory As Boolean)
  Call saveGuard(LoadedState)
  Dim p As Page
  Set p = model.makeHistoryPage(pName)
  Call mainForm.showCookedPage(p)
  If fromHistory = False Then
     model.getSingleUserState.history.append ("#history " + pName)
  End If
End Sub

Public Sub actionRevert(tail As String, fromHistory As Boolean)
    Call saveGuard(LoadedState)
    Dim parts() As String
    Dim s As String
    s = Replace(tail, "+", " ")
    parts = Split(s, " ")
    Dim p As Page
    Set p = model.getWikiAnnotatedDataStore.store.loadOldPage(CStr(parts(0)), CInt(parts(1)))
    Call mainForm.showRawPage(p)
    If fromHistory = False Then
      model.getSingleUserState.history.append ("#revert " & tail)
    End If
End Sub

Public Sub actionDir(path As String, fromHistory As Boolean)
  Call saveGuard(LoadedState)
  Dim p As Page
  Set p = model.getLocalFileSystem.makeDirectoryPage(path)
  Call mainForm.showCookedPage(p)
  If fromHistory = False Then
    model.getSingleUserState.history.append ("#dir " & path)
  End If
End Sub

Public Sub actionBack()
  Call saveGuard(True)
  mainForm.vm.hideAll
  Call model.getSingleUserState.history.back
  Dim cs As String
  cs = model.getSingleUserState.history.getAtIndex()
  Call Me.processCommand(cs, True)
End Sub


Public Sub actionForward()
  Call saveGuard(True)
  mainForm.vm.hideAll
  Dim cs As String
  model.getSingleUserState.history.forward
  cs = model.getSingleUserState.history.getAtIndex()
  Call Me.processCommand(cs, True)
End Sub

' Page Manipulation Actions
 
Public Sub actionNew(fromHistory As Boolean)
  Call saveGuard(LoadedState)
  Call model.newPage
  Call showRaw
  If fromHistory = False Then
    model.getSingleUserState.history.append ("#new")
  End If
End Sub

Public Sub actionNewNetwork(fromHistory As Boolean)
   Call saveGuard(LoadedState)
   Call model.newNetworkPage
   'Call showRaw
   mainForm.PageNameText.text = ""
   mainForm.vse.setMode (Edit)
   mainForm.vm.showVse
   Call mainForm.vse.draw(model.getCurrentPage, mainForm.vse.mode)

   If fromHistory = False Then
     model.getSingleUserState.history.append ("#newNet")
   End If
End Sub

Public Sub actionDelete(pName As String)
   Call model.deletePage(pName)
   Call actionLoad(pName, True)
End Sub

' == Special pages ==================

Public Sub actionStart(fromHistory As Boolean)
  Call actionLoad(model.getSystemConfigurations.startPage, fromHistory)
End Sub

Public Sub actionHelp(fromHistory As Boolean)
  Call actionLoad(model.getSystemConfigurations.helpIndexPage, fromHistory)
End Sub

Public Sub actionConfig(fromHistory As Boolean)
  Call actionLoad(model.getSystemConfigurations.configPage, fromHistory)
End Sub

Public Sub actionAll(fromHistory As Boolean)
  Call actionLoad(model.getSystemConfigurations.allPage, fromHistory)
End Sub

Public Sub actionRecent(fromHistory As Boolean)
  Call actionLoad(model.getSystemConfigurations.recentChangesPage, fromHistory)
End Sub

Public Sub actionAbout(fromHistory As Boolean)
  Call saveGuard(LoadedState)
  Call model.newPage
  Dim s As String
  s = "== About SdiDesk ==" + vbCrLf
  s = s + "=== SdiDesk version 0.2.2 alpha === " + vbCrLf
  s = s + "SdiDesk is copyright : Phil Jones ( http://www.synaesmedia.net ), 2004-2005 "
  s = s + " and released under the Gnu General Public Licence ( http://www.gnu.org/licenses/gpl.html )" + vbCrLf + vbCrLf
  s = s + "<p>" + vbCrLf + "BOX<" + vbCrLf
  s = s + "=== Author's note ===" + vbCrLf
  s = s + "It's not a license condition, but I'd like it if you preserve an active and obvious link to http://www.nooranch.com/sdidesk/ "
  s = s + " or to any alternative I later suggest / designate."
  s = s + " Please report any bugs or suggestions to the author via the NooRanch site. "
  s = s + " And talk to me about sponsoring bug-fixes and modifications. :-)" & vbCrLf & vbCrLf & "-- PhilJones" + vbCrLf
  s = s + ">BOX"
  model.getCurrentPage.raw = s
  model.getCurrentPage.pageName = "#about"
  Call model.getCurrentPage.cook(model.getPagePreparer, model.getPageCooker, model.getSingleUserState.backlinks)
  Call showCooked
  If fromHistory = False Then
    model.getSingleUserState.history.append ("#about")
  End If
End Sub


' == Exports =========================

Public Sub actionExport(exportName As String)
   ' this action runs an export
   Call model.getExportSubsystem.doExport(exportName)
End Sub

Public Sub actionExports(fromHistory As Boolean)
    Call mainForm.showCookedPage(model.getExportSubsystem.makeExportsPage())
    If fromHistory = False Then
      model.getSingleUserState.history.append ("#exports")
   End If
End Sub

Public Sub actionExporters(fromHistory As Boolean)
    Call mainForm.showCookedPage(model.getExportSubsystem.makeExportersPage())
    If fromHistory = False Then
      Call model.getSingleUserState.history.append("#exports")
   End If
End Sub

Public Sub actionInstantExport(pageName As String)
    Call mainForm.showCookedPage(model.getExportSubsystem.makeChooseExporterPage(model.getSingleUserState.currentPageName))
End Sub

Public Sub actionExportOne(tail As String)
    Dim parts() As String
    parts() = Split(tail, " ")
    If UBound(parts) > 0 Then
        Call model.getExportSubsystem.doInstantExport(parts(0), parts(1))
    Else
        MsgBox ("bad args to exportOne :: " & tail)
    End If
End Sub

' == Searching ========================
Public Sub actionFind(searchTerm As String, fromHistory As Boolean)
   ' we send the full string here, because the search string should be
   ' able to include spaces
      
  Call saveGuard(LoadedState)
    
  Dim p As Page
  
  Set p = model.makeSearchResultsPage(searchTerm)
  Call mainForm.showCookedPage(p)
  If fromHistory = False Then
     model.getSingleUserState.history.append ("#find " + searchTerm)
  End If
End Sub

Public Sub actionCrawl(full As String, fromHistory As Boolean)
   ' expected form
   ' #crawl crawlName startPage
      
  Call saveGuard(LoadedState)
  
  Dim parts() As String, crawlName As String, startPage As String
  
  parts = Split(full, " ")

  crawlName = parts(0)
  
  If UBound(parts) > 0 Then
    startPage = parts(1)
  Else
    startPage = "StartPage"
  End If
  
  Dim p As Page
  
  Set p = model.getCrawlerSubsystem.makeCrawlResultsPage(crawlName, startPage)
  
  Call mainForm.showCookedPage(p)
  If fromHistory = False Then
     model.getSingleUserState.history.append ("#crawl " + full)
  End If
End Sub

Public Sub actionCrawlers(fromHistory As Boolean)
    Call mainForm.showCookedPage(model.getCrawlerSubsystem.makeCrawlersPage())
    If fromHistory = False Then
      model.getSingleUserState.history.append ("#crawlers")
   End If
End Sub

Public Sub actionShowCrawler(crawlerName As String, fromHistory As Boolean)
    Dim p As Page
    Set p = POLICY_getFactory().getNewPageInstance()
    Dim pc As PageCrawler
    Set pc = model.getCrawlerSubsystem.crawlerManager.crawlers.Item(crawlerName)
    p.raw = pc.toString
    Call p.cook(model.getPagePreparer, model.getPageCooker, False)
    p.pageName = "#showCrawler " & crawlerName
    Call mainForm.showCookedPage(p)
    If fromHistory = False Then
      model.getSingleUserState.history.append ("#showCrawler " + crawlerName)
    End If
End Sub

' == Word Count ===================
Public Function actionWordCount(pageName As String)
    MsgBox ("Number of words in " & pageName & "=" & model.wordCount(pageName))
End Function

' == Shell to other programs ===========
Public Function actionShell(tail As String, fromHistory As Boolean)
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
    Call Shell(parts(0), vbNormalFocus)
End Function

' == Program State ================
Public Sub actionBackLinksOn(fromHistory As Boolean)
  model.getSingleUserState.backlinks = True
End Sub

Public Sub actionBackLinksOff(fromHistory As Boolean)
  model.getSingleUserState.backlinks = False
End Sub

' == Exit =============================
Public Sub actionExit()
  Call saveGuard(LoadedState)
  End
End Sub

' == Supporting methods ===============

Public Sub saveGuard(state As PageEditState)
   If model.getSingleUserState.changesSaved = False Then
      Dim s As String
      s = MsgBox("There are unsaved changes. Do you want to save?", 4, "Save?")
      If s = "6" Then ' 6 is "yes"
         Call actionSave(model.getCurrentPage().pageName, True)
      End If
      model.getSingleUserState.changesSaved = True
   End If
End Sub

Private Sub showRaw()
  Call mainForm.showRawPage(model.getSingleUserState.currentPage)
End Sub

Public Sub showCooked()
  Call mainForm.showCookedPage(model.getSingleUserState.currentPage)
End Sub

Private Function formToPage() As Page
   Dim p As Page
   Set p = model.getSingleUserState.currentPage
      
   If Left(mainForm.PageNameText, 1) = "#" Then
        ' this is actually a typed command
        ' so use saved page name
   Else
        p.pageName = mainForm.PageNameText
   End If
    
   If mainForm.vm.mode = vmmTable Then
        Call mainForm.td.updatePage(p)
   Else
        p.raw = mainForm.RawText.text
   End If
   
   Set formToPage = p
End Function

