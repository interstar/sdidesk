Option Strict Off
Option Explicit On
Friend Class WADSMainForm
	Inherits System.Windows.Forms.Form
	' SdiDesk 0.2.1
	
	' SdiDesk is Copyright (c) Philip Jones, 2004-2005. All rights reserved.
	
	'  This program is free software; you can redistribute it and/or modify
	'  it under the terms of the GNU General Public License as published by
	'  the Free Software Foundation; either version 2 of the License, or
	'  (at your option) any later version.
	
	'  This program is distributed in the hope that it will be useful,
	'  but WITHOUT ANY WARRANTY; without even the implied warranty of
	'  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
	'  GNU General Public License for more details.
	
	'  You should have received a copy of the GNU General Public License
	'  along with this program; if not, write to the Free Software
	'  Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
	'  or from the web-site : http://www.gnu.org/licenses/gpl.html
	
	'  Author's Note : this copyright only covers the SdiDesk source-code
	'  by myself, Phil Jones, and any other contributers. It isn't intended to
	'  apply to any runtime libraries belonging to third parties that you may have
	'  also received as part of this package, and on which SdiDesk may depend.
	'  These libraries are distributed on the understanding that I, having paid
	'  for my copy of Visual Basic, am entitled to distribute them with any
	'  application I develop using it. That may not apply to you. And this
	'  license explicitly shouldn't be interpretted as granting you any rights
	'  over them.
	
	
	' Terminology Note :
	' VSE stands for "Visual Structure Editor" ... the canvas for editing networks
	' WADS stands for "Wiki Annotated Data Store" (now an interface to the
	' sub-part of the program
	
	' the configuration factory should be the ONLY object
	' that currently knows which concrete classes implement
	' most of the major interfaces like ModelLevel, PageStore, Page etc.
	
	Private factory As SdiDeskConfigurationFactory
	
	' MVC "model"
	Public MagicNotebook As _ModelLevel
	
	' MVC View
	Public vm As ViewerManager ' use to show / hide / resize all viewers
	Public vse As VseCanvas ' where we draw networks for the Visual Structure Editor
	Public td As TableDisplay ' where we edit tables
	' + functions of this form are part of view
	
	' MVC controller
	Dim docHandler As DocumentHandler ' used to trap events from the WebBrowser control
	Public controller As ControlLevel ' all user actions (button clicks, command line) go via this
	
	
	Public Sub showRawPage(ByRef p As _Page)
		Me.PageNameText.Text = p.pageName
		Me.DateCreatedLabel.Text = "Created : " & CStr(p.createdDate)
		Me.DateLastEditedLabel.Text = "Last Edited : " & CStr(p.lastEdited)
		MagicNotebook.getSingleUserState.isLoading = True
		Me.RawText.Text = p.raw
		MagicNotebook.getSingleUserState.isLoading = False
		'UPGRADE_WARNING: Couldn't resolve default property of object WADSMainForm.HtmlView.Document.body. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Me.HtmlView.Document.DomDocument.body.innerHTML = p.cooked
		
		Call vm.showRaw() ' viewer manager controls the visibility
		
	End Sub
	
	Public Sub showCookedPage(ByRef p As _Page)
		Me.PageNameText.Text = p.pageName
		Me.DateCreatedLabel.Text = "Created : " & CStr(p.createdDate)
		Me.DateLastEditedLabel.Text = "Last Edited : " & CStr(p.lastEdited)
		MagicNotebook.getSingleUserState.isLoading = True
		Me.RawText.Text = p.raw
		MagicNotebook.getSingleUserState.isLoading = False
		
		' actually this decision should be somewhere else too
		Dim n As Network
		If p.isNetwork Then
			n = POLICY_getFactory().wrapPageInNetwork(p)
            Call vm.showVse() ' viewer manager controls visibility
            Application.DoEvents()
			Call Me.vse.draw(n, (vse.mode))
			'UPGRADE_WARNING: Couldn't resolve default property of object n. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call MagicNotebook.getControllableModel.setCurrentPage(n)
		Else
			' show the text in the WebBrowser
			'UPGRADE_WARNING: Couldn't resolve default property of object WADSMainForm.HtmlView.Document.body. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Me.HtmlView.Document.DomDocument.body.innerHTML = p.cooked
			Call vm.showHtml()
			
			Call waitPageLoad()
			If Not docHandler Is Nothing Then
				Call docHandler.recalc()
			End If
			
		End If
		
	End Sub
	
	Private Sub AllButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles AllButton.Click
		Call controller.actionAll(False)
	End Sub
	
	Private Sub BackButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles BackButton.Click
		Call controller.actionBack()
	End Sub
	
	Private Sub EditableCell_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles EditableCell.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim shift As Short = eventArgs.KeyData \ &H10000
		Call td.editableCellKeyDown(KeyCode)
	End Sub
	
	Private Sub EditableCell_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EditableCell.Leave
		EditableCell.Visible = False
	End Sub
	
	Private Sub EditorButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EditorButton.Click
		Call controller.actionEdit(MagicNotebook.getControllableModel.getCurrentPage.pageName, False)
	End Sub
	
	
	
	
    Private Sub ExportButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Call controller.actionInstantExport(MagicNotebook.getControllableModel.getCurrentPage.pageName)
    End Sub
	
	Private Sub FindButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles FindButton.Click
		Call controller.processCommand("#find " & PageNameText.Text, False)
	End Sub
	
	' Actual form control stuff
	
	Private Sub WADSMainForm_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'On Error Resume Next
        Try

        
            ' set html viewer to blank
            'UPGRADE_WARNING: Navigate2 was upgraded to Navigate and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            Me.HtmlView.Navigate(New System.Uri(("about:blank")))

            ' setup the factory
            factory = POLICY_getFactory()


            ' init model level
            MagicNotebook = factory.getModelLevel

            Call MagicNotebook.setForm(Me) ' dependency injection

            Dim wads As _WikiAnnotatedDataStore
            Dim sysConf As _SystemConfigurations
            wads = MagicNotebook.getWikiAnnotatedDataStore
            sysConf = MagicNotebook.getSystemConfigurations

            sysConf.configPage = "ConfigPage"
            sysConf.startPage = "StartPage"
            sysConf.helpIndexPage = "HelpIndex"
            sysConf.allPage = "AllPages"
            sysConf.recentChangesPage = "RecentChanges"


            ' init viwer
            docHandler = New DocumentHandler
            docHandler.HtmlView = Me.HtmlView

            vm = New ViewerManager
            Call vm.init((Me.RawText), (Me.HtmlView), (Me.NetworkCanvas), (Me.TableEditor))

            vse = New VseCanvas
            Call vse.init((Me.NetworkCanvas), (Me.RawText), MagicNotebook.getControllableModel.getPageCooker, Me)

            td = New TableDisplay
            Call td.init((Me.TableEditor), (Me.EditableCell), (Me.RawText), Me)

            ' connect the comboBox on the form with the
            ' navigation history in the model
            Dim nh As NavigationHistory
            nh = MagicNotebook.getSingleUserState.history
            Call nh.setComboBox((Me.HistoryList))
            'UPGRADE_NOTE: Object nh may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            nh = Nothing

            ' init controller
            controller = New ControlLevel
            'UPGRADE_WARNING: Couldn't resolve default property of object MagicNotebook. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call controller.init(Me, MagicNotebook)

            ' let's go
            Call MagicNotebook.getControllableModel.loadNewPage(MagicNotebook.getSystemConfigurations.startPage)
            Call waitPageLoad()
            MagicNotebook.getSingleUserState.changesSaved = True

            'UPGRADE_WARNING: Couldn't resolve default property of object MagicNotebook. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call MagicNotebook.getExportSubsystem.refreshExportManager(MagicNotebook)

            Call showCooked()
            MagicNotebook.getSingleUserState.history.append(("StartPage"))
            Call setSize(VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width), VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - 420)
            Left = 0
            Top = 0
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
	
	
	Public Sub setSize(ByRef w As Integer, ByRef h As Integer)
		Me.Width = VB6.TwipsToPixelsX(w)
		Me.Height = VB6.TwipsToPixelsY(h)
	End Sub
	
	
	'UPGRADE_WARNING: Event WADSMainForm.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub WADSMainForm_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		On Error Resume Next
		' Resize controls if Form is not minimized
		Dim w, h As Integer
		Dim c As Integer
		If Me.WindowState <> System.Windows.Forms.FormWindowState.Minimized Then
			w = VB6.PixelsToTwipsX(Me.Width)
			h = VB6.PixelsToTwipsY(Me.Height)
			
			If w >= 8000 And h >= 8000 Then
				' only if bigger than minimal size
				
				' set the frames
				
				' left
				NavFrame.Left = VB6.TwipsToPixelsX(60)
				PageFrame.Left = VB6.TwipsToPixelsX(60)
				ViewFrame.Left = VB6.TwipsToPixelsX(60)
				CategoryFrame.Left = VB6.TwipsToPixelsX(60)
				
				' widths
				NavFrame.Width = VB6.TwipsToPixelsX(w - 260)
				PageFrame.Width = VB6.TwipsToPixelsX(w - 260)
				ViewFrame.Width = VB6.TwipsToPixelsX(w - 260)
				CategoryFrame.Width = VB6.TwipsToPixelsX(w - 260)
				
				' vertical
				CategoryFrame.Top = VB6.TwipsToPixelsY(h - 1150)
				c = VB6.PixelsToTwipsY(Me.CategoryFrame.Top) - 100
				Me.ViewFrame.Height = VB6.TwipsToPixelsY(c - VB6.PixelsToTwipsY(Me.ViewFrame.Top))
				
				
				' set the dimensions of the content objects
				
				' viewers
				Call vm.resize(VB6.PixelsToTwipsX(ViewFrame.Width), VB6.PixelsToTwipsY(ViewFrame.Height), 100)
				
				' other components
				FindButton.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(NavFrame.Width) - 650)
				GoButton.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(NavFrame.Width) - 1300)
				PageNameText.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(NavFrame.Width) - 4700)
			Else
				Me.Width = VB6.TwipsToPixelsX(8000)
				Me.Height = VB6.TwipsToPixelsY(8000)
			End If
		End If
	End Sub
	
	Public Sub loadPage(ByRef n As String)
		Dim pt As String
		pt = MagicNotebook.getControllableModel.loadNewPage(n)
		Select Case pt
			Case "normal"
				showCooked()
			Case "newPage"
				showRaw()
			Case "network"
				showCooked()
			Case "table"
				showCooked()
			Case Else
				MsgBox("came back as something else WADSMainForm:loadPage")
		End Select
		MagicNotebook.getSingleUserState.changesSaved = True
	End Sub
	
	Private Sub showRaw()
		Call showRawPage(MagicNotebook.getControllableModel.getCurrentPage)
	End Sub
	
	Public Sub showCooked()
		Call showCookedPage(MagicNotebook.getControllableModel.getCurrentPage)
	End Sub
	
	'UPGRADE_NOTE: Form_Terminate was upgraded to Form_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_WARNING: WADSMainForm event Form.Terminate has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub Form_Terminate_Renamed()
        Call controller.saveGuard(PageEditState.EditedState)
	End Sub
	
	Private Sub ForwardButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ForwardButton.Click
		Call controller.actionForward()
	End Sub
	
	Private Sub GoButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles GoButton.Click
		Call controller.processCommand((PageNameText.Text), False)
	End Sub
	
	Private Sub HelpButton_Click()
		Call controller.actionHelp(False)
	End Sub
	
	Private Sub HistoryButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles HistoryButton.Click
		Call controller.actionPageHistory(MagicNotebook.getControllableModel.getCurrentPage.pageName, False)
	End Sub
	
	'UPGRADE_WARNING: Event HistoryList.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub HistoryList_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles HistoryList.SelectedIndexChanged
		If (HistoryList.Text <> "Future" And HistoryList.Text <> "History") Then
			Call controller.processCommand((HistoryList.Text), False)
		End If
	End Sub
	
	'UPGRADE_ISSUE: ShDocW.WebBrowser.DocumentComplete pDisp was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
	Private Sub HtmlView_DocumentCompleted(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles HtmlView.DocumentCompleted
		Dim url As String = eventArgs.URL.ToString()
		' this is here because I was getting a blank page after
		' loading a page through the interupt. May cause problems though
		' as it's a hack
		
		Call showCooked()
	End Sub
	
	
	
	Public Sub menuAbout_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuAbout.Click
		Call controller.actionAbout(False)
	End Sub
	
	Public Sub menuAll_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuAll.Click
		Call controller.actionAll(False)
	End Sub
	
	Public Sub menuBack_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuBack.Click
		Call controller.actionBack()
	End Sub
	
	Public Sub menuBackLinks_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuBackLinks.Click
		If menuBackLinks.Checked = True Then
			menuBackLinks.Checked = False
			MagicNotebook.getSingleUserState.backlinks = False
		Else
			menuBackLinks.Checked = True
			MagicNotebook.getSingleUserState.backlinks = True
		End If
	End Sub
	
	Public Sub menuCrawlers_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuCrawlers.Click
		Call controller.actionLoad("CrawlerDefinitions", False)
	End Sub
	
	Public Sub menuDelete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuDelete.Click
		Call controller.actionDelete(MagicNotebook.getControllableModel.getCurrentPage.pageName)
	End Sub
	
	Public Function chooseDirectory() As String
		'UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
        'DirectoryChooser.CancelError = True
		DirectoryChooserSave.InitialDirectory = My.Application.Info.DirectoryPath
		Dim doit As Boolean
		doit = False
		
		On Error GoTo Cancelled ' most likely cause of error
		Call DirectoryChooserSave.ShowDialog()
		doit = True
Cancelled: 
		
		Dim fName As String
		If doit = True Then
			fName = DirectoryChooserSave.FileName
		Else
			' cancelled (most likely)
			fName = ""
		End If
		
		chooseDirectory = fName
	End Function
	
	Public Sub menuDirectoryChooser_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuDirectoryChooser.Click
		' this stuff is here, rather than as an action,
		' because it depends on the DirectoryChooser.
		' Should probably refactor to an action
		
		Dim fName As String
		fName = chooseDirectory()
		If fName <> "" Then
			Call MagicNotebook.getLocalFileSystem.changeDirectory(fName)
		Else
			' do nothing, probably a cancel
		End If
	End Sub
	
	Private Sub menuDynamicExport_Click(ByRef index As Short)
		' hi
		
	End Sub
	
	Public Sub menuEdit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuEdit.Click
		Call controller.actionEdit(MagicNotebook.getControllableModel.getCurrentPage.pageName, False)
	End Sub
	
	Public Sub menuExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuExit.Click
		Call controller.actionExit()
	End Sub
	
	Public Sub menuExports_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuExports.Click
		Call controller.actionLoad("ExportDefinitions", False)
	End Sub
	
	Public Sub menuExportThisPageHtml_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuExportThisPageHtml.Click
		Dim dirName As String
		dirName = chooseDirectory()
		Call controller.actionExportOne("web,, " & MagicNotebook.getSingleUserState.currentPage.pageName & ",, " & dirName)
	End Sub
	
	Public Sub menuForward_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuForward.Click
		Call controller.actionForward()
	End Sub
	
	Public Sub menuHelpIndex_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuHelpIndex.Click
		Call controller.actionHelp(False)
	End Sub
	
	Public Sub menuHistory_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuHistory.Click
		Call controller.actionPageHistory(MagicNotebook.getControllableModel.getCurrentPage.pageName, False)
	End Sub
	
	Public Sub menuInterMap_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuInterMap.Click
		Call controller.actionLoad("InterMap", False)
	End Sub
	
	Public Sub menuLinkTypeDefinitions_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuLinkTypeDefinitions.Click
		Call controller.actionLoad("LinkTypeDefinitions", False)
	End Sub
	
	Public Sub menuNewNet_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuNewNet.Click
		Call controller.actionNewNetwork(False)
	End Sub
	
	Public Sub menuNewPage_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuNewPage.Click
		Call controller.actionNew(False)
	End Sub
	
	Public Sub menuPageVariables_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuPageVariables.Click
		MsgBox(MagicNotebook.getControllableModel.getCurrentPage.varsToString)
	End Sub
	
	Public Sub menuPreview_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuPreview.Click
		Call controller.actionPreview(False)
	End Sub
	
	
	
	Public Sub menuRaw_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuRaw.Click
		Call controller.actionRaw(MagicNotebook.getControllableModel.getCurrentPage.pageName, False)
	End Sub
	
	Public Sub menuRecentChanges_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuRecentChanges.Click
		Call controller.actionRecent(False)
	End Sub
	
	
	Public Sub menuSavePage_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuSavePage.Click
		Call controller.actionSave(MagicNotebook.getControllableModel.getCurrentPage.pageName, False)
	End Sub
	
	Public Sub menuShowCrawlers_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuShowCrawlers.Click
		Call controller.actionCrawlers(False)
	End Sub
	
	Public Sub menuShowExporters_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuShowExporters.Click
		Call controller.actionExporters(False)
	End Sub
	
	Public Sub menuShowExports_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuShowExports.Click
		Call controller.actionExports(False)
	End Sub
	
	Public Sub menuShowHtml_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuShowHtml.Click
		MsgBox(MagicNotebook.getControllableModel.getCurrentPage.cooked)
	End Sub
	
	Public Sub menuShowInterMap_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuShowInterMap.Click
		MsgBox(MagicNotebook.getSystemConfigurations.interMap.toString_Renamed())
	End Sub
	
	Public Sub menuShowOutlinks_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuShowOutlinks.Click
		Dim mg As New WikiMarkupGopher
		Dim s As String
		'UPGRADE_WARNING: Couldn't resolve default property of object MagicNotebook. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		s = mg.getAllTargets(MagicNotebook.getControllableModel.getCurrentPage.raw, MagicNotebook)
		MsgBox(s)
		'UPGRADE_NOTE: Object mg may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mg = Nothing
	End Sub
	
	Public Sub menuShowPrepared_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuShowPrepared.Click
		MsgBox(MagicNotebook.getControllableModel.getCurrentPage.prepared)
	End Sub
	
	Public Sub menuStart_Renamed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles menuStart_Renamed.Click
		Call controller.actionStart(False)
	End Sub
	
	Private Sub NetworkCanvas_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles NetworkCanvas.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        'Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
        'Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y) 

        Dim x As Single = eventArgs.X
        Dim y As Single = eventArgs.Y
		'UPGRADE_WARNING: Couldn't resolve default property of object MagicNotebook.getControllableModel.getCurrentPage(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call vse.MouseDownOnCanvas(MagicNotebook.getControllableModel.getCurrentPage(), Button, shift, x, y)
	End Sub
	
	Private Sub NetworkCanvas_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles NetworkCanvas.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Dim net As Network
		If TypeName(MagicNotebook.getControllableModel.getCurrentPage()) = "Network" Then
			net = MagicNotebook.getControllableModel.getCurrentPage()
			If ((net.hitNodeDetect(x, y) > -1) Or (net.hitArcDetect(x, y).x > -1)) Then
				'UPGRADE_WARNING: PictureBox property NetworkCanvas.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
				NetworkCanvas.Cursor = System.Windows.Forms.Cursors.Default
			Else
				NetworkCanvas.Cursor = System.Windows.Forms.Cursors.Cross
			End If
		End If
	End Sub
	
	Private Sub NetworkCanvas_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles NetworkCanvas.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        'Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
        'Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        Dim x As Single = eventArgs.X
        Dim y As Single = eventArgs.Y
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If TypeName(MagicNotebook.getControllableModel.getCurrentPage()) = "Network" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object MagicNotebook.getControllableModel.getCurrentPage(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call vse.endDrag(MagicNotebook.getControllableModel.getCurrentPage(), Button, shift, CInt(x), CInt(y))
		End If
	End Sub
	
	Private Sub NewButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles NewButton.Click
		Call controller.actionNew(False)
	End Sub
	
	Private Sub NewNetworkButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles NewNetworkButton.Click
		Call controller.actionNewNetwork(False)
	End Sub
	
	Private Sub PageNameText_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles PageNameText.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If KeyAscii = 13 Then
			Call controller.processCommand((PageNameText.Text), False)
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	
	Private Sub PresentationButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles PresentationButton.Click
		Call controller.actionPreview(False)
	End Sub
	
	Private Sub RawButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles RawButton.Click
		Call controller.actionRaw(MagicNotebook.getControllableModel.getCurrentPage.pageName, False)
	End Sub
	
	
	
	Private Sub RawText_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles RawText.TextChanged
		If MagicNotebook.getSingleUserState.isLoading = False Then
			MagicNotebook.getSingleUserState.changesSaved = False
		End If
	End Sub
	
	Private Sub RecentButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles RecentButton.Click
		Call controller.actionRecent(False)
	End Sub
	
	Private Sub SaveButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SaveButton.Click
		Call controller.actionSave(MagicNotebook.getControllableModel.getCurrentPage.pageName, False)
	End Sub
	
	Private Sub StartPageButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles StartPageButton.Click
		Call controller.actionStart(False)
	End Sub
	
	Public Sub waitPageLoad()
		' Make sure the Page has loaded
		Do 
			System.Windows.Forms.Application.DoEvents()
		Loop While HtmlView.IsBusy
	End Sub
	
	Private Sub WADSMainForm_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		ArcInfo.Close()
		NodeInfo.Close()
	End Sub
	
	Private Sub TableEditor_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TableEditor.DblClick
		Call Me.td.cellEdit()
	End Sub
	
	Private Sub TableEditor_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSFlexGridLib.DMSFlexGridEvents_MouseDownEvent) Handles TableEditor.MouseDownEvent
		Call td.startRange()
	End Sub
	
	Private Sub TableEditor_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSFlexGridLib.DMSFlexGridEvents_KeyDownEvent) Handles TableEditor.KeyDownEvent
		If eventArgs.KeyCode = System.Windows.Forms.Keys.Return Then
			Call Me.td.cellEdit()
		End If
		If eventArgs.KeyCode = System.Windows.Forms.Keys.Delete Then
			TableEditor.Text = ""
		End If
	End Sub

    Private Sub PageFrame_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PageFrame.Paint

    End Sub
End Class