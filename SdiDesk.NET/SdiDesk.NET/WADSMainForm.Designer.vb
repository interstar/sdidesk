<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class WADSMainForm
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			Static fTerminateCalled As Boolean
			If Not fTerminateCalled Then
				Form_Terminate_renamed()
				fTerminateCalled = True
			End If
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents TableEditor As AxMSFlexGridLib.AxMSFlexGrid
	Public WithEvents EditableCell As System.Windows.Forms.TextBox
	Public WithEvents DirListBox As Microsoft.VisualBasic.Compatibility.VB6.DirListBox
	Public DirectoryChooserSave As System.Windows.Forms.SaveFileDialog
	Public WithEvents HistoryButton As System.Windows.Forms.Button
	Public WithEvents DateCreatedLabel As System.Windows.Forms.Label
	Public WithEvents DateLastEditedLabel As System.Windows.Forms.Label
	Public WithEvents CategoryFrame As System.Windows.Forms.Panel
	Public WithEvents HtmlView As System.Windows.Forms.WebBrowser
	Public WithEvents NetworkCanvas As System.Windows.Forms.PictureBox
	Public WithEvents RawText As System.Windows.Forms.RichTextBox
	Public WithEvents ViewFrame As System.Windows.Forms.Panel
	Public WithEvents FindButton As System.Windows.Forms.Button
	Public WithEvents RecentButton As System.Windows.Forms.Button
	Public WithEvents GoButton As System.Windows.Forms.Button
	Public WithEvents AllButton As System.Windows.Forms.Button
	Public WithEvents StartPageButton As System.Windows.Forms.Button
	Public WithEvents BackButton As System.Windows.Forms.Button
	Public WithEvents PageNameText As System.Windows.Forms.TextBox
	Public WithEvents ForwardButton As System.Windows.Forms.Button
	Public WithEvents NavFrame As System.Windows.Forms.Panel
    Public WithEvents HistoryList As System.Windows.Forms.ComboBox
	Public WithEvents NewNetworkButton As System.Windows.Forms.Button
	Public WithEvents SaveButton As System.Windows.Forms.Button
	Public WithEvents NewButton As System.Windows.Forms.Button
	Public WithEvents EditorButton As System.Windows.Forms.Button
	Public WithEvents RawButton As System.Windows.Forms.Button
	Public WithEvents PresentationButton As System.Windows.Forms.Button
	Public WithEvents PageFrame As System.Windows.Forms.Panel
    Public WithEvents menuNewPage As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuNewNet As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuSavePage As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _menuSep1_0 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents menuDelete As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _menuSep4_1 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents menuDirectoryChooser As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuSepSDJI As System.Windows.Forms.ToolStripSeparator
	Public WithEvents menuExit As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuFile As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuEdit As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuPreview As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuRaw As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuSep2 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents menuHistory As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuPage As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuBack As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuForward As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuSep3 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents menuStart_Renamed As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuRecentChanges As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuAll As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuStandard As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuExportThisPageHtml As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuShowExporters As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuShowExports As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuShowCrawlers As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuExport As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuInterMap As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuLinkTypeDefinitions As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuCrawlers As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuExports As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuSep7 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents menuBackLinks As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuSep8 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents menuSettings As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuHelpIndex As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuAbout As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuSep5 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents menuShowOutlinks As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuPageVariables As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuShowPrepared As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuShowHtml As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuShowInterMap As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents menuHelp As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(WADSMainForm))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.TableEditor = New AxMSFlexGridLib.AxMSFlexGrid
        Me.EditableCell = New System.Windows.Forms.TextBox
        Me.DirListBox = New Microsoft.VisualBasic.Compatibility.VB6.DirListBox
        Me.DirectoryChooserSave = New System.Windows.Forms.SaveFileDialog
        Me.CategoryFrame = New System.Windows.Forms.Panel
        Me.HistoryButton = New System.Windows.Forms.Button
        Me.DateCreatedLabel = New System.Windows.Forms.Label
        Me.DateLastEditedLabel = New System.Windows.Forms.Label
        Me.ViewFrame = New System.Windows.Forms.Panel
        Me.RawText = New System.Windows.Forms.RichTextBox
        Me.HtmlView = New System.Windows.Forms.WebBrowser
        Me.NetworkCanvas = New System.Windows.Forms.PictureBox
        Me.NavFrame = New System.Windows.Forms.Panel
        Me.FindButton = New System.Windows.Forms.Button
        Me.RecentButton = New System.Windows.Forms.Button
        Me.GoButton = New System.Windows.Forms.Button
        Me.AllButton = New System.Windows.Forms.Button
        Me.StartPageButton = New System.Windows.Forms.Button
        Me.BackButton = New System.Windows.Forms.Button
        Me.PageNameText = New System.Windows.Forms.TextBox
        Me.ForwardButton = New System.Windows.Forms.Button
        Me.PageFrame = New System.Windows.Forms.Panel
        Me.HistoryList = New System.Windows.Forms.ComboBox
        Me.NewNetworkButton = New System.Windows.Forms.Button
        Me.SaveButton = New System.Windows.Forms.Button
        Me.NewButton = New System.Windows.Forms.Button
        Me.EditorButton = New System.Windows.Forms.Button
        Me.RawButton = New System.Windows.Forms.Button
        Me.PresentationButton = New System.Windows.Forms.Button
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip
        Me.menuFile = New System.Windows.Forms.ToolStripMenuItem
        Me.menuNewPage = New System.Windows.Forms.ToolStripMenuItem
        Me.menuNewNet = New System.Windows.Forms.ToolStripMenuItem
        Me.menuSavePage = New System.Windows.Forms.ToolStripMenuItem
        Me._menuSep1_0 = New System.Windows.Forms.ToolStripSeparator
        Me.menuDelete = New System.Windows.Forms.ToolStripMenuItem
        Me._menuSep4_1 = New System.Windows.Forms.ToolStripSeparator
        Me.menuDirectoryChooser = New System.Windows.Forms.ToolStripMenuItem
        Me.menuSepSDJI = New System.Windows.Forms.ToolStripSeparator
        Me.menuExit = New System.Windows.Forms.ToolStripMenuItem
        Me.menuPage = New System.Windows.Forms.ToolStripMenuItem
        Me.menuEdit = New System.Windows.Forms.ToolStripMenuItem
        Me.menuPreview = New System.Windows.Forms.ToolStripMenuItem
        Me.menuRaw = New System.Windows.Forms.ToolStripMenuItem
        Me.menuSep2 = New System.Windows.Forms.ToolStripSeparator
        Me.menuHistory = New System.Windows.Forms.ToolStripMenuItem
        Me.menuStandard = New System.Windows.Forms.ToolStripMenuItem
        Me.menuBack = New System.Windows.Forms.ToolStripMenuItem
        Me.menuForward = New System.Windows.Forms.ToolStripMenuItem
        Me.menuSep3 = New System.Windows.Forms.ToolStripSeparator
        Me.menuStart_Renamed = New System.Windows.Forms.ToolStripMenuItem
        Me.menuRecentChanges = New System.Windows.Forms.ToolStripMenuItem
        Me.menuAll = New System.Windows.Forms.ToolStripMenuItem
        Me.menuExport = New System.Windows.Forms.ToolStripMenuItem
        Me.menuExportThisPageHtml = New System.Windows.Forms.ToolStripMenuItem
        Me.menuShowExporters = New System.Windows.Forms.ToolStripMenuItem
        Me.menuShowExports = New System.Windows.Forms.ToolStripMenuItem
        Me.menuShowCrawlers = New System.Windows.Forms.ToolStripMenuItem
        Me.menuSettings = New System.Windows.Forms.ToolStripMenuItem
        Me.menuInterMap = New System.Windows.Forms.ToolStripMenuItem
        Me.menuLinkTypeDefinitions = New System.Windows.Forms.ToolStripMenuItem
        Me.menuCrawlers = New System.Windows.Forms.ToolStripMenuItem
        Me.menuExports = New System.Windows.Forms.ToolStripMenuItem
        Me.menuSep7 = New System.Windows.Forms.ToolStripSeparator
        Me.menuBackLinks = New System.Windows.Forms.ToolStripMenuItem
        Me.menuSep8 = New System.Windows.Forms.ToolStripSeparator
        Me.menuHelp = New System.Windows.Forms.ToolStripMenuItem
        Me.menuHelpIndex = New System.Windows.Forms.ToolStripMenuItem
        Me.menuAbout = New System.Windows.Forms.ToolStripMenuItem
        Me.menuSep5 = New System.Windows.Forms.ToolStripSeparator
        Me.menuShowOutlinks = New System.Windows.Forms.ToolStripMenuItem
        Me.menuPageVariables = New System.Windows.Forms.ToolStripMenuItem
        Me.menuShowPrepared = New System.Windows.Forms.ToolStripMenuItem
        Me.menuShowHtml = New System.Windows.Forms.ToolStripMenuItem
        Me.menuShowInterMap = New System.Windows.Forms.ToolStripMenuItem
        CType(Me.TableEditor, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.CategoryFrame.SuspendLayout()
        Me.ViewFrame.SuspendLayout()
        CType(Me.NetworkCanvas, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.NavFrame.SuspendLayout()
        Me.PageFrame.SuspendLayout()
        Me.MainMenu1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableEditor
        '
        Me.TableEditor.Location = New System.Drawing.Point(8, 104)
        Me.TableEditor.Name = "TableEditor"
        Me.TableEditor.OcxState = CType(resources.GetObject("TableEditor.OcxState"), System.Windows.Forms.AxHost.State)
        Me.TableEditor.Size = New System.Drawing.Size(497, 417)
        Me.TableEditor.TabIndex = 27
        Me.TableEditor.Visible = False
        '
        'EditableCell
        '
        Me.EditableCell.AcceptsReturn = True
        Me.EditableCell.BackColor = System.Drawing.SystemColors.Window
        Me.EditableCell.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.EditableCell.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.EditableCell.ForeColor = System.Drawing.SystemColors.WindowText
        Me.EditableCell.Location = New System.Drawing.Point(168, 124)
        Me.EditableCell.MaxLength = 0
        Me.EditableCell.Name = "EditableCell"
        Me.EditableCell.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.EditableCell.Size = New System.Drawing.Size(89, 20)
        Me.EditableCell.TabIndex = 26
        Me.EditableCell.Text = "Text1"
        Me.EditableCell.Visible = False
        '
        'DirListBox
        '
        Me.DirListBox.BackColor = System.Drawing.SystemColors.Window
        Me.DirListBox.Cursor = System.Windows.Forms.Cursors.Default
        Me.DirListBox.Font = New System.Drawing.Font("Lucida Console", 12.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DirListBox.ForeColor = System.Drawing.SystemColors.WindowText
        Me.DirListBox.FormattingEnabled = True
        Me.DirListBox.IntegralHeight = False
        Me.DirListBox.Location = New System.Drawing.Point(8, 152)
        Me.DirListBox.Name = "DirListBox"
        Me.DirListBox.Size = New System.Drawing.Size(137, 114)
        Me.DirListBox.TabIndex = 24
        Me.DirListBox.Visible = False
        '
        'DirectoryChooserSave
        '
        Me.DirectoryChooserSave.FileName = "none"
        Me.DirectoryChooserSave.InitialDirectory = "App"
        Me.DirectoryChooserSave.Title = "Choose a new directory"
        '
        'CategoryFrame
        '
        Me.CategoryFrame.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.CategoryFrame.Controls.Add(Me.HistoryButton)
        Me.CategoryFrame.Controls.Add(Me.DateCreatedLabel)
        Me.CategoryFrame.Controls.Add(Me.DateLastEditedLabel)
        Me.CategoryFrame.Cursor = System.Windows.Forms.Cursors.Default
        Me.CategoryFrame.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CategoryFrame.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CategoryFrame.Location = New System.Drawing.Point(4, 532)
        Me.CategoryFrame.Name = "CategoryFrame"
        Me.CategoryFrame.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CategoryFrame.Size = New System.Drawing.Size(505, 29)
        Me.CategoryFrame.TabIndex = 9
        '
        'HistoryButton
        '
        Me.HistoryButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.HistoryButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.HistoryButton.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.HistoryButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.HistoryButton.Location = New System.Drawing.Point(376, 4)
        Me.HistoryButton.Name = "HistoryButton"
        Me.HistoryButton.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.HistoryButton.Size = New System.Drawing.Size(65, 21)
        Me.HistoryButton.TabIndex = 25
        Me.HistoryButton.Text = "History"
        Me.HistoryButton.UseVisualStyleBackColor = False
        '
        'DateCreatedLabel
        '
        Me.DateCreatedLabel.BackColor = System.Drawing.Color.Transparent
        Me.DateCreatedLabel.Cursor = System.Windows.Forms.Cursors.Default
        Me.DateCreatedLabel.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateCreatedLabel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.DateCreatedLabel.Location = New System.Drawing.Point(208, 4)
        Me.DateCreatedLabel.Name = "DateCreatedLabel"
        Me.DateCreatedLabel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.DateCreatedLabel.Size = New System.Drawing.Size(177, 25)
        Me.DateCreatedLabel.TabIndex = 19
        Me.DateCreatedLabel.Text = "Created : "
        '
        'DateLastEditedLabel
        '
        Me.DateLastEditedLabel.BackColor = System.Drawing.Color.Transparent
        Me.DateLastEditedLabel.Cursor = System.Windows.Forms.Cursors.Default
        Me.DateLastEditedLabel.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateLastEditedLabel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.DateLastEditedLabel.Location = New System.Drawing.Point(0, 4)
        Me.DateLastEditedLabel.Name = "DateLastEditedLabel"
        Me.DateLastEditedLabel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.DateLastEditedLabel.Size = New System.Drawing.Size(185, 25)
        Me.DateLastEditedLabel.TabIndex = 18
        Me.DateLastEditedLabel.Text = "Last Edited"
        '
        'ViewFrame
        '
        Me.ViewFrame.BackColor = System.Drawing.Color.Teal
        Me.ViewFrame.Controls.Add(Me.RawText)
        Me.ViewFrame.Controls.Add(Me.HtmlView)
        Me.ViewFrame.Controls.Add(Me.NetworkCanvas)
        Me.ViewFrame.Cursor = System.Windows.Forms.Cursors.Default
        Me.ViewFrame.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ViewFrame.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ViewFrame.Location = New System.Drawing.Point(4, 100)
        Me.ViewFrame.Name = "ViewFrame"
        Me.ViewFrame.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ViewFrame.Size = New System.Drawing.Size(617, 425)
        Me.ViewFrame.TabIndex = 8
        '
        'RawText
        '
        Me.RawText.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RawText.Location = New System.Drawing.Point(4, 4)
        Me.RawText.Name = "RawText"
        Me.RawText.Size = New System.Drawing.Size(610, 409)
        Me.RawText.TabIndex = 12
        Me.RawText.Text = ""
        '
        'HtmlView
        '
        Me.HtmlView.Location = New System.Drawing.Point(4, 4)
        Me.HtmlView.Name = "HtmlView"
        Me.HtmlView.Size = New System.Drawing.Size(493, 401)
        Me.HtmlView.TabIndex = 13
        '
        'NetworkCanvas
        '
        Me.NetworkCanvas.BackColor = System.Drawing.Color.White
        Me.NetworkCanvas.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.NetworkCanvas.Cursor = System.Windows.Forms.Cursors.Cross
        Me.NetworkCanvas.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NetworkCanvas.ForeColor = System.Drawing.SystemColors.ControlText
        Me.NetworkCanvas.Location = New System.Drawing.Point(8, 8)
        Me.NetworkCanvas.Name = "NetworkCanvas"
        Me.NetworkCanvas.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.NetworkCanvas.Size = New System.Drawing.Size(497, 401)
        Me.NetworkCanvas.TabIndex = 16
        Me.NetworkCanvas.TabStop = False
        Me.NetworkCanvas.Visible = False
        '
        'NavFrame
        '
        Me.NavFrame.BackColor = System.Drawing.Color.Teal
        Me.NavFrame.Controls.Add(Me.FindButton)
        Me.NavFrame.Controls.Add(Me.RecentButton)
        Me.NavFrame.Controls.Add(Me.GoButton)
        Me.NavFrame.Controls.Add(Me.AllButton)
        Me.NavFrame.Controls.Add(Me.StartPageButton)
        Me.NavFrame.Controls.Add(Me.BackButton)
        Me.NavFrame.Controls.Add(Me.PageNameText)
        Me.NavFrame.Controls.Add(Me.ForwardButton)
        Me.NavFrame.Cursor = System.Windows.Forms.Cursors.Default
        Me.NavFrame.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NavFrame.ForeColor = System.Drawing.SystemColors.ControlText
        Me.NavFrame.Location = New System.Drawing.Point(4, 28)
        Me.NavFrame.Name = "NavFrame"
        Me.NavFrame.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.NavFrame.Size = New System.Drawing.Size(617, 33)
        Me.NavFrame.TabIndex = 0
        '
        'FindButton
        '
        Me.FindButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.FindButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.FindButton.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FindButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.FindButton.Location = New System.Drawing.Point(550, 4)
        Me.FindButton.Name = "FindButton"
        Me.FindButton.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.FindButton.Size = New System.Drawing.Size(44, 25)
        Me.FindButton.TabIndex = 23
        Me.FindButton.Text = "Find"
        Me.FindButton.UseVisualStyleBackColor = False
        '
        'RecentButton
        '
        Me.RecentButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.RecentButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.RecentButton.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RecentButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.RecentButton.Location = New System.Drawing.Point(132, 4)
        Me.RecentButton.Name = "RecentButton"
        Me.RecentButton.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.RecentButton.Size = New System.Drawing.Size(57, 25)
        Me.RecentButton.TabIndex = 21
        Me.RecentButton.Text = "Recent"
        Me.RecentButton.UseVisualStyleBackColor = False
        '
        'GoButton
        '
        Me.GoButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.GoButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.GoButton.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GoButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.GoButton.Location = New System.Drawing.Point(506, 4)
        Me.GoButton.Name = "GoButton"
        Me.GoButton.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.GoButton.Size = New System.Drawing.Size(41, 25)
        Me.GoButton.TabIndex = 17
        Me.GoButton.Text = "Go!"
        Me.GoButton.UseVisualStyleBackColor = False
        '
        'AllButton
        '
        Me.AllButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.AllButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.AllButton.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AllButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.AllButton.Location = New System.Drawing.Point(192, 4)
        Me.AllButton.Name = "AllButton"
        Me.AllButton.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.AllButton.Size = New System.Drawing.Size(41, 25)
        Me.AllButton.TabIndex = 15
        Me.AllButton.Text = "All"
        Me.AllButton.UseVisualStyleBackColor = False
        '
        'StartPageButton
        '
        Me.StartPageButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.StartPageButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.StartPageButton.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StartPageButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.StartPageButton.Location = New System.Drawing.Point(76, 4)
        Me.StartPageButton.Name = "StartPageButton"
        Me.StartPageButton.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPageButton.Size = New System.Drawing.Size(53, 25)
        Me.StartPageButton.TabIndex = 14
        Me.StartPageButton.Text = "Start"
        Me.StartPageButton.UseVisualStyleBackColor = False
        '
        'BackButton
        '
        Me.BackButton.BackColor = System.Drawing.SystemColors.Control
        Me.BackButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.BackButton.Font = New System.Drawing.Font("Arial", 13.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BackButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BackButton.Location = New System.Drawing.Point(4, 4)
        Me.BackButton.Name = "BackButton"
        Me.BackButton.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.BackButton.Size = New System.Drawing.Size(33, 25)
        Me.BackButton.TabIndex = 3
        Me.BackButton.Text = "<"
        Me.BackButton.UseVisualStyleBackColor = False
        '
        'PageNameText
        '
        Me.PageNameText.AcceptsReturn = True
        Me.PageNameText.BackColor = System.Drawing.SystemColors.Window
        Me.PageNameText.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.PageNameText.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PageNameText.ForeColor = System.Drawing.SystemColors.WindowText
        Me.PageNameText.Location = New System.Drawing.Point(240, 5)
        Me.PageNameText.MaxLength = 0
        Me.PageNameText.Name = "PageNameText"
        Me.PageNameText.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.PageNameText.Size = New System.Drawing.Size(260, 26)
        Me.PageNameText.TabIndex = 2
        '
        'ForwardButton
        '
        Me.ForwardButton.BackColor = System.Drawing.SystemColors.Control
        Me.ForwardButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.ForwardButton.Font = New System.Drawing.Font("Arial", 13.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForwardButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ForwardButton.Location = New System.Drawing.Point(40, 4)
        Me.ForwardButton.Name = "ForwardButton"
        Me.ForwardButton.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ForwardButton.Size = New System.Drawing.Size(33, 25)
        Me.ForwardButton.TabIndex = 1
        Me.ForwardButton.Text = ">"
        Me.ForwardButton.UseVisualStyleBackColor = False
        '
        'PageFrame
        '
        Me.PageFrame.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.PageFrame.Controls.Add(Me.HistoryList)
        Me.PageFrame.Controls.Add(Me.NewNetworkButton)
        Me.PageFrame.Controls.Add(Me.SaveButton)
        Me.PageFrame.Controls.Add(Me.NewButton)
        Me.PageFrame.Controls.Add(Me.EditorButton)
        Me.PageFrame.Controls.Add(Me.RawButton)
        Me.PageFrame.Controls.Add(Me.PresentationButton)
        Me.PageFrame.Cursor = System.Windows.Forms.Cursors.Default
        Me.PageFrame.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PageFrame.ForeColor = System.Drawing.SystemColors.ControlText
        Me.PageFrame.Location = New System.Drawing.Point(4, 64)
        Me.PageFrame.Name = "PageFrame"
        Me.PageFrame.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.PageFrame.Size = New System.Drawing.Size(617, 33)
        Me.PageFrame.TabIndex = 4
        '
        'HistoryList
        '
        Me.HistoryList.BackColor = System.Drawing.SystemColors.Window
        Me.HistoryList.Cursor = System.Windows.Forms.Cursors.Default
        Me.HistoryList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.HistoryList.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.HistoryList.ForeColor = System.Drawing.SystemColors.WindowText
        Me.HistoryList.Location = New System.Drawing.Point(356, 4)
        Me.HistoryList.Name = "HistoryList"
        Me.HistoryList.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.HistoryList.Size = New System.Drawing.Size(153, 23)
        Me.HistoryList.TabIndex = 22
        '
        'NewNetworkButton
        '
        Me.NewNetworkButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.NewNetworkButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.NewNetworkButton.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NewNetworkButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.NewNetworkButton.Location = New System.Drawing.Point(56, 4)
        Me.NewNetworkButton.Name = "NewNetworkButton"
        Me.NewNetworkButton.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.NewNetworkButton.Size = New System.Drawing.Size(65, 25)
        Me.NewNetworkButton.TabIndex = 20
        Me.NewNetworkButton.Text = "New Net"
        Me.NewNetworkButton.UseVisualStyleBackColor = False
        '
        'SaveButton
        '
        Me.SaveButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.SaveButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.SaveButton.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SaveButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.SaveButton.Location = New System.Drawing.Point(124, 4)
        Me.SaveButton.Name = "SaveButton"
        Me.SaveButton.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.SaveButton.Size = New System.Drawing.Size(49, 25)
        Me.SaveButton.TabIndex = 11
        Me.SaveButton.Text = "Save"
        Me.SaveButton.UseVisualStyleBackColor = False
        '
        'NewButton
        '
        Me.NewButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.NewButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.NewButton.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NewButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.NewButton.Location = New System.Drawing.Point(4, 4)
        Me.NewButton.Name = "NewButton"
        Me.NewButton.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.NewButton.Size = New System.Drawing.Size(49, 25)
        Me.NewButton.TabIndex = 10
        Me.NewButton.Text = "New"
        Me.NewButton.UseVisualStyleBackColor = False
        '
        'EditorButton
        '
        Me.EditorButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.EditorButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.EditorButton.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.EditorButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.EditorButton.Location = New System.Drawing.Point(176, 4)
        Me.EditorButton.Name = "EditorButton"
        Me.EditorButton.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.EditorButton.Size = New System.Drawing.Size(49, 25)
        Me.EditorButton.TabIndex = 7
        Me.EditorButton.Text = "Edit"
        Me.EditorButton.UseVisualStyleBackColor = False
        '
        'RawButton
        '
        Me.RawButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.RawButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.RawButton.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RawButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.RawButton.Location = New System.Drawing.Point(296, 4)
        Me.RawButton.Name = "RawButton"
        Me.RawButton.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.RawButton.Size = New System.Drawing.Size(57, 25)
        Me.RawButton.TabIndex = 6
        Me.RawButton.Text = "Raw"
        Me.RawButton.UseVisualStyleBackColor = False
        '
        'PresentationButton
        '
        Me.PresentationButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.PresentationButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.PresentationButton.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PresentationButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.PresentationButton.Location = New System.Drawing.Point(228, 4)
        Me.PresentationButton.Name = "PresentationButton"
        Me.PresentationButton.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.PresentationButton.Size = New System.Drawing.Size(65, 25)
        Me.PresentationButton.TabIndex = 5
        Me.PresentationButton.Text = "Preview"
        Me.PresentationButton.UseVisualStyleBackColor = False
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuFile, Me.menuPage, Me.menuStandard, Me.menuExport, Me.menuSettings, Me.menuHelp})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(633, 24)
        Me.MainMenu1.TabIndex = 28
        '
        'menuFile
        '
        Me.menuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuNewPage, Me.menuNewNet, Me.menuSavePage, Me._menuSep1_0, Me.menuDelete, Me._menuSep4_1, Me.menuDirectoryChooser, Me.menuSepSDJI, Me.menuExit})
        Me.menuFile.Name = "menuFile"
        Me.menuFile.Size = New System.Drawing.Size(37, 20)
        Me.menuFile.Text = "&File"
        '
        'menuNewPage
        '
        Me.menuNewPage.Name = "menuNewPage"
        Me.menuNewPage.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.N), System.Windows.Forms.Keys)
        Me.menuNewPage.Size = New System.Drawing.Size(191, 22)
        Me.menuNewPage.Text = "&New Page"
        '
        'menuNewNet
        '
        Me.menuNewNet.Name = "menuNewNet"
        Me.menuNewNet.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.M), System.Windows.Forms.Keys)
        Me.menuNewNet.Size = New System.Drawing.Size(191, 22)
        Me.menuNewNet.Text = "New Network"
        '
        'menuSavePage
        '
        Me.menuSavePage.Name = "menuSavePage"
        Me.menuSavePage.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.S), System.Windows.Forms.Keys)
        Me.menuSavePage.Size = New System.Drawing.Size(191, 22)
        Me.menuSavePage.Text = "&Save Page"
        '
        '_menuSep1_0
        '
        Me._menuSep1_0.Name = "_menuSep1_0"
        Me._menuSep1_0.Size = New System.Drawing.Size(188, 6)
        '
        'menuDelete
        '
        Me.menuDelete.Name = "menuDelete"
        Me.menuDelete.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.D), System.Windows.Forms.Keys)
        Me.menuDelete.Size = New System.Drawing.Size(191, 22)
        Me.menuDelete.Text = "&Delete Page"
        '
        '_menuSep4_1
        '
        Me._menuSep4_1.Name = "_menuSep4_1"
        Me._menuSep4_1.Size = New System.Drawing.Size(188, 6)
        '
        'menuDirectoryChooser
        '
        Me.menuDirectoryChooser.Name = "menuDirectoryChooser"
        Me.menuDirectoryChooser.ShortcutKeys = System.Windows.Forms.Keys.F1
        Me.menuDirectoryChooser.Size = New System.Drawing.Size(191, 22)
        Me.menuDirectoryChooser.Text = "Change Directory"
        '
        'menuSepSDJI
        '
        Me.menuSepSDJI.Name = "menuSepSDJI"
        Me.menuSepSDJI.Size = New System.Drawing.Size(188, 6)
        '
        'menuExit
        '
        Me.menuExit.Name = "menuExit"
        Me.menuExit.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.W), System.Windows.Forms.Keys)
        Me.menuExit.Size = New System.Drawing.Size(191, 22)
        Me.menuExit.Text = "Exit"
        '
        'menuPage
        '
        Me.menuPage.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuEdit, Me.menuPreview, Me.menuRaw, Me.menuSep2, Me.menuHistory})
        Me.menuPage.Name = "menuPage"
        Me.menuPage.Size = New System.Drawing.Size(45, 20)
        Me.menuPage.Text = "&Page"
        '
        'menuEdit
        '
        Me.menuEdit.Name = "menuEdit"
        Me.menuEdit.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.E), System.Windows.Forms.Keys)
        Me.menuEdit.Size = New System.Drawing.Size(157, 22)
        Me.menuEdit.Text = "&Edit"
        '
        'menuPreview
        '
        Me.menuPreview.Name = "menuPreview"
        Me.menuPreview.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.U), System.Windows.Forms.Keys)
        Me.menuPreview.Size = New System.Drawing.Size(157, 22)
        Me.menuPreview.Text = "Preview"
        '
        'menuRaw
        '
        Me.menuRaw.Name = "menuRaw"
        Me.menuRaw.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.R), System.Windows.Forms.Keys)
        Me.menuRaw.Size = New System.Drawing.Size(157, 22)
        Me.menuRaw.Text = "&Raw"
        '
        'menuSep2
        '
        Me.menuSep2.Name = "menuSep2"
        Me.menuSep2.Size = New System.Drawing.Size(154, 6)
        '
        'menuHistory
        '
        Me.menuHistory.Name = "menuHistory"
        Me.menuHistory.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.H), System.Windows.Forms.Keys)
        Me.menuHistory.Size = New System.Drawing.Size(157, 22)
        Me.menuHistory.Text = "&History"
        '
        'menuStandard
        '
        Me.menuStandard.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuBack, Me.menuForward, Me.menuSep3, Me.menuStart_Renamed, Me.menuRecentChanges, Me.menuAll})
        Me.menuStandard.Name = "menuStandard"
        Me.menuStandard.Size = New System.Drawing.Size(66, 20)
        Me.menuStandard.Text = "&Navigate"
        '
        'menuBack
        '
        Me.menuBack.Name = "menuBack"
        Me.menuBack.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.J), System.Windows.Forms.Keys)
        Me.menuBack.Size = New System.Drawing.Size(200, 22)
        Me.menuBack.Text = "Back"
        '
        'menuForward
        '
        Me.menuForward.Name = "menuForward"
        Me.menuForward.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.K), System.Windows.Forms.Keys)
        Me.menuForward.Size = New System.Drawing.Size(200, 22)
        Me.menuForward.Text = "Forward"
        '
        'menuSep3
        '
        Me.menuSep3.Name = "menuSep3"
        Me.menuSep3.Size = New System.Drawing.Size(197, 6)
        '
        'menuStart_Renamed
        '
        Me.menuStart_Renamed.Name = "menuStart_Renamed"
        Me.menuStart_Renamed.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.T), System.Windows.Forms.Keys)
        Me.menuStart_Renamed.Size = New System.Drawing.Size(200, 22)
        Me.menuStart_Renamed.Text = "S&tart Page"
        '
        'menuRecentChanges
        '
        Me.menuRecentChanges.Name = "menuRecentChanges"
        Me.menuRecentChanges.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.B), System.Windows.Forms.Keys)
        Me.menuRecentChanges.Size = New System.Drawing.Size(200, 22)
        Me.menuRecentChanges.Text = "Recent Changes"
        '
        'menuAll
        '
        Me.menuAll.Name = "menuAll"
        Me.menuAll.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.A), System.Windows.Forms.Keys)
        Me.menuAll.Size = New System.Drawing.Size(200, 22)
        Me.menuAll.Text = "&All"
        '
        'menuExport
        '
        Me.menuExport.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuExportThisPageHtml, Me.menuShowExporters, Me.menuShowExports, Me.menuShowCrawlers})
        Me.menuExport.Name = "menuExport"
        Me.menuExport.Size = New System.Drawing.Size(52, 20)
        Me.menuExport.Text = "Export"
        '
        'menuExportThisPageHtml
        '
        Me.menuExportThisPageHtml.Name = "menuExportThisPageHtml"
        Me.menuExportThisPageHtml.Size = New System.Drawing.Size(219, 22)
        Me.menuExportThisPageHtml.Text = "Export This Page (as HTML)"
        '
        'menuShowExporters
        '
        Me.menuShowExporters.Name = "menuShowExporters"
        Me.menuShowExporters.Size = New System.Drawing.Size(219, 22)
        Me.menuShowExporters.Text = "Show Exporters"
        '
        'menuShowExports
        '
        Me.menuShowExports.Name = "menuShowExports"
        Me.menuShowExports.Size = New System.Drawing.Size(219, 22)
        Me.menuShowExports.Text = "Show Exports"
        '
        'menuShowCrawlers
        '
        Me.menuShowCrawlers.Name = "menuShowCrawlers"
        Me.menuShowCrawlers.Size = New System.Drawing.Size(219, 22)
        Me.menuShowCrawlers.Text = "Show Crawlers"
        '
        'menuSettings
        '
        Me.menuSettings.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuInterMap, Me.menuLinkTypeDefinitions, Me.menuCrawlers, Me.menuExports, Me.menuSep7, Me.menuBackLinks, Me.menuSep8})
        Me.menuSettings.Name = "menuSettings"
        Me.menuSettings.Size = New System.Drawing.Size(61, 20)
        Me.menuSettings.Text = "Settings"
        '
        'menuInterMap
        '
        Me.menuInterMap.Name = "menuInterMap"
        Me.menuInterMap.Size = New System.Drawing.Size(130, 22)
        Me.menuInterMap.Text = "InterMap"
        '
        'menuLinkTypeDefinitions
        '
        Me.menuLinkTypeDefinitions.Name = "menuLinkTypeDefinitions"
        Me.menuLinkTypeDefinitions.Size = New System.Drawing.Size(130, 22)
        Me.menuLinkTypeDefinitions.Text = "Link Types"
        '
        'menuCrawlers
        '
        Me.menuCrawlers.Name = "menuCrawlers"
        Me.menuCrawlers.Size = New System.Drawing.Size(130, 22)
        Me.menuCrawlers.Text = "Crawlers"
        '
        'menuExports
        '
        Me.menuExports.Name = "menuExports"
        Me.menuExports.Size = New System.Drawing.Size(130, 22)
        Me.menuExports.Text = "Exports"
        '
        'menuSep7
        '
        Me.menuSep7.Name = "menuSep7"
        Me.menuSep7.Size = New System.Drawing.Size(127, 6)
        '
        'menuBackLinks
        '
        Me.menuBackLinks.Name = "menuBackLinks"
        Me.menuBackLinks.Size = New System.Drawing.Size(130, 22)
        Me.menuBackLinks.Text = "BackLinks"
        '
        'menuSep8
        '
        Me.menuSep8.Name = "menuSep8"
        Me.menuSep8.Size = New System.Drawing.Size(127, 6)
        '
        'menuHelp
        '
        Me.menuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuHelpIndex, Me.menuAbout, Me.menuSep5, Me.menuShowOutlinks, Me.menuPageVariables, Me.menuShowPrepared, Me.menuShowHtml, Me.menuShowInterMap})
        Me.menuHelp.Name = "menuHelp"
        Me.menuHelp.Size = New System.Drawing.Size(44, 20)
        Me.menuHelp.Text = "&Help"
        '
        'menuHelpIndex
        '
        Me.menuHelpIndex.Name = "menuHelpIndex"
        Me.menuHelpIndex.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.L), System.Windows.Forms.Keys)
        Me.menuHelpIndex.Size = New System.Drawing.Size(233, 22)
        Me.menuHelpIndex.Text = "He&lp Index"
        '
        'menuAbout
        '
        Me.menuAbout.Name = "menuAbout"
        Me.menuAbout.Size = New System.Drawing.Size(233, 22)
        Me.menuAbout.Text = "About SdiDesk"
        '
        'menuSep5
        '
        Me.menuSep5.Name = "menuSep5"
        Me.menuSep5.Size = New System.Drawing.Size(230, 6)
        '
        'menuShowOutlinks
        '
        Me.menuShowOutlinks.Name = "menuShowOutlinks"
        Me.menuShowOutlinks.ShortcutKeys = CType((System.Windows.Forms.Keys.Shift Or System.Windows.Forms.Keys.F1), System.Windows.Forms.Keys)
        Me.menuShowOutlinks.Size = New System.Drawing.Size(233, 22)
        Me.menuShowOutlinks.Text = "Show Outlinks"
        '
        'menuPageVariables
        '
        Me.menuPageVariables.Name = "menuPageVariables"
        Me.menuPageVariables.ShortcutKeys = CType((System.Windows.Forms.Keys.Shift Or System.Windows.Forms.Keys.F2), System.Windows.Forms.Keys)
        Me.menuPageVariables.Size = New System.Drawing.Size(233, 22)
        Me.menuPageVariables.Text = "Show Page Variables"
        '
        'menuShowPrepared
        '
        Me.menuShowPrepared.Name = "menuShowPrepared"
        Me.menuShowPrepared.ShortcutKeys = CType((System.Windows.Forms.Keys.Shift Or System.Windows.Forms.Keys.F3), System.Windows.Forms.Keys)
        Me.menuShowPrepared.Size = New System.Drawing.Size(233, 22)
        Me.menuShowPrepared.Text = "Show Prepared"
        '
        'menuShowHtml
        '
        Me.menuShowHtml.Name = "menuShowHtml"
        Me.menuShowHtml.ShortcutKeys = CType((System.Windows.Forms.Keys.Shift Or System.Windows.Forms.Keys.F4), System.Windows.Forms.Keys)
        Me.menuShowHtml.Size = New System.Drawing.Size(233, 22)
        Me.menuShowHtml.Text = "Show HTML"
        '
        'menuShowInterMap
        '
        Me.menuShowInterMap.Name = "menuShowInterMap"
        Me.menuShowInterMap.ShortcutKeys = CType((System.Windows.Forms.Keys.Shift Or System.Windows.Forms.Keys.F5), System.Windows.Forms.Keys)
        Me.menuShowInterMap.Size = New System.Drawing.Size(233, 22)
        Me.menuShowInterMap.Text = "Show InterMap"
        '
        'WADSMainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(633, 569)
        Me.Controls.Add(Me.EditableCell)
        Me.Controls.Add(Me.DirListBox)
        Me.Controls.Add(Me.CategoryFrame)
        Me.Controls.Add(Me.ViewFrame)
        Me.Controls.Add(Me.NavFrame)
        Me.Controls.Add(Me.PageFrame)
        Me.Controls.Add(Me.MainMenu1)
        Me.Controls.Add(Me.TableEditor)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(11, 57)
        Me.Name = "WADSMainForm"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "SdiDesk - (Version 0.2.2  ... another week, another release)"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.TableEditor, System.ComponentModel.ISupportInitialize).EndInit()
        Me.CategoryFrame.ResumeLayout(False)
        Me.ViewFrame.ResumeLayout(False)
        CType(Me.NetworkCanvas, System.ComponentModel.ISupportInitialize).EndInit()
        Me.NavFrame.ResumeLayout(False)
        Me.NavFrame.PerformLayout()
        Me.PageFrame.ResumeLayout(False)
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region 
End Class