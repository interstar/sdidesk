<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class ArcInfo
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents DeleteCheckBox As System.Windows.Forms.CheckBox
	Public WithEvents CancelButton_Renamed As System.Windows.Forms.Button
	Public WithEvents OkButton As System.Windows.Forms.Button
	Public WithEvents ArcDirectionality As System.Windows.Forms.CheckBox
	Public WithEvents ArcName As System.Windows.Forms.TextBox
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(ArcInfo))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.DeleteCheckBox = New System.Windows.Forms.CheckBox
		Me.CancelButton_Renamed = New System.Windows.Forms.Button
		Me.OkButton = New System.Windows.Forms.Button
		Me.ArcDirectionality = New System.Windows.Forms.CheckBox
		Me.ArcName = New System.Windows.Forms.TextBox
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.BackColor = System.Drawing.Color.FromARGB(255, 192, 128)
		Me.Text = "Arc Information"
		Me.ClientSize = New System.Drawing.Size(216, 174)
		Me.Location = New System.Drawing.Point(4, 23)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "ArcInfo"
		Me.DeleteCheckBox.BackColor = System.Drawing.Color.FromARGB(255, 128, 128)
		Me.DeleteCheckBox.Text = "Delete this arc ?"
		Me.DeleteCheckBox.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.DeleteCheckBox.Size = New System.Drawing.Size(193, 29)
		Me.DeleteCheckBox.Location = New System.Drawing.Point(12, 96)
		Me.DeleteCheckBox.TabIndex = 5
		Me.DeleteCheckBox.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.DeleteCheckBox.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.DeleteCheckBox.CausesValidation = True
		Me.DeleteCheckBox.Enabled = True
		Me.DeleteCheckBox.ForeColor = System.Drawing.SystemColors.ControlText
		Me.DeleteCheckBox.Cursor = System.Windows.Forms.Cursors.Default
		Me.DeleteCheckBox.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.DeleteCheckBox.Appearance = System.Windows.Forms.Appearance.Normal
		Me.DeleteCheckBox.TabStop = True
		Me.DeleteCheckBox.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.DeleteCheckBox.Visible = True
		Me.DeleteCheckBox.Name = "DeleteCheckBox"
		Me.CancelButton_Renamed.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.CancelButton_Renamed
		Me.CancelButton_Renamed.Text = "Cancel"
		Me.CancelButton_Renamed.Size = New System.Drawing.Size(53, 33)
		Me.CancelButton_Renamed.Location = New System.Drawing.Point(152, 132)
		Me.CancelButton_Renamed.TabIndex = 4
		Me.CancelButton_Renamed.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.CancelButton_Renamed.BackColor = System.Drawing.SystemColors.Control
		Me.CancelButton_Renamed.CausesValidation = True
		Me.CancelButton_Renamed.Enabled = True
		Me.CancelButton_Renamed.ForeColor = System.Drawing.SystemColors.ControlText
		Me.CancelButton_Renamed.Cursor = System.Windows.Forms.Cursors.Default
		Me.CancelButton_Renamed.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.CancelButton_Renamed.TabStop = True
		Me.CancelButton_Renamed.Name = "CancelButton_Renamed"
		Me.OkButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.OkButton.Text = "OK"
		Me.OkButton.Size = New System.Drawing.Size(45, 33)
		Me.OkButton.Location = New System.Drawing.Point(12, 132)
		Me.OkButton.TabIndex = 3
		Me.OkButton.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.OkButton.BackColor = System.Drawing.SystemColors.Control
		Me.OkButton.CausesValidation = True
		Me.OkButton.Enabled = True
		Me.OkButton.ForeColor = System.Drawing.SystemColors.ControlText
		Me.OkButton.Cursor = System.Windows.Forms.Cursors.Default
		Me.OkButton.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.OkButton.TabStop = True
		Me.OkButton.Name = "OkButton"
		Me.ArcDirectionality.BackColor = System.Drawing.Color.Yellow
		Me.ArcDirectionality.Text = "Directional ?"
		Me.ArcDirectionality.CausesValidation = False
		Me.ArcDirectionality.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ArcDirectionality.Size = New System.Drawing.Size(193, 29)
		Me.ArcDirectionality.Location = New System.Drawing.Point(12, 60)
		Me.ArcDirectionality.TabIndex = 2
		Me.ArcDirectionality.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.ArcDirectionality.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.ArcDirectionality.Enabled = True
		Me.ArcDirectionality.ForeColor = System.Drawing.SystemColors.ControlText
		Me.ArcDirectionality.Cursor = System.Windows.Forms.Cursors.Default
		Me.ArcDirectionality.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ArcDirectionality.Appearance = System.Windows.Forms.Appearance.Normal
		Me.ArcDirectionality.TabStop = True
		Me.ArcDirectionality.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.ArcDirectionality.Visible = True
		Me.ArcDirectionality.Name = "ArcDirectionality"
		Me.ArcName.AutoSize = False
		Me.ArcName.Size = New System.Drawing.Size(197, 25)
		Me.ArcName.Location = New System.Drawing.Point(8, 28)
		Me.ArcName.TabIndex = 0
		Me.ArcName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ArcName.AcceptsReturn = True
		Me.ArcName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.ArcName.BackColor = System.Drawing.SystemColors.Window
		Me.ArcName.CausesValidation = True
		Me.ArcName.Enabled = True
		Me.ArcName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.ArcName.HideSelection = True
		Me.ArcName.ReadOnly = False
		Me.ArcName.Maxlength = 0
		Me.ArcName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.ArcName.MultiLine = False
		Me.ArcName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ArcName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.ArcName.TabStop = True
		Me.ArcName.Visible = True
		Me.ArcName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.ArcName.Name = "ArcName"
		Me.Label1.Text = "New name for this Arc"
		Me.Label1.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Size = New System.Drawing.Size(217, 25)
		Me.Label1.Location = New System.Drawing.Point(8, 4)
		Me.Label1.TabIndex = 1
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.BackColor = System.Drawing.Color.Transparent
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(DeleteCheckBox)
		Me.Controls.Add(CancelButton_Renamed)
		Me.Controls.Add(OkButton)
		Me.Controls.Add(ArcDirectionality)
		Me.Controls.Add(ArcName)
		Me.Controls.Add(Label1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class