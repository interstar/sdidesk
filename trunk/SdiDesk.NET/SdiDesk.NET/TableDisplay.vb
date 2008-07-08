Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.Compatibility
Friend Class TableDisplay
	
	' This is the code which manages a TableEditor
	' Doesn't actually create the Editor (that's given by
	' the main form) but does know about and manipulate it
	
    Private Editor As AxMSFlexGridLib.AxMSFlexGrid
	Private EditableCell As System.Windows.Forms.TextBox
	Private RawText As System.Windows.Forms.RichTextBox
	
    Public mainForm As WADSMainForm ' System.Windows.Forms.Form
	
	Public noRows As Short
	Public noCols As Short
	
	Public startRow As Short ' for selecting a range
	Public startCol As Short
	
	Public comment As String
	
    Public Sub init(ByRef te As AxMSFlexGridLib.AxMSFlexGrid, ByRef ec As System.Windows.Forms.TextBox, ByRef rt As System.Windows.Forms.RichTextBox, ByRef mf As System.Windows.Forms.Form)
        Editor = te
        EditableCell = ec
        RawText = rt
        mainForm = mf
    End Sub
	
	
	Public Sub fillFromTable(ByRef t As Table)
		Dim i, j As Short
		Me.noRows = t.noRows
		Me.noCols = t.noCols
		
		Editor.rows = t.noRows + 2
		Editor.Cols = t.noCols + 2
		Editor.Row = 0
		For j = 0 To t.noCols
			Editor.col = j + 1
			Editor.Text = CStr(t.atHeader(j))
		Next j
		For i = 0 To t.noRows
			For j = 0 To t.noCols
				Editor.Row = i + 1
				Editor.col = j + 1
				'UPGRADE_WARNING: Couldn't resolve default property of object t.at(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Editor.Text = CStr(t.at(i, j))
			Next j
		Next i
		
		comment = t.comment
		
	End Sub
	
	Public Function toTable() As Table
		Dim i, j As Short
		Dim t As New Table
		
		Call t.setUp((Me.noRows), (Me.noCols))
		
		Editor.Row = 0
		For j = 0 To t.noCols
			Editor.col = j + 1
			Call t.setHeader(j, (Editor.Text))
		Next j
		
		For i = 0 To t.noRows
			For j = 0 To t.noCols
				Editor.col = j + 1
				Editor.Row = i + 1
				Call t.putIn(i, j, (Editor.Text))
			Next j
		Next i
		t.comment = comment
		toTable = t
	End Function
	
	Public Function updatePage(ByRef p As _Page) As _Page
		Dim t As Table
		t = toTable()
		p.raw = t.spitAsPrettyPersist
		
		'UPGRADE_NOTE: Object t may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		t = Nothing
	End Function
	
	Public Sub cellEdit()
		EditableCell.Visible = True
		EditableCell.Width = VB6.TwipsToPixelsX(Editor.CellWidth)
		EditableCell.Height = VB6.TwipsToPixelsY(Editor.CellHeight)
		EditableCell.Top = VB6.TwipsToPixelsY(Editor.CellTop + VB6.PixelsToTwipsY(Editor.Top))
		
		'UPGRADE_ISSUE: Control ViewFrame could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		EditableCell.Left = VB6.TwipsToPixelsX(Editor.CellLeft + VB6.PixelsToTwipsX(Editor.Left) + Me.mainForm.ViewFrame.Left)
		EditableCell.Text = Editor.Text
		EditableCell.SelectionStart = 0
		EditableCell.SelectionLength = Len(EditableCell.Text)
		EditableCell.BringToFront()
		EditableCell.Focus()
		
	End Sub
	
	Public Sub startRange()
		startRow = Editor.Row
		startCol = Editor.col
	End Sub
	
	
	Public Sub editableCellKeyDown(ByRef KeyCode As Short)
		If KeyCode = System.Windows.Forms.Keys.Return Then
			Editor.Text = EditableCell.Text
			If Editor.Row = Editor.rows - 1 Then
				Editor.Row = Editor.Row
			Else
				Editor.Row = Editor.Row + 1
			End If
			Editor.Focus()
			EditableCell.Visible = False
		End If
	End Sub
	
	'Public Sub changed(p As Page)
	'     Dim s As String
	'     Dim t As table
	
	'  p.raw = n.spitAsPrettyPersist
	'  Set p.myNetwork = New Network
	' Call p.myNetwork.init(1, 200, 0.75)
	'  p.myNetwork.parseFromPrettyPersist (p.raw)
	'  Call draw(p.myNetwork, mode)
	'  RawText.text = p.myNetwork.spitAsPrettyPersist
	'  mainForm.MagicNotebook.changesSaved = False
	'End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object Editor may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Editor = Nothing
		'UPGRADE_NOTE: Object EditableCell may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		EditableCell = Nothing
		'UPGRADE_NOTE: Object RawText may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		RawText = Nothing
		'UPGRADE_NOTE: Object mainForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mainForm = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class