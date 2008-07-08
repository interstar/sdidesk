Option Strict Off
Option Explicit On
Friend Class HtmlTemplate
	
	' This is an object which holds template information
	' to wrap an ExportHTML page
	
	' It also knows how to parse a string
	
	Private myStyleSheet As String
	Private myHeader As String
	Private myFooter As String
	
	Public varDict As VCollection
	
	Public ReadOnly Property styleSheet() As String
		Get
			styleSheet = myStyleSheet
		End Get
	End Property
	
	Public ReadOnly Property header() As String
		Get
			header = myHeader
		End Get
	End Property
	
	Public ReadOnly Property footer() As String
		Get
			footer = myFooter
		End Get
	End Property
	
	Public Sub init(ByRef p As _Page)
		' expecting the three things to be on a page, separated by ----
		' var definitions should come afterwards
		
		Dim s As String
		
		s = p.raw
		Dim parts() As String
		If s = "new page" Or s = "" Then
			myStyleSheet = "<style></style>"
			myHeader = "<body> " & vbCrLf
			myFooter = ""
		Else
			parts = Split(s, "----")
			myStyleSheet = parts(0)
			myHeader = parts(1)
			myFooter = parts(2)
			
		End If
		
		varDict = p.getDataDictionary
		
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object varDict may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		varDict = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class