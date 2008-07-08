Option Strict Off
Option Explicit On
Friend Class ExportManager
	
	' The new export model goes like this :
	' different types of exporting are handled by separate custom programs.
	
	' "Exporter" is the name we will use for these programs
	' "Export" is the name we'll use for an invocation of an export,
	' based on a particular page-set, template etc.
	
	' For example, HtmlExporter could be the name of a program to
	' export pages as a flat HTML site. It is an "Exporter"
	
	' my_site might be a call of HtmlExporter, giving it a simple_recursive
	' PageCrawler (to collect a set of pages) starting on a page called MySiteHome.
	' my_site is an "Export"
	
	' The ExportRecord now holds only three pieces of data, which define an Export :
	' the name of the Export
	' the name of the Exporter (program which will do the exporting)
	' the name of a page which contains all the parameters to define the export
	
	' The ExportManager is still a table of ExportRecords but
	' it does NOT *do* the export
	
	' The main program does *not* include Exporters!!!
	' It simply calls external Exporter programs through the VB shell command
	
	' ExportManager is the object which does this
	' although ExportSubsystem also has a call it needs to pass on
	
	Public exportTable As OCollection ' stores the table of different exports
	Public exportNames As VCollection ' stores the export names
	
	Public exportPrograms As VCollection ' stores list of export programs
	
	Private st As StringTool ' always useful
	
	Public Sub parseFromRawString(ByRef s As String)
		' we're expecting a simple, double-comma separated table
		' name,, program,, parameter-page
		
		Dim t As New Table
		Dim i As Short
		Dim e As ExportRecord
		
		Dim aName As String
		Dim aProgram As String
		Dim aParamPage As String
		
		exportTable = New OCollection
		exportNames = New VCollection
		
		Call t.parseFromDoubleCommaString(s)
		
		For i = 0 To t.noRows - 1
			e = New ExportRecord
			
			'UPGRADE_WARNING: Couldn't resolve default property of object t.at(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			aName = st.strip(CStr(t.at(i, 0)))
			'UPGRADE_WARNING: Couldn't resolve default property of object t.at(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			aProgram = st.strip(CStr(t.at(i, 1)))
			'UPGRADE_WARNING: Couldn't resolve default property of object t.at(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			aParamPage = st.strip(CStr(t.at(i, 2)))
			
			Call e.init(aName, aProgram, aParamPage)
			Call exportTable.Add(e, aName)
			Call exportNames.Add(aName, aName)
			
		Next i
		'UPGRADE_NOTE: Object e may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		e = Nothing
	End Sub
	
	Public Sub scanForExporters(ByRef lfs As _LocalFileSystem)
		Dim vc As New VCollection
		Dim v As Object
		If lfs.hasLocalFileSystem Then
			vc = lfs.getDirectoryListingAsVCollection(lfs.getExporterDirectory)
			For	Each v In vc.toCollection
				'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Right(CStr(v), 4) = ".exe" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Call exportPrograms.Add(CStr(v), CStr(v))
				End If
			Next v
		Else
			MsgBox("Look, I'm very sorry but I can't seem to scan for exports on this system. No exporting is possible.")
		End If
	End Sub
	
	'UPGRADE_NOTE: toString was upgraded to toString_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function toString_Renamed() As String
		Dim o As Object
		Dim s As String
		s = ""
		For	Each o In exportTable.toCollection
			'UPGRADE_WARNING: Couldn't resolve default property of object o.toString. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			s = CStr(CDbl(s) + o.toString + CDbl(vbCrLf))
		Next o
		toString_Renamed = s
	End Function
	
	Public Function exportersToString() As String
		Dim v As Object
		Dim s As String
		s = ""
		For	Each v In exportPrograms.toCollection
			'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			s = s & v & ", "
		Next v
		exportersToString = s
		
	End Function
	
	Public Sub callExport(ByRef aName As String, ByRef pageStoreIdentifier As String)
		' pageStoreIdentifier is a string which will let the exporter
		' find the page store. Currently, it will be a directory,
		' in future, may be a URL
		
		Dim e As ExportRecord
		e = exportTable.Item(aName)
		On Error GoTo notFound
		Dim pn As String
		'UPGRADE_WARNING: Couldn't resolve default property of object e.program. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pn = "exporters\" & e.program
		'UPGRADE_WARNING: Couldn't resolve default property of object e.paramPage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call Shell(pn & " -name " & e.name & " -param '" & e.paramPage & "' -psi '" & pageStoreIdentifier & "\'", AppWinStyle.NormalFocus)
		Exit Sub
		
notFound: 
		MsgBox("Couldn't find a program called '" & pn & "' in your exporters directory. Check what exporters are available (Export Menu:Show Exporters)")
		
	End Sub
	
	Public Sub callInstantExport(ByRef progName As String, ByRef pageStoreIdentifier As String, ByRef currentPageName As String)
		On Error GoTo notFound
		Dim pn As String
		pn = "exporters\" & progName
		Call Shell(pn & " -page '" & currentPageName & "' -psi '" & pageStoreIdentifier & "\'")
		Exit Sub
		
notFound: 
		MsgBox("Couldn't find a program called '" & pn & "' in your exporters directory. Check what exporters are available (Export Menu:Show Exporters)")
		
	End Sub
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		exportTable = New OCollection
		exportNames = New VCollection
		exportPrograms = New VCollection
		
		st = New StringTool
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object exportTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		exportTable = Nothing
		'UPGRADE_NOTE: Object exportNames may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		exportNames = Nothing
		'UPGRADE_NOTE: Object exportPrograms may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		exportPrograms = Nothing
		'UPGRADE_NOTE: Object st may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		st = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class