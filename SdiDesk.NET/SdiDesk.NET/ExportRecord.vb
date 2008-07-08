Option Strict Off
Option Explicit On
Friend Class ExportRecord
	
	' Record of one type of export that the system
	' can do
	
	Public name As String ' the name of the export
	Public program As Object ' the name of the exporter program
	Public paramPage As Object ' the page where the properties of this export are defined
	
	' name,, program,, paramPage
	Public Sub init(ByRef aName As String, ByRef prog As String, ByRef pPage As String)
		name = aName
		'UPGRADE_WARNING: Couldn't resolve default property of object program. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		program = prog
		'UPGRADE_WARNING: Couldn't resolve default property of object paramPage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		paramPage = pPage
	End Sub
	
	
	'UPGRADE_NOTE: toString was upgraded to toString_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function toString_Renamed() As String
		'UPGRADE_WARNING: Couldn't resolve default property of object paramPage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object program. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		toString_Renamed = name & ", " + program + ", " + paramPage
	End Function
End Class