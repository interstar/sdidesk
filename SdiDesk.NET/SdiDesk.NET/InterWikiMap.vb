Option Strict Off
Option Explicit On
Friend Class InterWikiMap
	
	' This is a mapping between other wiki names and their URLs
	
	Private map As VCollection
	Private nameMap As VCollection
	
	Public Sub add(ByRef url As String, ByRef key As String)
		Call map.add(url, key)
		Call nameMap.add(key, key)
	End Sub
	
	Public Function getUrl(ByRef key As String) As String
		If map.hasKey(key) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object map.Item(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			getUrl = CStr(map.Item(key))
		Else
			getUrl = "ERROR"
		End If
	End Function
	
	Public Sub parseFromString(ByRef s As String)
		Dim lines() As String
		Dim parts() As String
		Dim v As Object
		lines = Split(s, vbCrLf)
		For	Each v In lines
			'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			parts = Split(CStr(v), " ")
			If UBound(parts) > 0 Then
				Call Me.add(parts(1), parts(0))
			End If
		Next v
	End Sub
	
	'UPGRADE_NOTE: toString was upgraded to toString_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function toString_Renamed() As String
		Dim s As String
		Dim v As Object
		For	Each v In nameMap.toCollection
			'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			s = s & v & ", " & Me.getUrl(CStr(v)) & vbCrLf
		Next v
		toString_Renamed = s
	End Function
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		map = New VCollection
		nameMap = New VCollection
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object map may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		map = Nothing
		'UPGRADE_NOTE: Object nameMap may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		nameMap = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class