Option Strict Off
Option Explicit On
Friend Class VCollection
	
	' VB6s collections suck, bigtime.
	
	Private col As Collection
	Private myKeys As Collection
	
	Public Sub add(ByRef value As Object, ByRef key As String)
		Call col.Add(value, key)
		Call myKeys.Add(key, key)
	End Sub
	
	Public Function Count() As Short
		Count = col.Count()
	End Function
	
	Public Function Item(ByRef k As String) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object col.Item(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Item. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Item = col.Item(k)
	End Function
	
	Public Sub Remove(ByRef k As Object)
		Call col.Remove(k)
		Call myKeys.Remove(k)
	End Sub
	
	' Find if a collection has a key, return true if it does
	Public Function hasKey(ByRef k As String) As Boolean
		Dim a As Object
		Dim b As Boolean
		b = False
		On Error GoTo notHere
		'UPGRADE_WARNING: Couldn't resolve default property of object col.Item(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object a. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		a = col.Item(k)
		b = True
notHere: 
		hasKey = b
	End Function
	
	
	Public Function toCollection() As Collection
		' return the collection for "for each"ing
		toCollection = col
	End Function
	
	Public Function keyCollection() As Collection
		keyCollection = myKeys
	End Function
	
	'UPGRADE_NOTE: toString was upgraded to toString_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function toString_Renamed() As String
		Dim s As String
		Dim i As Object
		s = ""
		For	Each i In myKeys
			
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object col.Item(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			s = s & "* (" & CStr(i) & ", " & CStr(col.Item(CStr(i))) & ")" & vbCrLf
		Next i
		toString_Renamed = s
	End Function
	
	Public Function keysToString() As String
		Dim s As String
		Dim i As Object
		s = ""
		For	Each i In myKeys
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			s = s & "* " & CStr(i) & vbCrLf
		Next i
		keysToString = s
	End Function
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		col = New Collection
		myKeys = New Collection
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object col may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		col = Nothing
		'UPGRADE_NOTE: Object myKeys may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		myKeys = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class