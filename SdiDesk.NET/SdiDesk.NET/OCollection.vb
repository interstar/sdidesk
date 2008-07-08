Option Strict Off
Option Explicit On
Friend Class OCollection
	
	' VB6s collections suck, bigtime.
	
	Private col As Collection
	Private keys As Collection
	
	Public Sub Add(ByRef v As Object, ByRef k As String)
		Call col.Add(v, k)
		Call keys.Add(k, k)
	End Sub
	
	Public Function Count() As Short
		Count = col.Count()
	End Function
	
	Public Function Item(ByRef k As Object) As Object
		Item = col.Item(k)
	End Function
	
	Public Sub Remove(ByRef k As Object)
		Call col.Remove(k)
		Call keys.Remove(k)
	End Sub
	
	' Find if a collection has a key, return true if it does
	Public Function hasKey(ByRef k As String) As Boolean
		Dim a As Object
		Dim b As Boolean
		b = False
		On Error GoTo notHere
		a = col.Item(k)
		b = True
		'UPGRADE_NOTE: Object a may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		a = Nothing
notHere: 
		hasKey = b
	End Function
	
	Public Function toCollection() As Collection
		' return the collection for "for each"ing
		toCollection = col
	End Function
	
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		col = New Collection
		keys = New Collection
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object col may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		col = Nothing
		'UPGRADE_NOTE: Object keys may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		keys = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class