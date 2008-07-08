Option Strict Off
Option Explicit On
Friend Class Node
	
	' data about a node object for networks
	
	Public name As String
	Public x As Integer
	Public y As Integer
	Public nodeType As String
	
	Public boxLeft As Integer
	Public boxRight As Integer
	Public boxTop As Integer
	Public boxBottom As Integer
	
	Public Sub init(ByRef ex As Integer, ByRef wy As Integer, ByRef n As String)
		x = ex
		y = wy
		nodeType = ""
		name = n
	End Sub
	
	Public Sub setType(ByRef s As String)
		nodeType = s
	End Sub
	
	Public Sub setHitBox(ByRef l As Integer, ByRef t As Integer, ByRef r As Integer, ByRef b As Integer)
		boxLeft = l
		boxTop = t
		boxBottom = b
		boxRight = r
	End Sub
	
	Public Function boxToString() As String
		boxToString = "(" & boxLeft & "," & boxTop & ")-" & "(" & boxRight & "," & boxBottom & ")"
	End Function
	
	
	Public Function hit(ByRef x As Integer, ByRef y As Integer) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object hit. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hit = False
		If x > boxLeft And x < boxRight And y > boxTop And y < boxBottom Then
			'UPGRADE_WARNING: Couldn't resolve default property of object hit. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			hit = True
		End If
	End Function
End Class