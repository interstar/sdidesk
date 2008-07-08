Option Strict Off
Option Explicit On
Friend Class Arc
	
	' represents an Arc in a network diagram
	
	' the direction an arc goes
	Public Enum ArcDirectionality
		noDirection ' undirected
		one ' one-way link
		two ' two-way link
	End Enum
	
	
	Public exists As Boolean ' if there is an arc, true else false
	Public label As String ' label
	Public direction As ArcDirectionality ' is there a notion of directionality?
	Public angle As Single ' what's angle of this arc?
	Public x As Integer ' arcs need a notional location
	Public y As Integer ' as a target for hit detection
	
	Public n1 As Short ' index of from node
	Public n2 As Short ' index of to node
	
	
	Public Function hit(ByRef ex As Integer, ByRef wy As Integer) As Boolean
		' test if a point is in the active target of an arc
		hit = False
		Dim drawSize As Integer
		Dim drawAspect As Short
		drawAspect = 1
		drawSize = 100
		If ex > x - drawSize And ex < x + drawSize Then
			If wy > y - (drawSize * drawAspect) And wy < y + (drawSize * drawAspect) Then
				hit = True
			End If
		End If
	End Function
End Class