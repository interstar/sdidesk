Option Strict Off
Option Explicit On
Friend Class ArrayTool
	
	' couple of low level routines for arrays
	
	' this copies a subrange of one array of strings to another
	Public Function copyStringArray(ByRef a1() As String, ByRef a2() As String, ByRef s1 As Short, ByRef e1 As Short, ByRef s2 As Short) As Object
		Dim a2c, i As Short
		a2c = s2
		For i = s1 To e1
			a2(a2c) = a1(i)
			a2c = a2c + 1
		Next i
	End Function
	
	' this copies a subrange of one array of varianst to another
	Public Function copyVariantArray(ByRef a1() As Object, ByRef a2() As Object, ByRef s1 As Short, ByRef e1 As Short, ByRef s2 As Short) As Object
		Dim a2c, i As Short
		a2c = s2
		For i = s1 To e1
			a2(a2c) = a1(i)
			a2c = a2c + 1
		Next i
	End Function
	
	
	' return a string which lists the contents of an array
	' will break if the contents break cstr
	
	Public Function inspectArray(ByRef a1() As Object) As Object
		Dim l As Short
		Dim s As String
		l = UBound(a1) + 1
		s = ""
		Dim i As Short
		For i = 0 To l - 1
			s = s & CStr(a1(i)) & ",, "
		Next i
		'UPGRADE_WARNING: Couldn't resolve default property of object inspectArray. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		inspectArray = s
		
	End Function
End Class