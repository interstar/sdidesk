Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.Compatibility
Friend Class StringTool
	
	' a few standard string processing functions
	
	Public Function trimRight(ByRef s As String) As String
		trimRight = Left(s, Len(s) - 1)
	End Function
	
	Public Function trimLeft(ByRef s As String) As String
		trimLeft = Right(s, Len(s) - 1)
	End Function
	
	Public Function stripLeft(ByRef s2 As String, ByRef c As String) As String
		Dim s As String
		s = s2
		While Left(s, 1) = c
			s = trimLeft(s)
		End While
		stripLeft = s
	End Function
	
	
	Public Function stripRight(ByRef s2 As String, ByRef c As String) As String
		Dim s As String
		s = s2
		While Right(s, 1) = c
			s = trimRight(s)
		End While
		stripRight = s
	End Function
	
	Public Function strip(ByRef s2 As String) As String
		Dim s As String
		s = s2
		While Right(s, 1) = " " Or Right(s, 1) = vbCrLf
			s = trimRight(s)
		End While
		While Left(s, 1) = " " Or Left(s, 1) = vbCrLf
			s = trimLeft(s)
		End While
		
		strip = s
		
	End Function
	
	
	Public Function losta(ByRef s As String, ByRef sep As String) As Short
		' length of string to array
		' returns the length of an array formed by breaking string at sep
		Dim parts() As String
		parts = Split(s, sep)
		losta = UBound(parts) + 1
	End Function
	
	Public Function star(ByRef s As String, ByRef sep As String, ByRef b As Short, ByRef e As Short) As Object
		' string to array range
		' turns a string into an array (splits on sep)
		' and returns a new string made up of the desired range
		' for example
		' - get the first element : star(s,sep,1,1)
		' - get the last n elements : star(s,sep,losta(s,sep)-n,losta(s,sep))
		Dim parts() As String
		Dim build As String
		Dim i As Short
		build = ""
		parts = Split(s, sep)
		For i = b To e
			build = build & parts(i) & sep
		Next i
		'UPGRADE_WARNING: Couldn't resolve default property of object star. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		star = build
	End Function
	
	Public Function leftsa(ByRef s As String, ByRef sep As String, ByRef i As Short) As String
		'UPGRADE_WARNING: Couldn't resolve default property of object star(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		leftsa = star(s, sep, 0, i - 1)
	End Function
	
	Public Function rightsa(ByRef s As String, ByRef sep As String, ByRef i As Short) As String
		'UPGRADE_WARNING: Couldn't resolve default property of object star(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rightsa = star(s, sep, losta(s, sep) - i, losta(s, sep) - 1)
	End Function
	
	Public Function stripHead(ByRef s As String, ByRef sep As String, ByRef n As Short) As String
		' remove the first n from the front
		stripHead = rightsa(s, " ", losta(s, " ") - n)
	End Function
	
	Public Sub seeAscii(ByRef s As String)
		' diagnostic functions
		Dim i As Short
		Dim b As String
		For i = 1 To Len(s)
			b = b & (CStr(i) & " : " & Mid(s, i, 1) & " : " & CStr(Asc(Mid(s, i, 1)))) & vbCrLf
		Next i
		MsgBox(b)
	End Sub
	
	Public Function removeDoubleChar(ByRef s As String, ByRef c As String) As String
		' useful for removing double \\
		removeDoubleChar = Replace(s, c & c, c)
	End Function
	
	Public Function mySplit(ByRef s As String, ByRef sep As String, ByRef esc As String) As String()
		Dim parts() As String
		Dim i As Short
		Dim s2 As String
		Dim parts2() As String
		If esc <> "" Then
			parts2 = Split(s, esc)
			If UBound(parts2) > 0 Then
				i = 0
				While i <= UBound(parts2)
					parts2(i) = Replace(parts2(i), sep, "MYSEPARATORBYPHILJONES")
					' nb : that's an unlikely string, but the function *will* fail if the argument contains it
					' "this record can not be played on record player B"
					i = i + 2
				End While
				s2 = Join(parts2, esc)
				parts = Split(s2, "MYSEPARATORBYPHILJONES")
			Else
				parts = Split(s, sep)
			End If
		Else
			parts = Split(s, sep)
		End If
		mySplit = VB6.CopyArray(parts)
	End Function
End Class