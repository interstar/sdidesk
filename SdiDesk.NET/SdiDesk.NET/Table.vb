Option Strict Off
Option Explicit On
Friend Class Table
	
	' table class (includes column names)
	' and can do some tricks like totalling numeric columns etc
	
	Private headers() As String
    Private body(0, 0) As Object
	Private sums() As String
	Private means() As String
	
	Public noRows As Short
	Public noCols As Short
	Public comment As String
	
	Private st As StringTool
	
	Private hasHeaders As Boolean
	
	Public Sub setUp(ByRef r As Short, ByRef c As Short)
		noRows = r
		noCols = c
		ReDim headers(c + 1)
		ReDim body(r + 1, c + 1)
		ReDim sums(c + 1)
		ReDim means(c + 1)
		hasHeaders = False
	End Sub
	
	Public Sub putIn(ByRef r As Short, ByRef c As Short, ByRef v As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object body(r, c). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		body(r, c) = v
	End Sub
	
	Public Function at(ByRef r As Short, ByRef c As Short) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object body(r, c). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		at = body(r, c)
	End Function
	
	Public Sub setHeader(ByRef c As Short, ByRef h As String)
		headers(c) = h
	End Sub
	
	Public Function atHeader(ByRef c As Short) As String
		atHeader = headers(c)
	End Function
	
	Public Function isValidTable(ByRef s As String) As Boolean
		Dim b As Boolean
		Dim t As New Table
		isValidTable = t.parseFromDoubleCommaString(s)
		'UPGRADE_NOTE: Object t may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		t = Nothing
	End Function
	
	Public Function parseFromDoubleCommaString(ByRef t2 As String) As Boolean
		
		' format is like this
		
		' head,, head,, head
		' ____
		' body,, body,, body
		' body,, body,, body
		'
		' optional comments
		
		' header is optional, and inferred from the following ____ line
		' note the comments must be separated from the end by at least
		' one blank line
		
		' returns true if succesful, false if not
		
		Dim lines() As String
		Dim parts() As String
		Dim t As String
		Dim success As Boolean
		Dim i As Short
		
		t = t2 ' makes sure what we're processing isn't the argument
		
		If InStr(t, vbCrLf & vbCrLf) > 0 Then
			' strip off the comment at the bottom
			i = InStr(t, vbCrLf & vbCrLf)
			
			comment = Right(t, (Len(t) - i) + 1)
			comment = st.strip(comment)
			comment = st.trimLeft(comment)
			comment = st.trimLeft(comment)
			
			t = st.strip(Left(t, i))
			
		End If
		
		Dim startRow As Short
		Dim rowCount As Short
		Dim j As Short
		If InStr(t, vbCrLf) Then
			
			lines = Split(t, vbCrLf)
			
			' first guess at number of rows
			' though we'll correct if there's a header
			
			noRows = UBound(lines) + 1
			
			' now see if the first line contains headers by seeing if the
			' second line is composed of ____
			
			startRow = 0
			rowCount = 0
			
			
			If InStr(CStr(lines(1)), "____") > 0 Then
				' has headers in line 0
				hasHeaders = True
				
				parts = Split(st.strip(CStr(lines(0))), ",,")
				noCols = UBound(parts) + 1
				startRow = 2 ' skip past header lines
				noRows = noRows - 2 ' lose headers and ====
				
				' now we know enough to redim the arrays
				Call setUp(noRows, noCols)
				
				' now we can fill the headers
				For i = 0 To noCols - 1
					headers(i) = st.strip(CStr(parts(i)))
				Next i
				
			Else
				startRow = 0 ' no header, so start from top
				' but still must count cols
				parts = Split(CStr(lines(0)), ",,")
				noCols = UBound(parts) + 1
				Call setUp(noRows, noCols) ' and redim the arrays
			End If
			
			
			' now let's read the table body
			For i = 0 To noRows - 1
				On Error GoTo failedRow
				
				parts = Split(CStr(lines(i + startRow)), ",,")
				For j = 0 To noCols - 1
					If j <= UBound(parts) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object body(i, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						body(i, j) = st.strip(CStr(parts(j)))
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object body(i, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						body(i, j) = ""
					End If
				Next j
				rowCount = rowCount + 1
				
failedRow: 
			Next i
			
			success = True
		Else
			success = False
		End If
		
		parseFromDoubleCommaString = success
	End Function
	
	
	
	'UPGRADE_NOTE: isNumeric was upgraded to isNumeric_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function isNumeric_Renamed(ByRef col As Short) As Boolean
		' returns true if the column only contains numbers
		Dim i As Short
		Dim flag As Boolean
		flag = False
		On Error GoTo notNumeric
		Dim d As Double
		For i = 0 To noRows
			'UPGRADE_WARNING: Couldn't resolve default property of object at(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			d = CDbl(at(i, col))
		Next i
		' if we got here, all in column could be
		' turned into double, ie. were numeric,
		' so
		flag = True
		
notNumeric: 
		isNumeric_Renamed = flag
	End Function
	
	Public Function allNumeric() As Boolean
		Dim i As Short
		Dim flag As Boolean
		flag = True
		For i = 0 To noCols - 1
			If isNumeric_Renamed(i) Then
				flag = False
			End If
		Next i
		allNumeric = flag
	End Function
	
	Public Function calc() As Object
		Dim i, j As Short
		Dim t As Double
		For j = 0 To noCols - 1
			If isNumeric_Renamed(j) Then
				t = 0
				For i = 0 To noRows
					'UPGRADE_WARNING: Couldn't resolve default property of object at(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					t = t + CDbl(at(i, j))
				Next i
				sums(j) = CStr(t)
				means(j) = CStr(t / noRows)
			Else
				sums(j) = ""
				means(j) = ""
			End If
		Next j
		
	End Function
	
	Public Function rows() As Short
		rows = noRows
	End Function
	
	
	Public Sub project(ByRef t As Table, ByRef query As String)
		' this table becomes a copy of some cols from another table
		' query = "colNo colNo colNo"
		Dim parts() As String
		parts = Split(query, " ")
		Dim c As Short
		c = UBound(parts) + 1
		
		' dimension self as appropriate
		Call setUp(t.rows, c)
		
		Dim p As Object
		Dim cn As Short
		Dim cc As Short
		cc = 0
		On Error GoTo endOfQueryLine
		Dim i As Short
		For	Each p In parts
			'UPGRADE_WARNING: Couldn't resolve default property of object p. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			cn = CShort(p)
			headers(cc) = t.atHeader(cn)
			For i = 0 To noRows
				Call putIn(i, cc, t.at(i, cn))
			Next i
			cc = cc + 1
		Next p
		
endOfQueryLine: 
		' got here when we ran out of columns
		' is this a good way of handling error?
		
	End Sub
	
	Public Function toWikiFormat() As Object
		Dim i As Short
		Dim j As Short
		Dim s As String
		Call calc()
		s = " ,,"
		For j = 0 To noCols - 1
			s = s & "'''" & CStr(headers(j)) & "''',, "
		Next j
		s = s & vbCrLf & "____" & vbCrLf
		For i = 0 To noRows - 1
			s = s & " ,,"
			For j = 0 To noCols - 1
				'UPGRADE_WARNING: Couldn't resolve default property of object body(i, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If body(i, j) < 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object body(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					s = s & "<font color=#660000>" + body(i, j) + "</font>,, "
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object body(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					s = s + body(i, j) + ",, "
				End If
			Next j
			s = s & vbCrLf
		Next i
		s = s & "tot,, "
		For j = 0 To noCols - 1
			s = s & "<font color=#009900>" & sums(j) & "</font>,, "
		Next j
		s = s & vbCrLf & "av.,, "
		For j = 0 To noCols - 1
			s = s & "<font color=#000099>" & means(j) & "</font>,, "
		Next j
		
		
		'UPGRADE_WARNING: Couldn't resolve default property of object toWikiFormat. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		toWikiFormat = s
	End Function
	
	
	Public Sub inspect()
		MsgBox(toWikiFormat())
	End Sub
	
	
	Public Function spitAsPrettyPersist() As String
		Dim i, j As Short
		Dim s As String
		
		s = ""
		
		For j = 0 To noCols
			If st.strip(CStr(atHeader(j))) <> "" Then
				s = s & atHeader(j) & ",, "
			End If
		Next j
		
		s = st.stripRight(s, " ")
		If Right(s, 2) = ",," Then
			s = st.trimRight(s)
			s = st.trimRight(s)
		End If
		s = s & vbCrLf & "____" & vbCrLf
		
		For i = 0 To noRows
			For j = 0 To noCols
				'UPGRADE_WARNING: Couldn't resolve default property of object at(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If st.strip(CStr(at(i, j))) <> "" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object at(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					s = s & at(i, j) & ",, "
				End If
			Next j
			
			s = st.stripRight(s, " ")
			If Right(s, 2) = ",," Then
				s = st.trimRight(s)
				s = st.trimRight(s)
				s = st.stripRight(s, " ")
			End If
			s = s & vbCrLf
		Next i
		
		s = st.strip(s)
		
		s = s & st.stripLeft(comment, vbCrLf)
		
		spitAsPrettyPersist = s
	End Function
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		st = New StringTool
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object st may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		st = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class