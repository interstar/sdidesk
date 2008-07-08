Option Strict Off
Option Explicit On
Friend Class LinkTypeManager
	
	' stores the collection of linkTypes and colours
	
	Public c As VCollection ' will store in the format of key = typeName, val = colourDef
	
	
	Public Sub init()
		c = New VCollection
	End Sub
	
	
	Public Function setupLinkTypes(ByRef raw As String) As String
		' we're expecting raw to be in table format
		' be
		' typeName,, colourDef
		' typeName,, colourDef
		' etc.
		
		' reset the collection
		Call Me.init()
		
		' Put in a default colour which may be over-ridden by data from raw
		Call c.Add("#006633", "normal")
		
		' now parse the data from raw into a table, and then use to
		' fill the collection
		Dim t As New Table
		Call t.parseFromDoubleCommaString(raw)
		
		Dim st As New StringTool
		'UPGRADE_NOTE: typeName was upgraded to typeName_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim i As Short
		Dim typeName_Renamed, colDef As String
		For i = 0 To t.noRows - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object t.at(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			typeName_Renamed = st.strip(CStr(t.at(i, 0)))
			'UPGRADE_WARNING: Couldn't resolve default property of object t.at(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			colDef = st.strip(CStr(t.at(i, 1)))
			If c.hasKey(typeName_Renamed) Then
				c.Remove((typeName_Renamed))
			End If
			
			Call c.Add(colDef, typeName_Renamed)
		Next i
		'UPGRADE_NOTE: Object t may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		t = Nothing
		'UPGRADE_NOTE: Object st may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		st = Nothing
	End Function
	
	Public Function getColour(ByRef tn As String) As String
		If c.hasKey(tn) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object c.Item(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			getColour = c.Item(tn)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object c.Item(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			getColour = c.Item("normal")
		End If
	End Function
	
	'UPGRADE_NOTE: toString was upgraded to toString_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function toString_Renamed() As String
		toString_Renamed = c.toString_Renamed()
	End Function
	
	Public Function toStyleDefs() As String
		MsgBox("LinkTypeManager:toStyleDefs unwritten")
		'  For Each s2 In d
		'    Dim s As String
		'    s = CStr(s2)
		'    If s <> "" Then
		'      Dim parts() As String
		'      parts = Split(s, ",, ")
		'      build = build + "." + parts(0) + " {color:" + parts(1) + "}" + vbCrLf
		'
		
		'    End If
		'  Next s2
		
		'  build = build + "</style>" + vbCrLf + "</head>" + vbCrLf
		'  styleSheet = build ' may need to change this later
		'  setupLinkTypes = build
		toStyleDefs = ""
	End Function
End Class