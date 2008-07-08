Option Strict Off
Option Explicit On
Friend Class NavigationHistory
	
	' where the user browsed, basically a pair of stacks :
	' past and future
	' when we go to a new page, add the last
	' the history
	' when we go back, pop it off the history and push it on the future
	' when we go forward again, pop it off the future and push it back on the history
	
	Private history As Collection
	Private future As Collection
	Private buildIndex As Short ' building up the history
	Private futureIndex As Short ' counts through future
	Private walker As Short
	Private lb As System.Windows.Forms.ComboBox
	
	Public Sub clear()
		walker = 0
		buildIndex = 0
		futureIndex = 0
		history = New Collection
		future = New Collection
	End Sub
	
	
	Public Sub setComboBox(ByRef cb As System.Windows.Forms.ComboBox)
		lb = cb
	End Sub
	
	Public Sub wipeFuture()
		future = New Collection
		futureIndex = 0
	End Sub
	
	Public Sub append(ByRef pageName As String)
		If buildIndex > 1 Then
			If pageName <> Me.getAtIndex() Then
				Call history.Add(pageName, CStr(buildIndex))
				buildIndex = buildIndex + 1
				Me.inspectInList()
			End If
		Else
			Call history.Add(pageName, CStr(buildIndex))
			buildIndex = buildIndex + 1
			Me.inspectInList()
		End If
	End Sub
	
	Public Function getAtIndex() As String
		Dim i As Short
		If buildIndex = 0 Then
			i = 0
		Else
			i = buildIndex - 1
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object history.Item(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		getAtIndex = history.Item(CStr(i))
	End Function
	
	
	
	Public Sub back()
		
		Dim s As String
		If buildIndex > 1 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object history.Item(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			s = history.Item(CStr(buildIndex - 1))
			Call future.Add(s, CStr(futureIndex))
			futureIndex = futureIndex + 1
			history.Remove((CStr(buildIndex - 1)))
			buildIndex = buildIndex - 1
		End If
		Me.inspectInList()
	End Sub
	
	
	Public Sub forward()
		Dim s As String
		If future.Count() > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object future.Item(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			s = future.Item(CStr(futureIndex - 1))
			Call future.Remove(CStr(futureIndex - 1))
			futureIndex = futureIndex - 1
			Call history.Add(s, CStr(buildIndex))
			buildIndex = buildIndex + 1
		End If
		Me.inspectInList()
	End Sub
    'The Whole function is commented due to print issue
	Public Sub printOn(ByRef c As System.Windows.Forms.Form)
        '' ''UPGRADE_ISSUE: Form method c.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        ' ''c.Print("History")
        ' ''Dim i As Short
        ' ''For i = 0 To history.Count() - 1
        ' ''	'UPGRADE_WARNING: Couldn't resolve default property of object history.Item(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ' ''	'UPGRADE_ISSUE: Form method c.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        ' ''	c.Print(CStr(i), history.Item(CStr(i)))
        ' ''Next i
        '' ''UPGRADE_ISSUE: Form method c.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        ' ''c.Print("Future")
        ' ''For i = 0 To future.Count() - 1
        ' ''	'UPGRADE_WARNING: Couldn't resolve default property of object future.Item(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ' ''	'UPGRADE_ISSUE: Form method c.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        ' ''	c.Print(CStr(i), future.Item(CStr(i)))
        ' ''Next i
	End Sub
	
	Public Function inspectToString() As String
		Dim s As String
		s = "History" & vbCrLf
		Dim i As Short
		For i = 0 To history.Count() - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object history.Item(CStr(i)). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			s = s & CStr(i) & "  " & history.Item(CStr(i)) & vbCrLf
		Next i
		s = s & "Future" & vbCrLf
		For i = 0 To future.Count() - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object future.Item(CStr(i)). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			s = s & CStr(i) & ", " & future.Item(CStr(i)) & vbCrLf
		Next i
		inspectToString = s
	End Function
	
	Public Function inspectInList() As Object
		Dim i As Short
		If Not lb Is Nothing Then
			Call lb.Items.Clear()
			Call lb.Items.Add("History")
			
			For i = 0 To history.Count() - 1
				'UPGRADE_WARNING: Couldn't resolve default property of object history.Item(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call lb.Items.Add(history.Item(CStr(i)))
			Next i
			Call lb.Items.Add("Future")
			For i = future.Count() - 1 To 0 Step -1
				'UPGRADE_WARNING: Couldn't resolve default property of object future.Item(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call lb.Items.Add(future.Item(CStr(i)))
			Next i
		End If
	End Function
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		Call clear()
		'UPGRADE_NOTE: Object lb may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		lb = Nothing
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object history may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		history = Nothing
		'UPGRADE_NOTE: Object future may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		future = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class