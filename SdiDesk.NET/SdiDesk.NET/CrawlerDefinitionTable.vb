Option Strict Off
Option Explicit On
Friend Class CrawlerDefinitionTable
	
	' this table defines the crawlers in the system
	
	Public crawlerNames As VCollection
	Public crawlers As OCollection
	
	Private st As StringTool
	
	Private Sub clear()
		crawlerNames = New VCollection
		crawlers = New OCollection
	End Sub
	
	Public Sub parseFromTableString(ByRef tabDef As String, ByRef model As _ModelLevel, ByRef store As _PageStore, ByRef chef As Object)
		' we are expecting a crawler definition
		' name,, type,, maxDepth,, excluded pages,, excluded types
		Call clear()
		
		Dim i As Short
		Dim pc As _PageCrawler
		
		' parse a table string to a crawler
		Dim aTable As New Table
		aTable.parseFromDoubleCommaString(tabDef)
		
		'UPGRADE_NOTE: cType was upgraded to cType_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim exp, name, cType_Renamed, ext As String
		Dim md As Short
		
		For i = 0 To aTable.noRows - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object aTable.at(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			name = CStr(aTable.at(i, 0))
			If name <> "" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object aTable.at(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				cType_Renamed = st.strip(CStr(aTable.at(i, 1)))
				If (cType_Renamed = "recursive") Then
					'UPGRADE_WARNING: Couldn't resolve default property of object aTable.at(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					md = CShort(aTable.at(i, 2))
				Else
					md = 0
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object aTable.at(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				exp = CStr(aTable.at(i, 3))
				'UPGRADE_WARNING: Couldn't resolve default property of object aTable.at(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ext = CStr(aTable.at(i, 4))
                Try
                    pc = POLICY_getFactory().getNewPageCrawlerInstance(cType_Renamed, name, md, exp, ext)
                    Call crawlers.Add(pc, name)
                    Call crawlerNames.add(name, name)
                Catch ex As Exception

                End Try
				
			End If
		Next i
		
	End Sub
	
	Public Function getCrawler(ByRef name As String) As _PageCrawler
		If crawlers.hasKey(name) Then
			getCrawler = crawlers.Item(name)
		Else
			MsgBox("Error : no crawler called " & name)
		End If
	End Function
	
	'UPGRADE_NOTE: toString was upgraded to toString_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function toString_Renamed() As String
		Dim c As _PageCrawler
		Dim build As String
		build = ""
		For	Each c In Me.crawlers.toCollection
			build = build & c.toString_Renamed() & vbCrLf
		Next c
		toString_Renamed = build
	End Function
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		crawlerNames = New VCollection
		crawlers = New OCollection
		st = New StringTool
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object crawlerNames may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		crawlerNames = Nothing
		'UPGRADE_NOTE: Object crawlers may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		crawlers = Nothing
		'UPGRADE_NOTE: Object st may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		st = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class