Option Strict Off
Option Explicit On
'UPGRADE_NOTE: ScriptEngine was upgraded to ScriptEngine_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
Friend Class ScriptEngine_Renamed
	
	' performs some of the more complex stuff which
	' you can do with ##Inlines
	' not much of a scripting language yet, but
	' the name represents the aspiration of where we're going ;-)
	
	Private ti As TimeIndex
	
	Public Function perform(ByRef line As String, ByRef model As _ModelLevel) As String
		'UPGRADE_NOTE: command was upgraded to command_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim command_Renamed, argList As String
		Dim st As New StringTool
		
		If model Is Nothing Then
			MsgBox("perform, model is nothing")
		End If
		
		command_Renamed = st.strip(st.leftsa(line, " ", 1)) ' grab the first left word of line as command
		argList = st.stripHead(line, " ", 1) ' rest of the line
		
		Dim tokens() As String
		tokens = Split(st.strip(argList), ",, ")
		
		Select Case command_Renamed
			Case "##AllPages"
				perform = AllPages(model)
			Case "##Include"
				If UBound(tokens) > -1 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object model. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					perform = include(st.strip(tokens(0)), model)
				Else
					perform = "<font color='red'>Error in ##Include " & argList
				End If
			Case "##TableInclude"
				perform = tableInclude(tokens, model)
			Case "##Image"
				perform = imageInclude(tokens, model)
			Case "##Month"
				perform = monthChanges(tokens, model)
			Case "##CalendarMonth"
				'UPGRADE_WARNING: Couldn't resolve default property of object CalendarMonth(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				perform = CalendarMonth(tokens, model)
			Case "##IncludingCalendarMonth"
				'UPGRADE_WARNING: Couldn't resolve default property of object IncludingCalendarMonth(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				perform = IncludingCalendarMonth(tokens, model)
			Case "##CalendarEntries"
				'UPGRADE_WARNING: Couldn't resolve default property of object CalendarEntries(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				perform = CalendarEntries(tokens, model)
			Case "##RecentChanges"
				perform = recentChanges(model)
			Case "##Find"
				perform = find(argList, model)
			Case "##Local"
				If UBound(tokens) > 0 Then
					perform = localResource(st.strip(tokens(0)), st.strip(tokens(1)), model)
				Else
					perform = localResource(st.strip(tokens(0)), st.strip(tokens(0)), model)
				End If
			Case "##Dir"
				perform = Me.localDir(argList, model)
			Case "##WordCount"
				'UPGRADE_WARNING: Couldn't resolve default property of object model.getControllableModel(tokens(0)). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                '	perform = "Word count of " & tokens(0) & " is " & model.getControllableModel(tokens(0""))
                MessageBox.Show("reached at error posiont 1")
			Case "##Button"
				If UBound(tokens) > 0 Then
					perform = makeButton(tokens(0), tokens(1))
				Else
					perform = err_Renamed("Sorry, bad arguments for Button")
				End If
			Case Else
				perform = err_Renamed("Sorry, don't know how to '" & line & "'")
		End Select
	End Function
	
	
	'UPGRADE_NOTE: err was upgraded to err_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function err_Renamed(ByRef s As String) As String
		err_Renamed = "<font color=##ff0000 size=+1>" & s & "</font>" & vbCrLf
	End Function
	
	Public Function AllPages(ByRef model As _ModelLevel) As String
		Dim ps As PageSet
		Dim returnVal As String
		returnVal = err_Renamed("Sorry, problem getting all pages")
		On Error GoTo err1
		ps = model.getWikiAnnotatedDataStore.store.AllPages()
		Dim s As String
		s = CStr(ps.size()) & " pages " & vbCrLf
		returnVal = s & ps.toWikiMarkup
		
err1: 
		
		AllPages = returnVal
		
	End Function
	
	
	Public Function include(ByRef otherPage As String, ByRef wads As _WikiAnnotatedDataStore) As String
		include = wads.getRawPageData(otherPage)
	End Function
	
	
	Public Function tableInclude(ByRef tokens() As String, ByRef model As _ModelLevel) As String
		
		Dim tableRaw, returnValue As String
		
		returnValue = err_Renamed("Sorry, error including this table")
		On Error GoTo err1
		
		' get the raw page data
		tableRaw = model.getWikiAnnotatedDataStore.getRawPageData(tokens(0))
		
		returnValue = err_Renamed("Sorry, problem TableIncluding " & tokens(0))
		
		Dim t As New Table
		Dim t2 As New Table
		' parse page data into table
		
		Dim at As New ArrayTool
		If t.isValidTable(tableRaw) Then
			
			Call t.parseFromDoubleCommaString(tableRaw)
			
			If UBound(tokens) = 0 Then
				' this has no further arguments, show whole thing
				t2 = t
			Else
				' collect arguments 2 +
				Dim colIndexes(UBound(tokens)) As String
				Call at.copyStringArray(tokens, colIndexes, 1, UBound(tokens), 0)
				
				' now get columns out
				Call t2.project(t, Join(colIndexes, " "))
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object t2.toWikiFormat. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			returnValue = t2.toWikiFormat
		Else
			returnValue = err_Renamed("Error trying to TableInclude [[#raw " & tokens(0) & "]]")
		End If
		
err1: 
		tableInclude = returnValue
	End Function
	
	Public Function qq(ByRef s As String) As String
		qq = Chr(34) & s & Chr(34)
	End Function
	
	Public Function imageInclude(ByRef tokens() As String, ByRef model As _ModelLevel) As String
		Dim d, linkTarget, imageName, returnValue, iTag As String
		returnValue = err_Renamed("Sorry, couldn't include image, no name")
		On Error GoTo err1
		imageName = tokens(0)
		
		returnValue = err_Renamed("Sorry, couldn't include image called " & imageName)
		
		d = model.getWikiAnnotatedDataStore.store.pictureLocality & imageName
		iTag = "<img src='" & d & "' border=0>"
		
		' if second argument, it's a link
		Dim around As String
		If UBound(tokens) = 1 Then
			If tokens(1) <> "" Then ' make sure they didn't just leave a trailing commas
				around = ""
				around = around & "<a href=" & qq("about:blank") & " id=" & qq(tokens(1)) & ">" & iTag & "</a>"
				iTag = around
			End If
		End If
		
		' add a NoWiki so that picture-paths don't get interpretted as wiki-words
		iTag = "#NoWiki" & vbCrLf & iTag & vbCrLf & "#Wiki" & vbCrLf
		returnValue = iTag
err1: 
		
		imageInclude = returnValue
	End Function
	
	
	
	Public Function monthChanges(ByRef tokens() As String, ByRef model As _ModelLevel) As String
		Dim s, returnValue As String
		returnValue = err_Renamed("Sorry, couldn't show a month. Did you specify month and year correctly? ")
		On Error GoTo err1
		s = "=== " & ti.monthName_Renamed(CShort(tokens(0))) & ", " & tokens(1) & " === " & vbCrLf
		s = s & ti.monthName_Renamed(CDbl(tokens(0)) - 1) & " : "
		s = s & ti.monthName_Renamed(CDbl(tokens(0)) + 1) & vbCrLf
		'UPGRADE_WARNING: Couldn't resolve default property of object model.getWikiAnnotatedDataStore.store.timeIndexAsWikiFormat(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		returnValue = s + model.getWikiAnnotatedDataStore.store.timeIndexAsWikiFormat(CShort(tokens(0)), CShort(tokens(1)), False)
err1: 
		monthChanges = returnValue
	End Function
	
	Public Function CalendarMonth(ByRef tokens() As String, ByRef model As _ModelLevel) As Object
		Dim s, returnValue As String
		Dim cm As New CalendarMonth
		returnValue = err_Renamed("Sorry, couldn't show a month. Did you specify month and year correctly? ")
		On Error GoTo err1
		returnValue = cm.monthAsString(CShort(tokens(0)), CShort(tokens(1)), False, model)
err1: 
		'UPGRADE_WARNING: Couldn't resolve default property of object CalendarMonth. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CalendarMonth = returnValue
	End Function
	
	Public Function IncludingCalendarMonth(ByRef tokens() As String, ByRef model As _ModelLevel) As Object
		Dim s, returnValue As String
		Dim cm As New CalendarMonth
		returnValue = err_Renamed("Sorry, couldn't show a month. Did you specify month and year correctly? ")
		On Error GoTo err1
		returnValue = cm.monthAsString(CShort(tokens(0)), CShort(tokens(1)), True, model)
err1: 
		'UPGRADE_WARNING: Couldn't resolve default property of object IncludingCalendarMonth. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		IncludingCalendarMonth = returnValue
	End Function
	
	Public Function CalendarEntries(ByRef tokens() As String, ByRef model As _ModelLevel) As Object
		Dim s, returnValue As String
		Dim cm As New CalendarMonth
		returnValue = err_Renamed("Sorry, couldn't show these entries. Did you specify months and years correctly? ")
		On Error GoTo err1
		If UBound(tokens) > 2 Then
			If UBound(tokens) > 3 Then
				If tokens(4) = "back" Then
					returnValue = cm.includeAllBetween(CShort(tokens(0)), CShort(tokens(1)), CShort(tokens(2)), CShort(tokens(3)), 0, model.getWikiAnnotatedDataStore.store)
				Else
					returnValue = cm.includeAllBetween(CShort(tokens(0)), CShort(tokens(1)), CShort(tokens(2)), CShort(tokens(3)), 1, model.getWikiAnnotatedDataStore.store)
				End If
			Else
				returnValue = cm.includeAllBetween(CShort(tokens(0)), CShort(tokens(1)), CShort(tokens(2)), CShort(tokens(3)), 1, model.getWikiAnnotatedDataStore.store)
			End If
		End If
err1: 
		'UPGRADE_WARNING: Couldn't resolve default property of object CalendarEntries. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CalendarEntries = returnValue
	End Function
	
	
	Public Function recentChanges(ByRef model As _ModelLevel) As String
		Dim d As Date
		Dim monthNumber, yearNumber As Short
		Dim mArg As String
        d = DateTime.Now
        monthNumber = d.Month ' Month(d)
        yearNumber = d.Year ' Year(d)
		
		Dim m, s As String
		m = ti.monthName_Renamed(monthNumber)
		
		s = "=== " & m & ", " & yearNumber & " === " & vbCrLf
		'UPGRADE_WARNING: Couldn't resolve default property of object model.getWikiAnnotatedDataStore.store.timeIndexAsWikiFormat(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		s = s + model.getWikiAnnotatedDataStore.store.timeIndexAsWikiFormat(monthNumber, yearNumber, False)
		
		' now one month earlier
		monthNumber = monthNumber - 1
		
		If monthNumber < 1 Then ' wrap around for before January
			monthNumber = 12
			yearNumber = yearNumber - 1
		End If
		
		mArg = CStr(monthNumber)
		
		m = ti.monthName_Renamed(monthNumber)
		s = s & vbCrLf & "=== " & m & "-" & yearNumber & " === " & vbCrLf
		'UPGRADE_WARNING: Couldn't resolve default property of object model.getWikiAnnotatedDataStore.store.timeIndexAsWikiFormat(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		s = s + model.getWikiAnnotatedDataStore.store.timeIndexAsWikiFormat(monthNumber, yearNumber, False)
		
		recentChanges = s
	End Function
	
	Public Function find(ByRef searchString As String, ByRef model As _ModelLevel) As String
		Dim ps As PageSet
		ps = model.getWikiAnnotatedDataStore.store.getPageSetContaining(searchString)
		find = ps.toWikiMarkup
		'UPGRADE_NOTE: Object ps may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		ps = Nothing
	End Function
	
	
	Public Function localResource(ByRef linkText As String, ByRef path As String, ByRef model As _ModelLevel) As String
		Dim build As String
		build = "#NoWiki" & vbCrLf
		build = build & "<a target= 'new' id='external' href='file:///" & path & "'>"
		build = build & linkText & "</a>" & vbCrLf & "#Wiki" & vbCrLf
		localResource = build
	End Function
	
	Public Function localDir(ByRef path As String, ByRef model As _ModelLevel) As String
		localDir = model.getLocalFileSystem.makeDirectoryPage(path).raw
	End Function
	
	
	Function varsInLine(ByRef l As String, ByRef p As _Page, ByRef model As _ModelLevel) As String
		Dim parts() As String
		Dim build As String
		Dim varName, pName As String
		
		build = ""
		parts = Split(l, " ")
		Dim v As Object
		Dim s As String
		Dim subParts() As String
		For	Each v In parts
			'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			s = CStr(v)
			If Left(s, 2) = "$$" Then
				' it's a variable, either remote ie. $$PageName:VarName
				' or local ie. $$VarName
				
				If InStr(s, ":") Then
					subParts = Split(s, ":")
					
					pName = Right(subParts(0), Len(subParts(0)) - 2)
					varName = subParts(1)
					
					build = build & " " & model.getWikiAnnotatedDataStore.getPageVar(pName, varName)
				Else
					varName = Right(s, Len(s) - 2)
					build = build & " " & p.getVal(varName)
					' not a var include after all, just some random two $$ thing
				End If
				
			Else
				build = build & " " & s
			End If
			
		Next v
		
		varsInLine = build
	End Function
	
	
	Public Function makeButton(ByRef destination As Object, ByRef text As Object) As String
		Dim x As String
		'UPGRADE_WARNING: Couldn't resolve default property of object text. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object destination. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		x = "#NoWiki" & vbCrLf & "<form action='' method=get><a href='about:blank' id='" + destination + "'>" + "<input type='button' value='" + text + "'></a></form>" + vbCrLf + "#Wiki"
		'MsgBox (x)
		makeButton = x
	End Function
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		ti = New TimeIndex
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object ti may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		ti = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class