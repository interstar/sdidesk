Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class TimeIndex
	
	' The purpose of the TimeIndex is to manage a time-based index
	' of all pages.
	' This is stored in a set of files named after months, in
	' directories named years
	' ie. 2004\Febuary.mnp
	
	' The format of the month is just a double comma separated
	' list ... day number,, page
	
	' Gets read into this time-based index
	
	' ----------------------
	
	Private monthNames As Object
	
	Dim oneMonth As VCollection
	Dim store As _PageStore
	
	Public Sub init(ByRef ps As _PageStore)
		store = ps
	End Sub
	
	'UPGRADE_NOTE: month was upgraded to month_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: year was upgraded to year_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub loadMonthFromStore(ByRef year_Renamed As Short, ByRef month_Renamed As Short)
		Dim s As String
		s = store.loadMonth(month_Renamed, year_Renamed)
		Call parseMonthFromString(s)
	End Sub
	
	Public Sub loadMonthByDate(ByRef d As Date)
        Call loadMonthFromStore(d.Year, d.Month)
	End Sub
	
	Public Sub parseMonthFromString(ByRef s As String)
		' expected format is
		' day,, PageName
		' day,, PageName etc.
		' may have two days the same, not a problem
		oneMonth = New VCollection
		Dim lines() As String
		lines = Split(s, vbCrLf)
		Dim i As Short
		Dim l As Object
		i = 0
		For	Each l In lines
			'UPGRADE_WARNING: Couldn't resolve default property of object l. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If CStr(l) <> "" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object l. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call oneMonth.Add(CStr(l), CStr(l))
			End If
		Next l
	End Sub
	
	
	'UPGRADE_NOTE: month was upgraded to month_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: year was upgraded to year_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub saveMonth(ByRef year_Renamed As Short, ByRef month_Renamed As Short)
		Call store.saveMonth(month_Renamed, year_Renamed, Me.toRawString)
	End Sub
	
	
	Public Sub saveMonthByDate(ByRef d As Date)
        Call saveMonth(CShort(d.Year), CShort(d.Month))
	End Sub
	
	Public Function oneDayToWikiString(ByRef dayNum As Short) As String
		Dim build As String
		build = ""
		Dim r As Object
		Dim parts() As String
		For	Each r In oneMonth.toCollection
			' NB : stepping through record entries, NOT days
			'UPGRADE_WARNING: Couldn't resolve default property of object r. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If CStr(r) <> "" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object r. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				parts = Split(CStr(r), ",, ")
				If CShort(parts(0)) = dayNum Then
					build = build & "** [[" & parts(1) & "]]" & vbCrLf '+ ", "
				End If
			End If
		Next r
		oneDayToWikiString = build
	End Function
	
	'UPGRADE_NOTE: monthName was upgraded to monthName_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: month was upgraded to month_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function monthName_Renamed(ByRef month_Renamed As Short) As String
		'UPGRADE_WARNING: Couldn't resolve default property of object monthNames(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		monthName_Renamed = monthNames(month_Renamed - 1)
	End Function
	
	Public Function toDate(ByRef aYear As Short, ByRef aMonth As Short, ByRef aDay As Short) As String
		Dim cm As New CalendarMonth
		toDate = "[[" & aDay & "-" & cm.monthName_Renamed(aMonth - 1) & "-" & aYear & "]]"
		'UPGRADE_NOTE: Object cm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		cm = Nothing
	End Function
	
	'UPGRADE_NOTE: year was upgraded to year_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: month was upgraded to month_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function toWikiString(ByRef month_Renamed As Short, ByRef year_Renamed As Short, ByRef order As Boolean) As String
		' order = true = 1->31, order=false = 31 -> 1
		Dim build, ds As String
		Dim i As Short
		
		build = ""
		If order = True Then
			For i = 1 To 31
				ds = oneDayToWikiString(i)
				If ds <> "" Then
					build = build & "* " & toDate(year_Renamed, month_Renamed, i) & " : " & vbCrLf & ds & vbCrLf
				End If
			Next i
		Else
			For i = 31 To 1 Step -1
				ds = oneDayToWikiString(i)
				If ds <> "" Then
					build = build & "* " & toDate(year_Renamed, month_Renamed, i) & " : " & vbCrLf & ds & vbCrLf
				End If
			Next i
		End If
		toWikiString = build
	End Function
	
	Public Function toRawString() As String
		Dim build As String
		Dim i As Object
		build = ""
		For	Each i In oneMonth.toCollection
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			build = build & CStr(i) & vbCrLf
		Next i
		toRawString = build
	End Function
	
	Public Sub addWord(ByRef word As String, ByRef dayNum As Short)
		Dim i As Object
		Dim flag As Boolean
		Dim c As New VCollection
		Dim s As String
		Dim dummy As Short
		
		flag = False
		dummy = 0
		
		Dim parts() As String
		For	Each i In oneMonth.toCollection
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			parts = Split(CStr(i), ",, ")
			If CShort(parts(0)) < dayNum Then
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call c.Add(CStr(i), CStr(dummy))
				dummy = dummy + 1
			Else
				If parts(0) = CStr(dayNum) And parts(1) = word Then
					' do nothing
				Else
					
					If flag = False Then
						' first on this date
						Call c.Add(dayNum & ",, " & word, CStr(dummy))
						'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Call c.Add(CStr(i), CStr(dummy + 1))
						flag = True
						dummy = dummy + 2
					Else
						' now rest
						'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Call c.Add(CStr(i), CStr(dummy))
						dummy = dummy + 1
					End If
				End If
			End If
		Next i
		If flag = False Then
			Call c.Add(dayNum & ",, " & word, CStr(dummy))
		End If
		oneMonth = c
	End Sub
	
	Public Sub removeWord(ByRef word As String)
		Dim i As Object
		Dim c As New VCollection
		Dim parts() As String
		For	Each i In oneMonth.toCollection
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			parts = Split(CStr(i), ",, ")
			If parts(1) <> word Then
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call c.Add(CStr(i), CStr(i))
			End If
		Next i
		oneMonth = c
	End Sub
	
	Public Sub updateWord(ByRef word As String, ByRef oldDate As Date, ByRef newDate As Date)
		' OK, this is what we actually call
		' we take the word out of the oldDate and put it into newDate
		' load old date
		Me.loadMonthByDate((oldDate))
		' remove word
		Me.removeWord((word))
		' save old date
		Me.saveMonthByDate((oldDate))
		' load new date
		Me.loadMonthByDate((newDate))
		' add word
		Call Me.addWord(word, VB.Day(newDate))
		' save new date
		Me.saveMonthByDate((newDate))
	End Sub
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object monthNames. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		monthNames = New Object(){"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"}
		oneMonth = New VCollection
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object oneMonth may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oneMonth = Nothing
		ReDim monthNames(0)
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class