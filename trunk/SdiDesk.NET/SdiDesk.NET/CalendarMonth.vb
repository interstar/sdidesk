Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class CalendarMonth
	
	' Calendar class month
	
	Dim monthNames As Object
	Dim noDays(12) As Short
	
	
	
	'UPGRADE_NOTE: month was upgraded to month_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: year was upgraded to year_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function getFirst(ByRef year_Renamed As Short, ByRef month_Renamed As Short) As Date
		getFirst = CDate(year_Renamed & "-" & month_Renamed & "-1")
	End Function
	
	'UPGRADE_NOTE: day was upgraded to day_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: month was upgraded to month_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: year was upgraded to year_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function datePageName(ByRef year_Renamed As Short, ByRef month_Renamed As Short, ByRef day_Renamed As Short) As String
		datePageName = day_Renamed & "-" & monthName_Renamed(month_Renamed - 1) & "-" & year_Renamed
	End Function
	
	'UPGRADE_NOTE: monthName was upgraded to monthName_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function monthName_Renamed(ByRef m As Short) As String
		'UPGRADE_WARNING: Couldn't resolve default property of object monthNames(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		monthName_Renamed = CStr(monthNames(m))
	End Function
	
	'UPGRADE_NOTE: day was upgraded to day_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Function nativeLink(ByRef link As String, ByRef day_Renamed As Short) As String
		nativeLink = "[[calendar>" & link & "|" & day_Renamed & "]]"
	End Function
	
	
	'UPGRADE_NOTE: month was upgraded to month_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: year was upgraded to year_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function monthAsString(ByRef year_Renamed As Short, ByRef month_Renamed As Short, ByRef include As Boolean, ByRef model As _ModelLevel) As String
		Dim first As Short
		first = WeekDay(getFirst(year_Renamed, month_Renamed))
		
		Dim cal(5, 7) As Short
		Dim j, i, dCount As Short
		Dim s, dd As String
		Dim nL As String
		dCount = 1
		
		'UPGRADE_WARNING: Couldn't resolve default property of object monthNames(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		s = "== " & monthNames(month_Renamed - 1) & ", " & year_Renamed & " ==" & vbCrLf
		s = s & "<font size='2'><table border='1'>"
		
		If include = False Then
			s = s & "<tr bgcolor='#eeffaa'><td>S</td><td>M</td><td>Tu</td><td>We</td><td>Th</td><td>F</td><td>S</td></tr>" & vbCrLf
		Else
			s = s & "<tr bgcolor='#eeffaa'><td>Sunday</td><td>Monday</td><td>Tuesday</td><td>Wedenesday</td><td>Thursday</td><td>Friday</td><td>Saturday</td></tr>" & vbCrLf
		End If
		j = 0
		Do While j < noDays(month_Renamed) + 1
			j = j + 1
			s = s & "<tr>"
			For i = 1 To 7
				If dCount > 1 Or i >= first Then
					nL = nativeLink(datePageName(year_Renamed, month_Renamed, dCount), dCount)
					dd = model.getWikiAnnotatedDataStore.getRawPageData(datePageName(year_Renamed, month_Renamed, dCount))
					
					If dd = "new page" Then
						s = s & "<td bgcolor='#eeeeee'>"
						dd = ""
					Else
						s = s & "<td>"
					End If
					s = s & nL
					If include Then
						s = s & "<br>" & Left(dd, 25)
					End If
					s = s & "</td>" & vbCrLf
					dCount = dCount + 1
					
				Else
					s = s & "<td bgcolor='#999999'></td>" & vbCrLf
					
				End If
				
				If dCount > noDays(month_Renamed) Then
					Exit Do
				End If
				
			Next i
			s = s & "</tr>" & vbCrLf
		Loop 
		s = s & vbCrLf & "</table></font>" & vbCrLf
		
		monthAsString = s
	End Function
	
	
	Public Function includedDay(ByRef p As String, ByRef store As _PageStore) As String
		includedDay = "BOX<" & vbCrLf & "===" & p & "===" & vbCrLf & "##Include " & p & vbCrLf & ">BOX" & vbCrLf & "----" & vbCrLf
		
	End Function
	
	Public Function includeAllBetween(ByRef year1 As Short, ByRef month1 As Short, ByRef year2 As Short, ByRef month2 As Short, ByRef direction As Short, ByRef store As _PageStore) As String
		Dim d As Date
		Dim s, p As String
		s = ""
		If direction = 1 Then
			d = getFirst(year1, month1)
			While d < getFirst(year2, month2)
				p = datePageName(Year(d), Month(d), VB.Day(d))
				If store.pageExists(p) Then
					s = s & includedDay(p, store)
				End If
				d = System.Date.FromOADate(d.ToOADate + 1)
			End While
		Else
			d = getFirst(year2, month2)
			While d >= getFirst(year1, month1)
				p = datePageName(Year(d), Month(d), VB.Day(d))
				If store.pageExists(p) Then
					s = s & includedDay(p, store)
				End If
				d = System.Date.FromOADate(d.ToOADate - 1)
			End While
		End If
		includeAllBetween = s
	End Function
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		noDays(1) = 31
		noDays(2) = 28
		noDays(3) = 31
		noDays(4) = 30
		noDays(5) = 31
		noDays(6) = 30
		noDays(7) = 31
		noDays(8) = 31
		noDays(9) = 30
		noDays(10) = 31
		noDays(11) = 30
		noDays(12) = 31
		
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object monthNames. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		monthNames = New Object(){"Jan", "Feb", "Mar", "Apr", "May", "June", "July", "Aug", "Sep", "Oct", "Nov", "Dec"}
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
End Class