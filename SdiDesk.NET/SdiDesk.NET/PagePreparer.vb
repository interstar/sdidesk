Option Strict Off
Option Explicit On
Friend Class PagePreparer
	
	' this module now takes the pre-processing / preparation of a page from
	' the PageCooker.
	
	Public wads As _WikiAnnotatedDataStore ' needs to know about the model
	
	
	Public Function processInlines(ByRef raw As String, ByRef aPage As _Page) As String
		Dim lines() As String
		Dim l2 As Object
		Dim l As String
		Dim build, current As String ' what we build, and current state
		'UPGRADE_NOTE: ScriptEngine was upgraded to ScriptEngine_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim ScriptEngine_Renamed As New ScriptEngine_Renamed
		
		build = ""
		current = raw ' this time, at end of loop it will be set to build
		
		' we do this in a loop because it's recursive
		Dim finished As Boolean
		
		Dim parts() As String
		Do 
			build = ""
			finished = True
			lines = Split(current, vbCrLf)
			
			' now through each line of the page, substituting the inlines
			
			For	Each l2 In lines
				'UPGRADE_WARNING: Couldn't resolve default property of object l2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				l = CStr(l2) ' ensure it's a string
				
				' interpret #= assignments
				If InStr(l, "#=") > 0 Then
					parts = Split(l, "#=")
					Call aPage.setVal(parts(0), parts(1))
					l = vbCrLf & "<font size=+1 color=#339999> " & parts(0) & " #''''''= " & parts(1) & "</font>"
				End If
				
				' interpret $$ variables
				If InStr(l, "$$") > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object wads. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					l = ScriptEngine_Renamed.varsInLine(l, aPage, wads)
				End If
				
				' interpret ##Inlines
				If Left(l, 2) = "##" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object wads. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					l = ScriptEngine_Renamed.perform(l, wads)
					finished = False ' this is dirty, try it again
				End If
				
				build = build & l & vbCrLf
			Next l2
			
			' at this point we did all the inlines in the current state
			' but we want to make sure that if any inlines brought in NEW inlines
			' they get processed too.
			' so set current = build and maybe go round again
			
			current = build
		Loop Until finished = True ' loop until there are no more ##inlines left
		
		' now current should have all inlines fully substituted
		processInlines = current
		
		' clean up
		'UPGRADE_NOTE: Object ScriptEngine_Renamed may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		ScriptEngine_Renamed = Nothing
		
	End Function
	
	
	Public Function prepare(ByRef aPage As _Page, ByRef backlinks As Boolean) As String
		Dim s As String
		s = Me.processInlines((aPage.raw), aPage)
		
		Dim ps As PageSet
		If backlinks = True Then
			ps = wads.getPageSetContaining((aPage.pageName))
			s = s & "----" & vbCrLf & "<table bgcolor=#ccffcc width=100%><tr>" & vbCrLf
			s = s & "<td> <h3>Backlinks</h3> " & vbCrLf & ps.toWikiMarkup & "</td></tr></table>" & vbCrLf
		End If
		aPage.prepared = Me.processInlines(s, aPage)
		prepare = aPage.prepared
	End Function
End Class