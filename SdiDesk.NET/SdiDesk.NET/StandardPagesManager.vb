Option Strict Off
Option Explicit On
Friend Class StandardPagesManager
	
	' this object just encapsulates the responsibility of ensuring
	' a set of default pages to kick off your new SdiDesk
	' it makes RecentChanges, AllPages, some LinkTypeDefinitions,
	' CrawlerDefinitions, ExportDefinitions and a default BasicHtmlTemplate
	
	Public Function ensurePage(ByRef store As _PageStore, ByRef pageName As String, ByRef defaultRaw As String) As Object
		Dim p As _Page

		
		If Not store.pageExists(pageName) Then
            p = POLICY_getFactory().getNewPageInstance
			p.pageName = pageName
			p.raw = defaultRaw
			
			Call store.savePage(p)
		End If
		'UPGRADE_NOTE: Object p may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		p = Nothing
	End Function
	
	Public Function ensureStandardPages(ByRef store As _PageStore) As Object
		Dim s As String
		Call ensurePage(store, "RecentChanges", "##RecentChanges")
		Call ensurePage(store, "AllPages", "##AllPages")
		Call ensurePage(store, "LinkTypeDefinitions", "Type,, Colour" & vbCrLf & "____" & vbCrLf & "example,, #0066aa" & vbCrLf & "explanation,, #ff9900" & vbCrLf & "definition,, #339966" & vbCrLf & "counter,, #aa6633" & vbCrLf & "normal,, #000099")
		
		s = "name,, type,, maxDepth,, excluded pages,, excluded link types" & vbCrLf & "____" & vbCrLf & "simple_recursive ,, recursive,, -1,, ,, " & vbCrLf & "depth_one,, recursive,, 1,, ,, ,, " & vbCrLf & "recent_changes,, recent,, ,, ,, ,,"
		
		Call ensurePage(store, "CrawlerDefinitions", s)
		
		s = "Name,, Program,, Parameters" & vbCrLf & "____" & vbCrLf & "Dummy,, DummyExporter.exe,, Exports/MySite" & vbCrLf
		
		Call ensurePage(store, "ExportDefinitions", s)
		
		s = "<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN' " & vbCrLf & "'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'>" & vbCrLf & "<html xmlns='http://www.w3.org/1999/xhtml'>" & vbCrLf & "<head>" & vbCrLf & vbCrLf & "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1' />" & vbCrLf & "<meta name='generator' content='SdiDesk 0.2.0' />" & vbCrLf & "</head>" & vbCrLf & "----" & vbCrLf & "<body bgcolor='#ffffff' text='#000000'>" & vbCrLf & "----" & vbCrLf & "</body> </html>"
		
		Call ensurePage(store, "BasicHtmlTemplate", s)
		
	End Function
End Class