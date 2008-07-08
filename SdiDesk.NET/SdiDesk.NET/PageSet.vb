Option Strict Off
Option Explicit On
Friend Class PageSet
	
	' A PageSet is a collection of pages
	' You can fill it with pages and search them
	
	Public pages As OCollection
	
	Public Sub init()
		pages = New OCollection
	End Sub
	
	Public Sub addPage(ByRef p As _Page)
		If pages.hasKey((p.pageName)) Then
			pages.Remove((p.pageName))
			Call pages.Add(p, (p.pageName))
		Else
			Call pages.Add(p, (p.pageName))
		End If
	End Sub
	
	Public Sub addPageFromName(ByRef pageName As String, ByRef store As _PageStore)
		Dim p As _Page
		p = store.loadRaw(pageName)
		Call Me.addPage(p)
	End Sub
	
	Public Function hasPage(ByRef pName As String) As Boolean
		If pages.hasKey(pName) Then
			hasPage = True
		Else
			hasPage = False
		End If
	End Function
	
	Public Sub removePage(ByRef pName As String)
		If hasPage(pName) Then
			Call pages.Remove(pName)
		End If
	End Sub
	
	Public Sub clearOut()
		pages = New OCollection
	End Sub
	
	Public Sub merge(ByRef ps2 As PageSet)
		Dim size As Short
		Dim o As Object
		For	Each o In ps2.pages.toCollection
			'UPGRADE_WARNING: Couldn't resolve default property of object o. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call Me.addPage(o)
		Next o
	End Sub
	
	Public Function toWikiMarkup() As String
		Dim s As String
		s = ""
		Dim i As Object
		For	Each i In pages.toCollection
			'UPGRADE_WARNING: Couldn't resolve default property of object i.pageName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			s = s & "* [[" & i.pageName & "]]" & vbCrLf
		Next i
		toWikiMarkup = s
	End Function
	
	Public Sub saveAll(ByRef store As _PageStore)
		Dim i As Object
		For	Each i In pages.toCollection
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call store.savePage(i)
		Next i
	End Sub
	
	Public Function size() As Short
		size = pages.count
	End Function
End Class