Option Strict Off
Option Explicit On
Friend Class FileSystemPageStore
	Implements _PageStore
	' FileSystemPageStore is an implementation of the PageStore interface
	' which uses the file system :-)
	
	
	
	Public mainDataDirectory As String ' name of the main directory
	Public pagesDirectory As String ' pages subdir, usually main\pages
	Public timeIndexDirectory As String ' where timeIndex goes, usually main\timeIndex
	Public picturesDirectory As String ' where the pictures are
	Public exporterDirectory As String ' where the exporters are
	
	Public slash As String
	Public subPageSeparator As String
	
	Private ti As TimeIndex ' for managing the time index of pages
	
	Private st As StringTool ' always useful
	
	Public Function asPageStore() As _PageStore
		asPageStore = Me
	End Function
	
	Public Sub setDataDirectory(ByRef dd As String)
		mainDataDirectory = dd
		pagesDirectory = dd & slash & "pages" & slash
		timeIndexDirectory = dd & slash & "pages" & slash & "timeIndex" & slash
		picturesDirectory = dd & slash & "pages" & slash & "pictures" & slash
		exporterDirectory = dd & slash & "exporters" & slash
		Call Me.ensureFullNameDirectory(mainDataDirectory)
		Call Me.ensureFullNameDirectory(pagesDirectory)
		Call Me.ensureFullNameDirectory(timeIndexDirectory)
		Call Me.ensureFullNameDirectory(picturesDirectory)
		Call Me.ensureFullNameDirectory(exporterDirectory)
	End Sub
	
	Public Sub ensureDirectory(ByRef dirName As String)
		' must be a nicer way of testing if directory exists
		' but not in the manual :-(
		' so we try to make it, and catch
		' the error raised if it's already there
		
		Dim d As String
		d = mainDataDirectory & dirName
		Call ensureFullNameDirectory(d)
	End Sub
	
	Public Sub ensureFullNameDirectory(ByRef dirName As String)
		
		On Error GoTo AlreadyThere
		MkDir(dirName)
		
AlreadyThere: 
		
	End Sub
	
	' file name processing
	
	Public Function ensureTrailingSlash(ByRef s As String) As String
		' makes sure any string has just one trailing slash
		' eg. path\ becomes path\ and path becomes path\
		Dim path As String
		path = s
		If Right(path, 2) = (slash & slash) Then
			path = Left(path, Len(path) - 1)
		End If
		
		If Right(path, 1) <> slash Then
			path = path & slash
		End If
		
		ensureTrailingSlash = path
	End Function
	
	
	Public Function pathFromFileName(ByRef fName As String) As String
		' strips off the fileName from the right of a path + file name
		pathFromFileName = Left(fName, InStrRev(fName, slash))
	End Function
	
	Public Function pageNameToFileName(ByRef pageName As String) As String
		Dim d, pn2, pn3 As String
		
		pn2 = Replace(pageName, " ", "_")
		
		d = Me.ensureTrailingSlash(pagesDirectory & Left(pn2, 1))
		
		If InStr(pn2, "/") Then
			pn3 = Right(pn2, Len(pn2) - InStr(pn2, "/"))
			pn2 = Left(pn2, InStr(pn2, "/") - 1)
			'MsgBox (pn2 & " : " & pn3)
			d = d & pn2 & slash
			Call ensureFullNameDirectory(d)
			pageNameToFileName = d & pn3 & ".mnp"
		Else
			pageNameToFileName = d & pn2 & ".mnp"
		End If
		
	End Function
	
	Function SubPageSeparatorForFileSystem(ByRef s As String) As String
		' when we want to export html pages which are sub-pages we need a
		' separator which is OK on both unix and windows.
		' choose "--"
		SubPageSeparatorForFileSystem = Replace(s, "/", subPageSeparator)
	End Function
	
	Function FileExists(ByRef fileName As String) As Boolean
		Dim Msg As String
		' Turn on error trapping so error handler responds
		' if any error is detected.
		On Error GoTo CheckError
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileExists = (Dir(fileName) <> "")
		' Avoid executing error handler if no error
		' occurs.
		Exit Function
		
CheckError: ' Branch here if error occurs.
		' Define constants to represent intrinsic Visual
		' Basic error codes.
		Const mnErrDiskNotReady As Short = 71
		Const mnErrDeviceUnavailable As Short = 68
		' vbExclamation, vbOK, vbCancel, vbCritical, and
		' vbOKCancel are constants defined in the VBA type
		' library.
		If (Err.Number = mnErrDiskNotReady) Then
			Msg = "Put a floppy disk in the drive "
			Msg = Msg & "and close the door."
			' Display message box with an exclamation mark
			' icon and with OK and Cancel buttons.
			If MsgBox(Msg, CDbl(MsgBoxStyle.Exclamation & MsgBoxStyle.OKCancel)) = MsgBoxResult.OK Then
				Resume 
			Else
				Resume Next
			End If
		ElseIf Err.Number = mnErrDeviceUnavailable Then 
			Msg = "This drive or path does not exist: "
			Msg = Msg & fileName
			MsgBox(Msg, MsgBoxStyle.Exclamation)
			Resume Next
		Else
			Msg = "Unexpected error #" & Str(Err.Number)
			Msg = Msg & " occurred: " & Err.Description
			' Display message box with Stop sign icon and
			' OK button.
			MsgBox(Msg, MsgBoxStyle.Critical)
			Stop
		End If
		Resume 
	End Function
	
	
	
	
	
	Public Function loadPageFromFile(ByRef fileName As String, ByRef pageName As String) As _Page
		
		' this loads the raw data into a page object
		' and sets it's type
		' But DOES NOT cook
		
		Dim mrp As New MemoryResidentPage
		Dim p As _Page
		p = mrp
		
		Dim line As String
		Dim stream As Short
		stream = FreeFile
		
		p.raw = ""
		Dim pt As String
		
		Dim Item As String
		Dim cd As String
		If FileExists(fileName) Then
			' file exists
			
			FileOpen(stream, fileName, OpenMode.Input)
			
			On Error GoTo inputError
			
			Input(stream, line)
			p.pageName = line
			p.categories = ""
			line = LineInput(stream)
			p.categories = line
			
			Input(stream, cd)
			p.createdDate = PageStore_safeDate(cd)
			
			Input(stream, cd)
			p.lastEdited = PageStore_safeDate(cd)
			
			Do Until EOF(1)
				line = LineInput(stream)
				p.raw = p.raw & line & vbCrLf
			Loop 
			FileClose(stream)
			Call mrp.trimSpacesFromEnd()
			
		Else
			' new page
			p.raw = "new page"
			p.categories = ""
			p.pageName = pageName
		End If
		
		loadPageFromFile = p
		Exit Function
		
inputError: 
		MsgBox("Error reading file " & fileName)
		
	End Function
	
	
	Public Sub savePageToFile(ByRef p As _Page, ByRef fName As String)
		' don't call this directly, call savePage below
		' (which calls it)
		' We need this if we want to save a page
		' into a file with a non-standard name
		
		' let's make sure we have a directory
		Dim path As String
		Dim stream As Short
		
		path = Me.pathFromFileName(fName)
		Call Me.ensureFullNameDirectory(path)
		
		stream = FreeFile
		' now save this pages
		FileOpen(stream, fName, OpenMode.Output)
		PrintLine(stream, p.pageName)
		PrintLine(stream, p.categories)
		p.lastEdited = Today
		PrintLine(stream, p.createdDate)
		PrintLine(stream, p.lastEdited)
		PrintLine(stream, p.raw)
		FileClose(stream)
		
	End Sub
	
	
	
	Public Sub renameFile(ByRef oldName As String, ByRef newName As String)
		
		Dim l As String
		Dim stream1, stream2 As Short
		
		On Error GoTo forgetIt
		
		stream1 = FreeFile
		FileOpen(stream1, oldName, OpenMode.Input)
		stream2 = FreeFile
		FileOpen(stream2, newName, OpenMode.Output)
		
		Do Until EOF(stream1)
			l = LineInput(stream1)
			PrintLine(stream2, l)
		Loop 
		
forgetIt: 
		' file wasn't there so forget it
		FileClose(stream1)
		FileClose(stream2)
		
	End Sub
	
	Public Sub shiftOldFiles(ByRef fileName As String)
		Dim oName, oName2 As String
		Dim i As Short
		For i = 5 To 2 Step -1
			oName = Left(fileName, Len(fileName) - 1)
			oName2 = Left(fileName, Len(fileName) - 1)
			oName = oName & CStr(i)
			oName2 = oName2 & CStr(i - 1)
			Call renameFile(oName2, oName) ' copy eg. file.mn2 to file.mn3
		Next i
		Call renameFile(fileName, oName2)
		
	End Sub
	
	
	
	Public Function dirAsVCollection(ByRef path2 As String) As VCollection
		Dim nextOne, path As String
		Dim vc As New VCollection
		
		path = Me.ensureTrailingSlash(path2)
		path = st.removeDoubleChar(path, slash)
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		nextOne = Dir(path)
		
		While nextOne <> ""
			Call vc.add(nextOne, nextOne)
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			nextOne = Dir()
		End While
		dirAsVCollection = vc
	End Function
	
	Public Function dirAsPage(ByRef path2 As String, ByRef dlb As Microsoft.VisualBasic.Compatibility.VB6.DirListBox) As Object
		Dim nextOne, build, path As String
		Dim i As Short
		
		Dim wmg As New WikiMarkupGopher
		
		path = Me.ensureTrailingSlash(path2)
		path = st.removeDoubleChar(path, slash)
		dlb.Path = path
		
		build = ""
		
		If dlb.DirListCount > 0 Then
			build = build & "==== Subdirectories ====" & vbCrLf
		End If
		
		For i = 0 To dlb.DirListCount
			nextOne = dlb.DirList(i)
			nextOne = Replace(nextOne, " " & slash, slash, 1, -1)
			If nextOne <> "" Then
				nextOne = st.removeDoubleChar(nextOne, slash)
				build = build & "* [[#dir " & nextOne & "]]" & vbCrLf
			End If
		Next i
		
		build = build & ">BOX<" & vbCrLf
		
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		nextOne = Dir(path)
		If nextOne <> "" Then
			build = build & "==== Files ====" & vbCrLf
		End If
		
		While nextOne <> ""
			build = build & "##Local " & nextOne & ",, " & path & nextOne & vbCrLf
			If wmg.isImage(nextOne) Then
				build = build & "#NoWiki" & vbCrLf & "<img src='" & nextOne & "'>" & vbCrLf & "#Wiki" & vbCrLf
			End If
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			nextOne = Dir()
		End While
		
		'UPGRADE_NOTE: Object wmg may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		wmg = Nothing
		'UPGRADE_WARNING: Couldn't resolve default property of object dirAsPage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		dirAsPage = build
	End Function
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		slash = "\"
		subPageSeparator = "--"
		Call Me.setDataDirectory(My.Application.Info.DirectoryPath)
		ti = New TimeIndex
		'UPGRADE_WARNING: Couldn't resolve default property of object Me. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call ti.init(Me)
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
		'UPGRADE_NOTE: Object ti may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		ti = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Private Function PageStore_getPageStoreIdentifier() As String Implements _PageStore.getPageStoreIdentifier
		PageStore_getPageStoreIdentifier = mainDataDirectory
	End Function
	
	'UPGRADE_NOTE: year was upgraded to year_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: month was upgraded to month_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Function PageStore_loadMonth(ByRef month_Renamed As Short, ByRef year_Renamed As Short) As String Implements _PageStore.loadMonth
		' loads month from a file
		Dim p As _Page
		Dim fileName As String
		fileName = ensureTrailingSlash(timeIndexDirectory) & year_Renamed & slash & month_Renamed & ".mnp"
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If (Dir(fileName) <> "") Then
			' file already exists, load it
			p = loadPageFromFile(fileName, "" & year_Renamed & "-" & month_Renamed)
		Else
			' file doesn't exist, so assume this month blank
			p = POLICY_getFactory().getNewPageInstance
		End If
		PageStore_loadMonth = p.raw
	End Function
	
	'UPGRADE_NOTE: month was upgraded to month_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: year was upgraded to year_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub ensureDirectoryByYearAndMonth(ByRef year_Renamed As Short, ByRef month_Renamed As Short)
		' must be a nicer way of testing if directory exists
		' but not in the manual :-(
		' so we try to make it, and catch
		' the error raised if it's already there
		
		Dim d As String
		d = ensureTrailingSlash(timeIndexDirectory) & year_Renamed
		On Error GoTo AlreadyThere
		MkDir(d)
		
AlreadyThere: 
		
	End Sub
	
	
	Private Function PageStore_loadUntilNotRedirectRaw(ByRef pageName As String) As _Page Implements _PageStore.loadUntilNotRedirectRaw
		Dim p2 As _Page
		Dim isRedirect As Boolean
		Dim pName As String
		pName = pageName
		
		' keep reading data page until not a #REDIRECT
		Do 
			p2 = PageStore_loadRaw(pName)
			isRedirect = False
			If p2.isRedirect Then
				isRedirect = True
				pName = Right(p2.getFirstLine, Len(p2.getFirstLine) - 10)
			End If
		Loop Until isRedirect = False
		PageStore_loadUntilNotRedirectRaw = p2
	End Function
	
	Private Function PageStore_pageExists(ByRef pageName As String) As Boolean Implements _PageStore.pageExists
		Dim fName As String
		fName = pageNameToFileName(pageName)
		If FileExists(fName) Then
			PageStore_pageExists = True
		Else
			PageStore_pageExists = False
		End If
	End Function
	
	
	
	Private Property PageStore_pictureLocality() As String Implements _PageStore.pictureLocality
		Get
			PageStore_pictureLocality = "file:/" & picturesDirectory
		End Get
		Set(ByVal Value As String)
			picturesDirectory = Value
		End Set
	End Property
	
	
	
	Private Function PageStore_safeDate(ByRef s As String) As Date Implements _PageStore.safeDate
		' turns a string into a date but doesn't baulk if it breaks
		Dim d1, d2 As Date
		d2 = Today
		On Error GoTo broken
		d1 = CDate(s)
		If d1 <> System.Date.FromOADate(0) Then ' make sure you overwrite any old zeroes
			d2 = d1
		End If
broken: 
		PageStore_safeDate = d2
	End Function
	
	Public Function PageStore_loadRaw(ByRef pageName As String) As _Page Implements _PageStore.loadRaw
		Dim fileName As String
		fileName = Me.pageNameToFileName(pageName)
		PageStore_loadRaw = loadPageFromFile(fileName, pageName)
	End Function
	
	Public Function PageStore_loadOldPage(ByRef pageName As String, ByRef version As Short) As _Page Implements _PageStore.loadOldPage
		Dim fileName, f1 As String
		Dim p As _Page
		fileName = Me.pageNameToFileName(pageName)
		If version > 5 Then
			MsgBox("only 4 backups")
		Else
			f1 = Left(fileName, Len(fileName) - 1) & CStr(version)
			p = Me.loadPageFromFile(f1, pageName)
			PageStore_loadOldPage = p
			'UPGRADE_NOTE: Object p may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			p = Nothing
		End If
	End Function
	
	Public Function PageStore_AllPages() As PageSet Implements _PageStore.AllPages
		Dim ps As New PageSet
		Dim j As Short
		Dim ps2 As PageSet
		
		Call ps.init()
		
		For j = 0 To 9
			ps2 = PageStore_getPageSetOfAllPagesStartingWith(CStr(j))
			Call ps.merge(ps2)
		Next j
		
		For j = 65 To 90
			ps2 = PageStore_getPageSetOfAllPagesStartingWith(Chr(j + 33))
			Call ps.merge(ps2)
			ps2 = PageStore_getPageSetOfAllPagesStartingWith(Chr(j))
			Call ps.merge(ps2)
		Next j
		
		For j = 192 To 253
			ps2 = PageStore_getPageSetOfAllPagesStartingWith(Chr(j))
			Call ps.merge(ps2)
		Next j
		
		
		PageStore_AllPages = ps
	End Function
	
	
	Public Function PageStore_getPageSetContaining(ByRef searchText As Object) As PageSet Implements _PageStore.getPageSetContaining
		Dim ps As PageSet
		Dim ps2 As New PageSet
		Call ps2.init()
		ps = PageStore_AllPages()
		Dim o As Object
		For	Each o In ps.pages.toCollection
			'UPGRADE_WARNING: Couldn't resolve default property of object searchText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object o.raw. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If InStr(1, o.raw, searchText, 1) > 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object o. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call ps2.addPage(o)
			End If
		Next o
		
		PageStore_getPageSetContaining = ps2
	End Function
	
	'UPGRADE_NOTE: year was upgraded to year_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: month was upgraded to month_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub PageStore_saveMonth(ByRef month_Renamed As Short, ByRef year_Renamed As Short, ByRef body As String) Implements _PageStore.saveMonth
		Dim p As _Page
		p = POLICY_getFactory().getNewPageInstance
		p.pageName = "" & year_Renamed & "-" & month_Renamed
		
		Dim fileName As String
		Call ensureDirectoryByYearAndMonth(year_Renamed, month_Renamed)
		
		fileName = ensureTrailingSlash(timeIndexDirectory) & year_Renamed & slash & month_Renamed & ".mnp"
		p.raw = body
		Call savePageToFile(p, fileName)
	End Sub
	
	Public Sub PageStore_savePage(ByRef p As _Page) Implements _PageStore.savePage
		Dim fileName, firstLetter As String
		
		' ensure the directory exists
		firstLetter = Left(p.pageName, 1)
		Call ensureDirectory(slash & "pages" & slash & firstLetter)
		
		' now move the old files out of the way
		fileName = Me.pageNameToFileName((p.pageName))
		Call shiftOldFiles(fileName)
		
		
		' now update the timeIndex
		Call ti.updateWord((p.pageName), (p.lastEdited), Today)
		
		' finally, save it
		Call savePageToFile(p, fileName)
	End Sub
	
	Public Function PageStore_deletePage(ByRef pageName As String) As Object Implements _PageStore.deletePage
		Dim s As String
		s = pageNameToFileName(pageName)
		Dim x As Short
		x = MsgBox("Sure you want to remove page " & pageName & "?", 4)
		If x = 6 Then
			Kill(s)
		End If
		
	End Function
	
	'UPGRADE_NOTE: year was upgraded to year_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: month was upgraded to month_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function PageStore_timeIndexAsWikiFormat(ByRef month_Renamed As Short, ByRef year_Renamed As Short, ByRef order As Boolean) As Object Implements _PageStore.timeIndexAsWikiFormat
		Dim s As String
		s = PageStore_loadMonth(month_Renamed, year_Renamed)
		Call ti.parseMonthFromString(s)
		'UPGRADE_WARNING: Couldn't resolve default property of object PageStore_timeIndexAsWikiFormat. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PageStore_timeIndexAsWikiFormat = ti.toWikiString(month_Renamed, year_Renamed, order)
	End Function
	
	Public Function PageStore_pageContains(ByRef pageName As String, ByRef searchText As String) As Boolean Implements _PageStore.pageContains
		Dim r As String
		r = PageStore_loadRaw(pageName).raw
		If InStr(r, searchText) > 0 Then
			PageStore_pageContains = True
		Else
			PageStore_pageContains = False
		End If
	End Function
	
	
	Public Function PageStore_getPageSetOfAllPagesStartingWith(ByRef s As String) As PageSet Implements _PageStore.getPageSetOfAllPagesStartingWith
		Dim d As String
		Dim p As _Page
		Dim ps As New PageSet
		
		Call ps.init()
		
		d = Me.ensureTrailingSlash(pagesDirectory & s) & "*.mnp"
		Dim s3, pageName As String ' s3 is the directory name
		Dim c As New Collection ' to store names in
		
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		s3 = Dir(d)
		
		' here we loop through getting names,
		' then we loop through turning names into pages (1)
		' Why? Because Me.loadRaw screws up the state of dir
		
		Do While s3 <> ""
			pageName = Left(s3, Len(s3) - 4) ' makes it a pageName
			Call c.Add(pageName)
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			s3 = Dir()
		Loop 
		
		' (1) turn those names into pages
		Dim v As Object
		For	Each v In c
			'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pageName = CStr(v)
			p = PageStore_loadRaw(pageName)
			Call ps.addPage(p)
		Next v
		
		PageStore_getPageSetOfAllPagesStartingWith = ps
		
	End Function
End Class