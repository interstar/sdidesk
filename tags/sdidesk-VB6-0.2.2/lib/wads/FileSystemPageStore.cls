VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileSystemPageStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FileSystemPageStore is an implementation of the PageStore interface
' which uses the file system :-)

Option Explicit

Implements PageStore

Public mainDataDirectory As String ' name of the main directory
Public pagesDirectory As String ' pages subdir, usually main\pages
Public timeIndexDirectory As String ' where timeIndex goes, usually main\timeIndex
Public picturesDirectory As String ' where the pictures are
Public exporterDirectory As String ' where the exporters are

Public slash As String
Public subPageSeparator As String

Private ti As TimeIndex ' for managing the time index of pages

Private st As StringTool ' always useful

Public Function asPageStore() As PageStore
    Set asPageStore = Me
End Function

Public Sub setDataDirectory(dd As String)
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

Public Sub ensureDirectory(dirName As String)
   ' must be a nicer way of testing if directory exists
   ' but not in the manual :-(
   ' so we try to make it, and catch
   ' the error raised if it's already there
   
   Dim d As String
   d = mainDataDirectory & dirName
   Call ensureFullNameDirectory(d)
End Sub

Public Sub ensureFullNameDirectory(dirName As String)
      
   On Error GoTo AlreadyThere
     MkDir dirName

AlreadyThere:

End Sub

' file name processing

Public Function ensureTrailingSlash(s As String) As String
   ' makes sure any string has just one trailing slash
   ' eg. path\ becomes path\ and path becomes path\
   Dim path As String
   path = s
   If Right(path, 2) = (slash + slash) Then
      path = Left(path, Len(path) - 1)
   End If
   
   If Right(path, 1) <> slash Then
     path = path & slash
   End If
   
   ensureTrailingSlash = path
End Function


Public Function pathFromFileName(fName As String) As String
  ' strips off the fileName from the right of a path + file name
  pathFromFileName = Left(fName, InStrRev(fName, slash))
End Function

Public Function pageNameToFileName(pageName As String) As String
    Dim pn2 As String, d As String, pn3 As String
   
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

Function SubPageSeparatorForFileSystem(s As String) As String
    ' when we want to export html pages which are sub-pages we need a
    ' separator which is OK on both unix and windows.
    ' choose "--"
    SubPageSeparatorForFileSystem = Replace(s, "/", subPageSeparator)
End Function

Function FileExists(fileName As String) As Boolean
   Dim Msg As String
   ' Turn on error trapping so error handler responds
   ' if any error is detected.
   On Error GoTo CheckError
      FileExists = (Dir(fileName) <> "")
      ' Avoid executing error handler if no error
      ' occurs.
      Exit Function

CheckError:         ' Branch here if error occurs.
   ' Define constants to represent intrinsic Visual
   ' Basic error codes.
   Const mnErrDiskNotReady = 71, _
   mnErrDeviceUnavailable = 68
   ' vbExclamation, vbOK, vbCancel, vbCritical, and
   ' vbOKCancel are constants defined in the VBA type
   ' library.
   If (err.Number = mnErrDiskNotReady) Then
      Msg = "Put a floppy disk in the drive "
      Msg = Msg & "and close the door."
      ' Display message box with an exclamation mark
      ' icon and with OK and Cancel buttons.
      If MsgBox(Msg, vbExclamation & vbOKCancel) = _
      vbOK Then
         Resume
      Else
         Resume Next
      End If
   ElseIf err.Number = mnErrDeviceUnavailable Then
      Msg = "This drive or path does not exist: "
      Msg = Msg & fileName
      MsgBox Msg, vbExclamation
      Resume Next
   Else
      Msg = "Unexpected error #" & Str(err.Number)
      Msg = Msg & " occurred: " & err.Description
      ' Display message box with Stop sign icon and
      ' OK button.
      MsgBox Msg, vbCritical
      Stop
   End If
   Resume
End Function





Public Function loadPageFromFile(fileName As String, pageName As String) As Page

    ' this loads the raw data into a page object
    ' and sets it's type
    ' But DOES NOT cook
    
    Dim mrp As New MemoryResidentPage
    Dim p As Page
    Set p = mrp
    
    Dim line As String
    Dim stream As Integer
    stream = FreeFile
    
    p.raw = ""
    Dim pt As String
    
    If FileExists(fileName) Then
      ' file exists

      Open fileName For Input As #stream
      
      On Error GoTo inputError
      
      Input #stream, line
      p.pageName = line
      Dim Item As String
      p.categories = ""
      Line Input #stream, line
      p.categories = line
      
      Dim cd As String
      Input #stream, cd
      p.createdDate = PageStore_safeDate(cd)

      Input #stream, cd
      p.lastEdited = PageStore_safeDate(cd)

      Do Until EOF(1)
        Line Input #stream, line
        p.raw = p.raw + line + vbCrLf
      Loop
      Close #stream
      Call mrp.trimSpacesFromEnd

    Else
      ' new page
      p.raw = "new page"
      p.categories = ""
      p.pageName = pageName
    End If
    
    Set loadPageFromFile = p
    Exit Function
    
inputError:
    MsgBox ("Error reading file " & fileName)
    
End Function


Public Sub savePageToFile(p As Page, fName As String)
    ' don't call this directly, call savePage below
    ' (which calls it)
    ' We need this if we want to save a page
    ' into a file with a non-standard name
    
    ' let's make sure we have a directory
    Dim path As String
    Dim stream As Integer
    
    path = Me.pathFromFileName(fName)
    Call Me.ensureFullNameDirectory(path)
    
    stream = FreeFile
    ' now save this pages
    Open fName For Output As #stream
    Print #stream, p.pageName
    Print #stream, p.categories
    p.lastEdited = Date
    Print #stream, p.createdDate
    Print #stream, p.lastEdited
    Print #stream, p.raw
    Close #stream
    
End Sub



Public Sub renameFile(oldName As String, newName As String)
   
  Dim l As String
  Dim stream1 As Integer, stream2 As Integer

On Error GoTo forgetIt
  
  stream1 = FreeFile
  Open oldName For Input As #stream1
  stream2 = FreeFile
  Open newName For Output As #stream2

  Do Until EOF(stream1)
    Line Input #stream1, l
    Print #stream2, l
  Loop

forgetIt:
  ' file wasn't there so forget it
  Close #stream1
  Close #stream2

End Sub

Public Sub shiftOldFiles(fileName As String)
  Dim oName As String, oName2 As String
  Dim i As Integer
  For i = 5 To 2 Step -1
    oName = Left(fileName, Len(fileName) - 1)
    oName2 = Left(fileName, Len(fileName) - 1)
    oName = oName + CStr(i)
    oName2 = oName2 + CStr(i - 1)
    Call renameFile(oName2, oName) ' copy eg. file.mn2 to file.mn3
  Next i
  Call renameFile(fileName, oName2)
    
End Sub



Public Function dirAsVCollection(path2 As String) As VCollection
  Dim nextOne As String, path As String
  Dim vc As New VCollection
  
  path = Me.ensureTrailingSlash(path2)
  path = st.removeDoubleChar(path, slash)
  nextOne = Dir(path)
  
  While nextOne <> ""
     Call vc.add(nextOne, nextOne)
     nextOne = Dir()
  Wend
  Set dirAsVCollection = vc
End Function

Public Function dirAsPage(path2 As String, dlb As DirListBox)
  Dim build As String, nextOne As String, path As String, i As Integer
  
  Dim wmg As New WikiMarkupGopher
  
  path = Me.ensureTrailingSlash(path2)
  path = st.removeDoubleChar(path, slash)
  dlb.path = path
  
  build = ""
  
  If dlb.ListCount > 0 Then
    build = build + "==== Subdirectories ====" + vbCrLf
  End If
  
  For i = 0 To dlb.ListCount
     nextOne = dlb.List(i)
     nextOne = Replace(nextOne, (" " & slash), slash, 1, -1)
     If nextOne <> "" Then
       nextOne = st.removeDoubleChar(nextOne, slash)
       build = build + "* [[#dir " + nextOne + "]]" + vbCrLf
     End If
  Next i
  
  build = build + ">BOX<" + vbCrLf
  
  nextOne = Dir(path)
  If nextOne <> "" Then
    build = build + "==== Files ====" + vbCrLf
  End If
  
  While nextOne <> ""
     build = build + "##Local " + nextOne + ",, " + path + nextOne + vbCrLf
     If wmg.isImage(nextOne) Then
       build = build + "#NoWiki" + vbCrLf + "<img src='" + nextOne + "'>" + vbCrLf + "#Wiki" + vbCrLf
     End If
     nextOne = Dir()
  Wend
  
  Set wmg = Nothing
  dirAsPage = build
End Function

Private Sub Class_Initialize()
    slash = "\"
    subPageSeparator = "--"
    Call Me.setDataDirectory(App.path)
    Set ti = New TimeIndex
    Call ti.init(Me)
    Set st = New StringTool
End Sub

Private Sub Class_Terminate()
    Set st = Nothing
    Set ti = Nothing
End Sub

Private Function PageStore_getPageStoreIdentifier() As String
    PageStore_getPageStoreIdentifier = mainDataDirectory
End Function

Private Function PageStore_loadMonth(month As Integer, year As Integer) As String
   ' loads month from a file
   Dim p As Page
   Dim fileName As String
   fileName = ensureTrailingSlash(timeIndexDirectory) & year & slash & month & ".mnp"
   If (Dir(fileName) <> "") Then
     ' file already exists, load it
     Set p = loadPageFromFile(fileName, "" & year & "-" & month)
   Else
     ' file doesn't exist, so assume this month blank
     Set p = POLICY_getFactory().getNewPageInstance
   End If
   PageStore_loadMonth = p.raw
End Function

Public Sub ensureDirectoryByYearAndMonth(year As Integer, month As Integer)
   ' must be a nicer way of testing if directory exists
   ' but not in the manual :-(
   ' so we try to make it, and catch
   ' the error raised if it's already there
   
   Dim d As String
   d = ensureTrailingSlash(timeIndexDirectory) & year
   On Error GoTo AlreadyThere
     MkDir d

AlreadyThere:

End Sub


Private Function PageStore_loadUntilNotRedirectRaw(pageName As String) As Page
    Dim p2 As Page
    Dim isRedirect As Boolean
    Dim pName As String
    pName = pageName
    
    ' keep reading data page until not a #REDIRECT
    Do
        Set p2 = PageStore_loadRaw(pName)
        isRedirect = False
        If p2.isRedirect Then
            isRedirect = True
            pName = Right(p2.getFirstLine, Len(p2.getFirstLine) - 10)
        End If
    Loop Until isRedirect = False
    Set PageStore_loadUntilNotRedirectRaw = p2
End Function

Private Function PageStore_pageExists(pageName As String) As Boolean
   Dim fName As String
   fName = pageNameToFileName(pageName)
   If FileExists(fName) Then
      PageStore_pageExists = True
   Else
      PageStore_pageExists = False
   End If
End Function


Private Property Let PageStore_pictureLocality(ByVal RHS As String)
    picturesDirectory = RHS
End Property

Private Property Get PageStore_pictureLocality() As String
    PageStore_pictureLocality = "file:/" & picturesDirectory
End Property



Private Function PageStore_safeDate(s As String) As Date
  ' turns a string into a date but doesn't baulk if it breaks
  Dim d1 As Date, d2 As Date
  d2 = Date
  On Error GoTo broken
    d1 = CDate(s)
    If d1 <> CDate(0) Then ' make sure you overwrite any old zeroes
      d2 = d1
    End If
broken:
  PageStore_safeDate = d2
End Function

Public Function PageStore_loadRaw(pageName As String) As Page
   Dim fileName As String
   fileName = Me.pageNameToFileName(pageName)
   Set PageStore_loadRaw = loadPageFromFile(fileName, pageName)
End Function

Public Function PageStore_loadOldPage(pageName As String, version As Integer) As Page
   Dim fileName As String, f1 As String
   Dim p As Page
   fileName = Me.pageNameToFileName(pageName)
   If version > 5 Then
     MsgBox ("only 4 backups")
   Else
     f1 = Left(fileName, Len(fileName) - 1) & CStr(version)
     Set p = Me.loadPageFromFile(f1, pageName)
     Set PageStore_loadOldPage = p
     Set p = Nothing
   End If
End Function

Public Function PageStore_AllPages() As PageSet
  Dim ps As New PageSet, j As Integer
  Dim ps2 As PageSet
  
  Call ps.init
  
  For j = 0 To 9
     Set ps2 = PageStore_getPageSetOfAllPagesStartingWith(CStr(j))
     Call ps.merge(ps2)
  Next j
  
  For j = 65 To 90
     Set ps2 = PageStore_getPageSetOfAllPagesStartingWith(Chr(j + 33))
     Call ps.merge(ps2)
     Set ps2 = PageStore_getPageSetOfAllPagesStartingWith(Chr(j))
     Call ps.merge(ps2)
  Next j
  
  For j = 192 To 253
     Set ps2 = PageStore_getPageSetOfAllPagesStartingWith(Chr(j))
     Call ps.merge(ps2)
  Next j
  
  
  Set PageStore_AllPages = ps
End Function


Public Function PageStore_getPageSetContaining(searchText) As PageSet
  Dim ps As PageSet, ps2 As New PageSet
  Call ps2.init
  Set ps = PageStore_AllPages()
  Dim o As Object
  For Each o In ps.pages.toCollection
    If InStr(1, o.raw, searchText, 1) > 0 Then
       Call ps2.addPage(o)
    End If
  Next o
  
  Set PageStore_getPageSetContaining = ps2
End Function

Private Sub PageStore_saveMonth(month As Integer, year As Integer, body As String)
   Dim p As Page
   Set p = POLICY_getFactory().getNewPageInstance
   p.pageName = "" & year & "-" & month
   
   Dim fileName As String
   Call ensureDirectoryByYearAndMonth(year, month)
       
   fileName = ensureTrailingSlash(timeIndexDirectory) & year & slash & month & ".mnp"
   p.raw = body
   Call savePageToFile(p, fileName)
End Sub

Public Sub PageStore_savePage(p As Page)
    Dim fileName As String, firstLetter As String
    
    ' ensure the directory exists
    firstLetter = Left(p.pageName, 1)
    Call ensureDirectory(slash & "pages" & slash & firstLetter)
    
    ' now move the old files out of the way
    fileName = Me.pageNameToFileName(p.pageName)
    Call shiftOldFiles(fileName)
    

    ' now update the timeIndex
    Call ti.updateWord(p.pageName, p.lastEdited, Date)
    
    ' finally, save it
    Call savePageToFile(p, fileName)
End Sub

Public Function PageStore_deletePage(pageName As String)
   Dim s As String
   s = pageNameToFileName(pageName)
   Dim x As Integer
   x = MsgBox("Sure you want to remove page " & pageName & "?", 4)
   If x = 6 Then
     Kill s
   End If
   
End Function

Public Function PageStore_timeIndexAsWikiFormat(month As Integer, year As Integer, order As Boolean)
  Dim s As String
  s = PageStore_loadMonth(month, year)
  Call ti.parseMonthFromString(s)
  PageStore_timeIndexAsWikiFormat = ti.toWikiString(month, year, order)
End Function

Public Function PageStore_pageContains(pageName As String, searchText As String) As Boolean
  Dim r As String
  r = PageStore_loadRaw(pageName).raw
  If InStr(r, searchText) > 0 Then
    PageStore_pageContains = True
  Else
    PageStore_pageContains = False
  End If
End Function


Public Function PageStore_getPageSetOfAllPagesStartingWith(s As String) As PageSet
  Dim d As String, p As Page, ps As New PageSet
  
  Call ps.init
  
  d = Me.ensureTrailingSlash(pagesDirectory & s) & "*.mnp"
  Dim s3 As String, pageName As String ' s3 is the directory name
  Dim c As New Collection ' to store names in
  
  s3 = Dir(d)
  
  ' here we loop through getting names,
  ' then we loop through turning names into pages (1)
  ' Why? Because Me.loadRaw screws up the state of dir
  
  Do While s3 <> ""
    pageName = Left(s3, (Len(s3) - 4)) ' makes it a pageName
    Call c.add(pageName)
    s3 = Dir()
  Loop
  
  ' (1) turn those names into pages
  Dim v As Variant
  For Each v In c
    pageName = CStr(v)
    Set p = PageStore_loadRaw(pageName)
    Call ps.addPage(p)
  Next v
   
  Set PageStore_getPageSetOfAllPagesStartingWith = ps

End Function

