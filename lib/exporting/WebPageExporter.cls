VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebPageExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements Exporter
Implements Export

' exports as a set of flat-web pages, now implements Exporter

Private myWads As WikiAnnotatedDataStore ' link to WADS
Private mySysConf As SystemConfigurations ' link to SysConf
Private myPrep As PagePreparer ' link to a page preparer

Private myFileSystemPageStore As FileSystemPageStore

' Export support
Private myExportName As String ' my name
Private myCrawlers As OCollection ' the collection of crawlers that define this export
Private myPages As PageSet ' the collection of pages to be exportet
Private myParameters As VCollection ' the parameters for the export

Private myPath As String ' my path
Private myStartPage As String ' my copy of the start page
Private myTemplate As HtmlTemplate ' the template for the HTML

' Exporter support
Private myExporterName As String
Private myParamPage As String
Private myPsi As String
Private myPageName As String

Private st As StringTool

Public Function asExporter() As Exporter
    Set asExporter = Me
End Function

Public Function asExport() As Export
    Set asExport = Me
End Function

Public Function getPsi() As String
    getPsi = myPsi
End Function


Public Sub setTemplatePage(p As page)
  Set myTemplate = New HtmlTemplate
  Call myTemplate.init(p)
End Sub

Private Function substituteVariables(s As String, p As page, tpl As HtmlTemplate) As String
    ' we want to do some interesting templating tricks
    ' here's a simple substitution language for HTML exports
    ' $PageName -> page.name
    ' $LastEdited -> page.last
    Dim s2 As String
    s2 = Replace(s, "$PageName", p.pageName)
    s2 = Replace(s2, "$LastEdited", CStr(p.lastEdited))
    s2 = Replace(s2, "$CreatedDate", CStr(p.createdDate))
    s2 = Replace(s2, "$ExportName", myExporterName)
    
    If InStr(s2, "$$") Then
        MsgBox ("has $$ in " & s2)
    End If
    substituteVariables = s2
End Function


Public Function pageToExportHTML(p As page, path As String, ps As PageSet, backlinks As Boolean) As String
    Dim s As String
    
    Dim prep As PagePreparer
    Set prep = POLICY_getFactory().getPagePreparer
    Call prep.prepare(p, backlinks)
    
    Dim chef As New ExportHtmlPageCooker
    
    Call chef.setPageSet(ps)
    
    s = "<html>" & vbCrLf & "<head>" & vbCrLf & _
    myTemplate.styleSheet & "</head>" & vbCrLf & myTemplate.header & _
    chef.asPageCooker.cook(p) & vbCrLf & myTemplate.footer & vbCrLf & "</html>"
    
    s = substituteVariables(s, p, myTemplate)
    ' clean up
    Set chef = Nothing
    Set prep = Nothing
    
    pageToExportHTML = s
End Function

Public Function saveStringInFile(s As String, fileName As String)

   On Error GoTo couldNotOpen
       Open fileName For Output As 1
       Print #1, s
       Close #1
       Exit Function

couldNotOpen:
   MsgBox ("sorry, couldn't open " + fileName + " to write to, this page wasn't exported")

End Function



Private Sub Class_Initialize()
   Set myCrawlers = New OCollection
   Set st = New StringTool
End Sub

Private Sub Class_Terminate()
   Set myCrawlers = Nothing
   Set myPages = Nothing
   Set myWads = Nothing
   
   Set myTemplate = Nothing
   Set st = Nothing
End Sub

Private Sub Export_addCrawler(pc As PageCrawler)
    Call myCrawlers.add(pc, pc.name)
End Sub

Private Sub Export_doExport()
    MsgBox ("Placeholder for doExport")
End Sub

Private Function Export_getPageSet() As Variant
    Set myPages = New PageSet
    Dim pc As PageCrawler
    For Each pc In myCrawlers.toCollection
        
    Next pc
End Function

Private Sub Export_readParameters(paramPage As String)
    Dim p As page
    Set p = myWads.store.loadRaw(paramPage)
    Call p.prepare(prep, False)
    Set myParameters = p.getDataDictionary
    myStartPage = myParameters.Item("startPage")
    myPath = myParameters.Item("path")
    Dim p2 As page
    Set p2 = myWads.store.loadRaw(myParameters.Item("template"))
    Call Me.setTemplatePage(p2)
    
    Set p = Nothing
    Set p2 = Nothing
End Sub

Private Sub Exporter_addCrawler(c As PageCrawler)
   Call myCrawlers.add(c, CStr(myCrawlers.Count))
End Sub

Private Function Exporter_canExportOnePage() As Boolean
    Exporter_canExportOnePage = True
End Function

Private Function Exporter_canInstant() As Boolean
    Exporter_canInstant = True
End Function

Private Function Exporter_canMulti() As Boolean
    Exporter_canMulti = True
End Function




Private Sub Exporter_doExport(exp As Export)
    MsgBox ("Placeholder for Export_doExport")
End Sub

Private Sub Exporter_doInstantExport(page As page)
    Set myTemplate = New HtmlTemplate
    Dim tPage As page
    Set tPage = myFileSystemPageStore.asPageStore.loadRaw("BasicHtmlTemplate")
    Call myTemplate.init(tPage)
    
End Sub

Private Sub Exporter_doInstantExportPageName(pageName As String)
    Dim p As page
    Set p = myFileSystemPageStore.asPageStore.loadRaw(pageName)
    Call Exporter_doInstantExport(p)
    Set p = Nothing
End Sub

Private Sub Exporter_export()
   Dim o As Object, fileName As String
   Dim fStore As New FileSystemPageStore
   For Each o In myPages.pages.toCollection
        Call exportOnePage(fStore, o, myPages, myPath, False)
   Next o
   Set fStore = Nothing
End Sub


Private Sub exportOnePage(p As page, path As String)
    Dim exportText As String, fileName As String
    fileName = myFileSystemPageStore.SubPageSeparatorForFileSystem(p.pageName)
    fileName = myFileSystemPageStore.ensureTrailingSlash(path) & fileName & ".html"
    exportText = Me.pageToExportHTML(p, path, ps, backlinks)
    If InStr(p.pageName, ":") > 0 Then
        MsgBox ("error in WebPageExporter " + p.pageName + ", " + fileName)
    Else
        Call Me.saveStringInFile(exportText, fileName)
    End If
End Sub


Private Sub XXexportThisPage(p As page)
    Dim ps As New PageSet
    Dim fs As New FileSystemPageStore
    Call exportOnePage(fs, p, ps, myPath, False)
    Set ps = Nothing
    Set fs = Nothing
End Sub


Private Function XXgetPages() As PageSet
  If myPages Is Nothing Then
    Set myPages = New PageSet
  End If
  If myPages.size = 0 Then
    Dim o As Object, cr As PageCrawler
    For Each o In myCrawlers
      Set cr = o
      Call cr.crawl(myStartPage)
      myPages.merge (cr.getPages)
    Next o
  End If
  Set Exporter_getPages = myPages
  Set cr = Nothing
  Set o = Nothing
End Function

Private Function Exporter_getVersion() As String
    Exporter_getVersion = "WebPageExporter 0.1.0"
End Function

Private Property Set Exporter_model(ByVal RHS As ModelLevel)
   Set myWads = RHS
End Property

Private Property Get Exporter_model() As ModelLevel
  Set Exporter_model = myWads
End Property

Private Property Let Exporter_name(ByVal RHS As String)
  myName = RHS
End Property

Private Property Get Exporter_name() As String
  Exporter_name = myName
End Property

Private Property Set Exporter_pages(ByVal RHS As PageSet)
  Set myPages = RHS
End Property

Private Property Get Exporter_pages() As PageSet
  Set Exporter_pages = myPages
End Property

Private Property Let Exporter_path(ByVal RHS As String)
  myPath = RHS
End Property

Private Property Get Exporter_path() As String
    Exporter_path = myPath
End Property

Private Function Exporter_readCommand() As VCollection
    Dim laa As New LineArgAnaliser
    laa.analise (command)
    If Not laa.argDict.hasKey("psi") Then
        MsgBox ("Error in WebPageExporter.Exporter_readCommand No command line arguments")
        End
    Else
        myPsi = laa.argDict.Item("psi")
        myExporterName = laa.argDict.Item("name")
        If laa.argDict.hasKey("page") Then
            ' this is an "instant" export
            myPageName = laa.argDict.Item("page")
            myPsi = laa.argDict.Item("psi")
        Else
            myParamPage = laa.argDict.Item("param")
        End If
    End If
    Set myFileSystemPageStore = New FileSystemPageStore
    myFileSystemPageStore.setDataDirectory (myPsi & "\")
    Set Exporter_readCommand = laa.asVCollection
End Function


