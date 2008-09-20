VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExportManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' The new export model goes like this :
' different types of exporting are handled by separate custom programs.

' "Exporter" is the name we will use for these programs
' "Export" is the name we'll use for an invocation of an export,
' based on a particular page-set, template etc.

' For example, HtmlExporter could be the name of a program to
' export pages as a flat HTML site. It is an "Exporter"

' my_site might be a call of HtmlExporter, giving it a simple_recursive
' PageCrawler (to collect a set of pages) starting on a page called MySiteHome.
' my_site is an "Export"

' The ExportRecord now holds only three pieces of data, which define an Export :
' the name of the Export
' the name of the Exporter (program which will do the exporting)
' the name of a page which contains all the parameters to define the export

' The ExportManager is still a table of ExportRecords but
' it does NOT *do* the export

' The main program does *not* include Exporters!!!
' It simply calls external Exporter programs through the VB shell command

' ExportManager is the object which does this
' although ExportSubsystem also has a call it needs to pass on

Public exportTable As OCollection ' stores the table of different exports
Public exportNames As VCollection ' stores the export names

Public exportPrograms As VCollection ' stores list of export programs

Private st As StringTool ' always useful
  
Public Sub parseFromRawString(s As String)
  ' we're expecting a simple, double-comma separated table
  ' name,, program,, parameter-page
  
  Dim t As New table
  Dim i As Integer
  Dim e As ExportRecord
  
  Dim aName As String
  Dim aProgram As String
  Dim aParamPage As String
  
  Set exportTable = New OCollection
  Set exportNames = New VCollection
  
  Call t.parseFromDoubleCommaString(s)
  
  For i = 0 To t.noRows - 1
    Set e = New ExportRecord
      
    aName = st.strip(CStr(t.at(i, 0)))
    aProgram = st.strip(CStr(t.at(i, 1)))
    aParamPage = st.strip(CStr(t.at(i, 2)))
        
    Call e.init(aName, aProgram, aParamPage)
    Call exportTable.Add(e, aName)
    Call exportNames.Add(aName, aName)
  
  Next i
  Set e = Nothing
End Sub

Public Sub scanForExporters(lfs As LocalFileSystem)
    If lfs.hasLocalFileSystem Then
        Dim vc As New VCollection
        Dim v As Variant
        Set vc = lfs.getDirectoryListingAsVCollection(lfs.getExporterDirectory)
        For Each v In vc.toCollection
            If Right(CStr(v), 4) = ".exe" Then
               Call exportPrograms.Add(CStr(v), CStr(v))
            End If
        Next v
    Else
        MsgBox ("Look, I'm very sorry but I can't seem to scan for exports on this system. No exporting is possible.")
    End If
End Sub

Public Function toString() As String
  Dim o As Object
  Dim s As String
  s = ""
  For Each o In exportTable.toCollection
    s = s + o.toString + vbCrLf
  Next o
  toString = s
End Function

Public Function exportersToString() As String
    Dim v As Variant
    Dim s As String
    s = ""
    For Each v In exportPrograms.toCollection
        s = s & v & ", "
    Next v
    exportersToString = s
    
End Function

Public Sub callExport(aName As String, pageStoreIdentifier As String)
    ' pageStoreIdentifier is a string which will let the exporter
    ' find the page store. Currently, it will be a directory,
    ' in future, may be a URL
    
    Dim e As ExportRecord
    Set e = exportTable.Item(aName)
    On Error GoTo notFound
    Dim pn As String
    pn = "exporters\" & e.program
    Call Shell(pn & " -name " & e.name & " -param '" & e.paramPage & "' -psi '" & pageStoreIdentifier & "\'", vbNormalFocus)
    Exit Sub

notFound:
    MsgBox ("Couldn't find a program called '" & pn & "' in your exporters directory. Check what exporters are available (Export Menu:Show Exporters)")
    
End Sub

Public Sub callInstantExport(progName As String, pageStoreIdentifier As String, currentPageName As String)
    On Error GoTo notFound
    Dim pn As String
    pn = "exporters\" & progName
    Call Shell(pn & " -page '" & currentPageName & "' -psi '" & pageStoreIdentifier & "\'")
    Exit Sub

notFound:
    MsgBox ("Couldn't find a program called '" & pn & "' in your exporters directory. Check what exporters are available (Export Menu:Show Exporters)")
    
End Sub

Private Sub Class_Initialize()
    Set exportTable = New OCollection
    Set exportNames = New VCollection
    Set exportPrograms = New VCollection

    Set st = New StringTool
End Sub

Private Sub Class_Terminate()
    Set exportTable = Nothing
    Set exportNames = Nothing
    Set exportPrograms = Nothing
    Set st = Nothing
End Sub
