VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CrawlerDefinitionTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' this table defines the crawlers in the system
 
Public crawlerNames As VCollection
Public crawlers As OCollection

Private st As StringTool

Private Sub clear()
    Set crawlerNames = New VCollection
    Set crawlers = New OCollection
End Sub

Public Sub parseFromTableString(tabDef As String, model As ModelLevel, store As PageStore, chef As Object)
  ' we are expecting a crawler definition
  ' name,, type,, maxDepth,, excluded pages,, excluded types
  Call clear
  
  Dim i As Integer
  Dim pc As PageCrawler
  
  ' parse a table string to a crawler
  Dim aTable As New table
  aTable.parseFromDoubleCommaString (tabDef)
  
  Dim name As String, cType As String, md As Integer, exp As String, ext As String
     
  For i = 0 To aTable.noRows - 1
    name = CStr(aTable.at(i, 0))
    If name <> "" Then
        cType = st.strip(CStr(aTable.at(i, 1)))
        If (cType = "recursive") Then
            md = CInt(aTable.at(i, 2))
        Else
            md = 0
        End If
        exp = CStr(aTable.at(i, 3))
        ext = CStr(aTable.at(i, 4))
        
        Set pc = POLICY_getFactory().getNewPageCrawlerInstance(cType, name, md, exp, ext)
        Call crawlers.Add(pc, name)
        Call crawlerNames.Add(name, name)
    End If
  Next i
    
End Sub

Public Function getCrawler(name As String) As PageCrawler
  If crawlers.hasKey(name) Then
    Set getCrawler = crawlers.Item(name)
  Else
    MsgBox ("Error : no crawler called " + name)
  End If
End Function

Public Function toString() As String
  Dim c As PageCrawler, build As String
  build = ""
  For Each c In Me.crawlers.toCollection
    build = build + c.toString + vbCrLf
  Next c
  toString = build
End Function

Private Sub Class_Initialize()
    Set crawlerNames = New VCollection
    Set crawlers = New OCollection
    Set st = New StringTool
End Sub


Private Sub Class_Terminate()
    Set crawlerNames = Nothing
    Set crawlers = Nothing
    Set st = Nothing
End Sub
