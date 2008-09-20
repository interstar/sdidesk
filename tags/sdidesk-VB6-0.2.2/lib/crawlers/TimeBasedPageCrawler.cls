VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimeBasedPageCrawler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements PageCrawler

' page crawler for finding recent pages / monthly pages
' at the moment, only gets recent
' as we need to work out how to give it all the parameters it
' needs ... the use of a single crawler definition table is breaking
' down!!1

Private myWads As WikiAnnotatedDataStore '
Private myStore As PageStore
Private myPages As PageSet ' where we keep the pages while crawling

Private myName As String ' useful to know the name of the crawler
Private myExcludedPages As VCollection  ' pages not to crawl to
Private myDefaultLinkTypeBehaviour As String ' undefined linkTypes
Private myLinkTypeBehaviours As VCollection ' linkTypes and what to do with them


Public Sub init(aName As String, ep As String, ltb As String)
   myName = aName
   Call parseExcludedPagesFromString(ep)
   Call parseLinkTypeBehavioursFromString(ltb)
End Sub

Public Sub parseExcludedPagesFromString(s As String)
  ' format of s is
  ' PageName1|PageName2|PageName3
  Set myExcludedPages = New VCollection
  Dim parts() As String
  If s <> "" Then ' if blank argument, do nothing
    If InStr(s, "|") > 0 Then ' multiple excluded pages
      parts = Split(s, "|")
      Dim v As Variant, v2 As String
      For Each v In parts
        v2 = CStr(v)
        If Not myExcludedPages.hasKey(v2) Then
          Call myExcludedPages.Add(v2, v2)
        End If
      Next v
    Else
      ' one excluded page
      Call myExcludedPages.Add(s, s)
    End If
  End If
End Sub


Public Sub parseLinkTypeBehavioursFromString(s As String)
  ' format of s is
  ' +explanation|definition|normal|counterArg
  ' which means, the default is exclude but include the following list
  ' alternatively
  ' -explanation|definition
  ' means that the default is include but exclude the following list
  
  Set myLinkTypeBehaviours = New VCollection

  If s = "" Or s = " " Then
    ' no args, defaults to +
    myDefaultLinkTypeBehaviour = "+"
    
  Else
    If Left(s, 1) = "+" Then
       myDefaultLinkTypeBehaviour = "-"
    Else
       myDefaultLinkTypeBehaviour = "+"
    End If
    ' the above might be confusing?
    ' if we put a + at the front, these are things we're explicitly *including*
    ' against a default of excluding
    ' if we put a - at the front, these are things we're explicity *excluding*
    ' against a default of including
    
    ' strip off the first char
    s = Right(s, Len(s) - 1)
    
    Dim parts() As String
    parts = Split(s, "|")
    Dim v As Variant, v2 As String
    For Each v In parts
      v2 = CStr(v)
      If Not myLinkTypeBehaviours.hasKey(v2) Then
        Call myLinkTypeBehaviours.Add(v2, v2)
      End If
    Next v
  End If
End Sub


Private Sub Class_Initialize()

  Set myPages = New PageSet
  Call myPages.init
End Sub

Private Sub PageCrawler_clear()
    ' resets the PageSet
    Set myPages = New PageSet
    Call myPages.init
End Sub

Private Sub PageCrawler_crawl(startPage As String)
    Dim se As New ScriptEngine, recentString As String
    myPages.clearOut
    recentString = se.recentChanges(myWads)
    Set myPages = PageCrawler_fillPageSetFromString(recentString)
End Sub

Private Function PageCrawler_fillPageSetFromPage(p As Page) As PageSet
    ' does nothing for this crawler
End Function

Private Function PageCrawler_fillPageSetFromString(s As String) As PageSet
    ' does nothing for this crawler
End Function

Private Function PageCrawler_getPages() As PageSet
    Set PageCrawler_getPages = myPages
End Function

Private Property Set PageCrawler_wads(ByVal RHS As WikiAnnotatedDataStore)
    Set myWads = RHS
End Property

Private Property Get PageCrawler_wads() As WikiAnnotatedDataStore
    Set PageCrawler_wads = myWads
End Property

Private Property Let PageCrawler_name(ByVal RHS As String)
    myName = RHS
End Property

Private Property Get PageCrawler_name() As String
    PageCrawler_name = myName
End Property

Private Property Set PageCrawler_pages(ByVal RHS As PageSet)
    Set PageCrawler_pages = RHS
End Property

Private Property Get PageCrawler_pages() As PageSet
    Set PageCrawler_pages = myPages
End Property

Private Property Set PageCrawler_store(ByVal RHS As PageStore)
    Set myStore = RHS
End Property

Private Property Get PageCrawler_store() As PageStore
    Set PageCrawler_store = myStore
End Property

Private Function PageCrawler_toString() As String
  Dim s As String
  s = "<p>'''" & myName & "''' is an example of a Time''''''Based''''''Page''''''Crawler </p>" & _
  "<p>It excludes these pages : " & myExcludedPages.toString & "</p>"
  PageCrawler_toString = s
End Function


