VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RecursivePageCrawler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements PageCrawler

Private myWads As WikiAnnotatedDataStore  ' where to get everything else

Private myPages As PageSet ' where we keep the pages while crawling

Private myName As String ' useful to know the name of the crawler
Private myMaxDepth As Integer ' how deep to crawl
Private myExcludedPages As VCollection  ' pages not to crawl to
Private myLinkTypeBehaviours As VCollection ' linkTypes and what to do with them
Private myDefaultLinkTypeBehaviour As String ' are we including or excluding

Public Sub init(aName As String, maxDepth As Integer, excludedPages As String, linkTypeBehaviours As String)
  myName = aName
  myMaxDepth = maxDepth
  Call parseExcludedPagesFromString(excludedPages)
  Call parseLinkTypeBehavioursFromString(linkTypeBehaviours)
End Sub

Public Function asPageCrawler() As PageCrawler
    Set asPageCrawler = Me
End Function

Public Sub parseExcludedPagesFromString(s As String)
  ' format of s is
  ' PageName|PageName2|PageName3
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


Public Sub recursiveCrawl(startPage As String, depth As Integer)
  ' this function starts with the name of a page
  ' gathers it's out-links, then
  ' follows out-links to any other pages not currently in the main set
  ' so DOESN'T go circular
  ' and has the possibility of a maximum depth restriction
  ' At the end, we should have gathered all relevant pages into pages
  
  ' let's go ...
  
  If depth < myMaxDepth Or myMaxDepth < 0 Then
    ' otherwise we're not going anywhere.
    ' note if you set maxDepth to say, -1, then there's no maximum
    
    Dim ps As New PageSet, r As String
    ps.init
    
    ' now fill it from the out-links from startPage
    ' this doesn't handle extra links from including. Should we?
    r = myWads.getRawPageData(startPage)
    Set ps = PageCrawler_fillPageSetFromString(r)
    
    ' now iterate through it
    
    Dim p As Object
    Dim recurse As Boolean
    
    For Each p In ps.pages.toCollection
      
      ' if this page has NOT yet been gathered into our "pages" set,
      ' we will recurse into it
      If Not myPages.hasPage(p.pageName) Then
         ' ie. we DON'T yet have this page
         ' let's have it
         Call myPages.addPage(p)
      
         ' and let's do the recursion
         Call Me.recursiveCrawl(p.pageName, depth + 1)
      End If
      
    Next p
      
  Else 'depth >= maxDepth, just come out
  End If
  
  Set p = Nothing
  Set ps = Nothing

End Sub



Private Sub Class_Initialize()

  Set myPages = New PageSet
  Call myPages.init
  
End Sub

Private Sub Class_Terminate()
    Set myPages = Nothing
    Set myWads = Nothing
    Set myPages = Nothing

    Set myExcludedPages = Nothing
    Set myLinkTypeBehaviours = Nothing
End Sub


Private Sub PageCrawler_clear()
    ' resets the PageSet
    Set myPages = New PageSet
    Call myPages.init
End Sub

Private Sub PageCrawler_crawl(startPage As String)
     myPages.clearOut
     Call Me.recursiveCrawl(startPage, 0)
End Sub

Private Function linkExcluded(lnk As link) As Boolean
    If lnk.isCommand() Then
        linkExcluded = True
        Exit Function
    End If

    If (myDefaultLinkTypeBehaviour = "+") Then
        ' exclusions are explicit,
        If myLinkTypeBehaviours.hasKey(lnk.linkType) Then
            ' it's in the excluded list
            linkExcluded = True
            Exit Function
        End If
    Else
        ' inclusions are explicit
        If Not myLinkTypeBehaviours.hasKey(lnk.linkType) Then
            ' it's not in the included list
            linkExcluded = True
            Exit Function
        End If
    End If
    
    If Not myWads.pageExists(lnk.target) Then
        linkExcluded = True
        Exit Function
    End If
    
    ' but now let's see if this page itself if explicitly excluded
    linkExcluded = myExcludedPages.hasKey(lnk.target)

End Function

Private Function filterOutlinks(initialOuts As OCollection) As PageSet
    Dim lnk As link
    Dim ps As New PageSet, p As Page
    Call ps.init
    Call ps.clearOut
    
    For Each lnk In initialOuts.toCollection
        If Not linkExcluded(lnk) Then
            Set p = myWads.store.loadRaw(lnk.target)
            Call ps.addPage(p)
        End If
    Next lnk
    
    Set filterOutlinks = ps
End Function

Private Function outLinksFromNetwork() As OCollection
    
End Function

Private Function PageCrawler_fillPageSetFromPage(p As Page) As PageSet
    If Not p.isNetwork Then
        Call p.cook(POLICY_getFactory().getPagePreparer, POLICY_getFactory().getNativePageCooker, False)
        PageCrawler_fillPageSetFromPage = PageCrawler_fillPageSetFromString(p.prepared)
    Else
    
    End If
End Function

Private Function PageCrawler_fillPageSetFromString(s As String) As PageSet
  ' Creates a new PageSet and fills it
  ' from the raw links defined in s

  Dim outlinks As OCollection
  
  Dim lp As LinkProcessor
  Set lp = POLICY_getFactory().getStandardLinkProcessor
  
  Set outlinks = lp.getAllLinksInBigDocument(s)
  
  Set PageCrawler_fillPageSetFromString = filterOutlinks(outlinks)
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


Private Function PageCrawler_toString() As String
  Dim s As String
  s = "BOX<" & vbCrLf & "<p>'''" & myName & "''' is an example of a Recursive''''''Page''''''Crawler</p>" & _
  "<p>It follows links to a maximum depth of " & myMaxDepth & " from the start-page</p>" & _
  "<p>It excludes these pages : " & vbCrLf & myExcludedPages.toString & "</p>" & vbCrLf
  If myDefaultLinkTypeBehaviour = "-" Then
    s = s & "<p>It also ignores all links except those of these types : " & vbCrLf & _
    myLinkTypeBehaviours.toString & vbCrLf & "</P>" & vbCrLf
  Else
    s = s & "<p>It also ignores all links of these types : " & _
    myLinkTypeBehaviours.toString & vbCrLf & "</p>" & vbCrLf
  End If
  s = s & ">BOX" & vbCrLf & "<p>To edit or add a new crawler, please go to the CrawlerDefinitions page.</p>"
  PageCrawler_toString = s
End Function

