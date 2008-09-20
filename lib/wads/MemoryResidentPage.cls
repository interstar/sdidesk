VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemoryResidentPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements Page

' this is the basic page object which holds data about a page
' As far as possible, EVERYTHING is a page in sdiDesk
 
' implement properties
Private myRaw As String ' the raw text of the page
Private myPrepared As String ' done includes and inlines, but not prettification
Private myCooked As String ' the presentation view of the page.

Private myPageName As String ' name of the page
Private myCategories As String ' the categories box
Private myCreatedDate As Date ' date this was created
Private myLastEdited As Date ' date last edited

' instance vars for this page
Private dataDictionary As VCollection

' alternative versions
Private myTable As table

Private myType As String

' helpers
Private st As StringTool

Public Function asPage() As Page
    Set asPage = Me
End Function

Public Function getFirstLine() As String
  Dim lines() As String
  If InStr(myRaw, vbCrLf) Then
      lines = Split(myRaw, vbCrLf)
      getFirstLine = lines(0)
  Else
      getFirstLine = myRaw
  End If
End Function

Public Sub trimSpacesFromEnd()
    Dim flag As Boolean
    flag = False
    While flag = False
      If Len(myRaw) > 1 Then
         If Right(myRaw, 1) = vbCrLf Or Right(myRaw, 1) = " " Or Asc(Right(myRaw, 1)) = 10 Then
            myRaw = Left(myRaw, Len(myRaw) - 2)
         Else
            flag = True
         End If
      Else
         flag = True
      End If
    Wend
End Sub

Private Sub Class_Initialize()
    Set dataDictionary = New VCollection
    Set st = New StringTool
    myType = "new page"
End Sub

Private Sub Class_Terminate()
    Set dataDictionary = Nothing
    Set st = Nothing
End Sub

Private Property Let Page_categories(ByVal RHS As String)
    myCategories = RHS
End Property

Private Property Get Page_categories() As String
    Page_categories = myCategories
End Property

Private Sub Page_cook(prep As PagePreparer, chef As PageCooker, backlinks As Boolean)
    Dim t As New table
    If t.isValidTable(myRaw) And InStr(myRaw, "____") Then
        myType = "table"
    Else
        myType = "normal"
    End If
    Set t = Nothing
  
    Call prep.prepare(Me, backlinks)
    myCooked = chef.cook(Me)
End Sub

Private Property Get Page_cooked() As String
    Page_cooked = myCooked
End Property

Private Property Let Page_cooked(ByVal RHS As String)
    myCooked = RHS
End Property

Private Property Let Page_createdDate(ByVal RHS As Date)
    myCreatedDate = RHS
End Property

Private Property Get Page_createdDate() As Date
    Page_createdDate = myCreatedDate
End Property

Private Function Page_getDataDictionary() As VCollection
    Set Page_getDataDictionary = dataDictionary
End Function

Private Function Page_getFirstLine() As String
    Page_getFirstLine = getFirstLine()
End Function

Private Function Page_getMyType() As String
   Dim t As String
   t = "normal"
   If Page_isNetwork Then t = "network"
   If Page_isTable Then t = "table"
   If Page_isRedirect Then t = "redirect"
   If Page_isNew Then t = "new page"
   Page_getMyType = t
End Function



Private Function Page_getRedirectPage() As String
    Dim s As String
    s = getFirstLine()
    Dim parts() As String
    parts = Split(s, " ")
    Page_getRedirectPage = parts(1)
End Function

Private Function Page_getTable() As table
    If Page_isTable Then
        Set Page_getTable = myTable
    Else
        MsgBox ("Error trying to get table")
        End
    End If
End Function

Private Function Page_getVal(key As String) As String
  Dim k As String
  k = st.strip(key)
  If Page_hasVar(k) Then
    Page_getVal = dataDictionary.Item(k)
  Else
    Page_getVal = "<font color='#990000'> undefined variable : " + k + " on page " + myPageName + "</font>"
  End If
End Function

Private Function Page_hasVar(key As String) As Boolean
  Dim k As String
  k = st.strip(key)
  Page_hasVar = dataDictionary.hasKey(k)
End Function

Private Function Page_isNetwork() As Boolean
  Dim f As String
  f = getFirstLine()
  If Len(f) > 8 And Left(f, 8) = "#Network" Then
      Page_isNetwork = True
    Else
      Page_isNetwork = False
  End If
End Function

Private Function Page_isNew() As Boolean
    If myType = "new page" Then
        Page_isNew = True
    Else
        Page_isNew = False
    End If
End Function

Private Function Page_isRedirect() As Boolean
  Dim f As String
  f = getFirstLine()
  If InStr(f, "#REDIRECT ") > 0 Then
    Page_isRedirect = True
  Else
    Page_isRedirect = False
  End If
End Function

Private Function Page_isTable() As Boolean
    If myType = "table" Then
        Page_isTable = True
    Else
        Page_isTable = False
    End If
End Function

Private Property Get Page_lastEdited() As Date
    Page_lastEdited = myLastEdited
End Property

Private Property Let Page_lastEdited(ByVal RHS As Date)
    myLastEdited = RHS
End Property

Private Property Let Page_pageName(ByVal RHS As String)
    myPageName = RHS
End Property

Private Property Get Page_pageName() As String
    Page_pageName = myPageName
End Property

Private Property Let Page_pageType(ByVal RHS As String)
    myType = RHS
End Property

Private Property Get Page_pageType() As String
    Page_pageType = myType
End Property

Private Sub Page_prepare(prep As PagePreparer, backlinks As Boolean)
  If Not Page_isNetwork() Then
    Dim t As New table
    If t.isValidTable(myRaw) And InStr(myRaw, "____") Then
        myType = "table"
    Else
        myType = "normal"
    End If
    Set t = Nothing
  
    Call prep.prepare(Me, backlinks)
  End If
End Sub

Private Property Let Page_prepared(ByVal RHS As String)
    myPrepared = RHS
End Property

Private Property Get Page_prepared() As String
    Page_prepared = myPrepared
End Property

Private Property Let Page_raw(ByVal RHS As String)
    myRaw = RHS
End Property

Private Property Get Page_raw() As String
    Page_raw = myRaw
End Property


Private Sub Page_setVal(aKey As String, aVal As String)
  Dim v As String, k As String
  v = st.strip(aVal)
  k = st.strip(aKey)
  If dataDictionary.hasKey(k) Then
    Call dataDictionary.Remove(k)
  End If
  
  Call dataDictionary.Add(v, k)
End Sub

Private Function Page_spawnCopy() As Page
  ' creates a new object of class MemoryResidentPage, and populates it with copies
  ' of all data except cooked and network
  Dim p2 As Page
  Set p2 = POLICY_getFactory().getNewPageInstance()
  p2.raw = myRaw
  p2.pageName = myPageName
  p2.categories = myCategories
  p2.createdDate = myCreatedDate
  p2.lastEdited = myLastEdited
  Set Page_spawnCopy = p2
  Set p2 = Nothing
End Function

Private Function Page_varsToString() As String
  Page_varsToString = dataDictionary.toString
End Function

Private Function Page_wordCount() As Integer
    ' nb : a rough word-count
    Dim lines() As String
    Dim words() As String
    Dim v As Variant, w As Variant
    Dim s As String
    Dim wmg As New WikiMarkupGopher
    Dim Count As Integer
    Count = 0
    lines = Split(myRaw, vbCrLf)
    For Each w In lines
        words = Split(CStr(w), " ")
        For Each v In words
            s = CStr(v)
            If s <> "" And wmg.isAlpha(Left(s, 1)) Then
                Count = Count + 1
            End If
        Next v
    Next w
    Page_wordCount = Count
    Set wmg = Nothing
End Function
