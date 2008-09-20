VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScriptEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' performs some of the more complex stuff which
' you can do with ##Inlines
' not much of a scripting language yet, but
' the name represents the aspiration of where we're going ;-)

Private ti As TimeIndex

Public Function perform(line As String, model As ModelLevel) As String
  Dim command As String, argList As String, st As New StringTool
  
  If model Is Nothing Then
    MsgBox ("perform, model is nothing")
  End If
  
  command = st.strip(st.leftsa(line, " ", 1)) ' grab the first left word of line as command
  argList = st.stripHead(line, " ", 1) ' rest of the line
  
  Dim tokens() As String
  tokens = Split(st.strip(argList), ",, ")
  
  Select Case command
    Case "##AllPages":
       perform = AllPages(model)
    Case "##Include":
       If UBound(tokens) > -1 Then
            perform = include(st.strip(tokens(0)), model)
       Else
            perform = "<font color='red'>Error in ##Include " + argList
       End If
    Case "##TableInclude":
       perform = tableInclude(tokens, model)
    Case "##Image":
       perform = imageInclude(tokens, model)
    Case "##Month":
       perform = monthChanges(tokens, model)
    Case "##CalendarMonth":
        perform = CalendarMonth(tokens, model)
    Case "##IncludingCalendarMonth":
        perform = IncludingCalendarMonth(tokens, model)
    Case "##CalendarEntries":
        perform = CalendarEntries(tokens, model)
    Case "##RecentChanges":
       perform = recentChanges(model)
    Case "##Find":
       perform = find(argList, model)
    Case "##Local":
       If UBound(tokens) > 0 Then
           perform = localResource(st.strip(tokens(0)), st.strip(tokens(1)), model)
       Else
           perform = localResource(st.strip(tokens(0)), st.strip(tokens(0)), model)
       End If
    Case "##Dir":
       perform = Me.localDir(argList, model)
    Case "##WordCount":
        perform = "Word count of " & tokens(0) & " is " & model.getControllableModel(tokens(0))
    Case "##Button":
        If UBound(tokens) > 0 Then
            perform = makeButton(tokens(0), tokens(1))
        Else
            perform = err("Sorry, bad arguments for Button")
        End If
    Case Else:
       perform = err("Sorry, don't know how to '" + line + "'")
  End Select
End Function


Public Function err(s As String) As String
  err = "<font color=##ff0000 size=+1>" + s + "</font>" + vbCrLf
End Function

Public Function AllPages(model As ModelLevel) As String
  Dim ps As PageSet, returnVal As String
  returnVal = err("Sorry, problem getting all pages")
  On Error GoTo err1
    Set ps = model.getWikiAnnotatedDataStore.store.AllPages()
    Dim s As String
    s = CStr(ps.size()) + " pages " + vbCrLf
    returnVal = s + ps.toWikiMarkup
  
err1:
  
  AllPages = returnVal
  
End Function


Public Function include(otherPage As String, wads As WikiAnnotatedDataStore) As String
  include = wads.getRawPageData(otherPage)
End Function


Public Function tableInclude(tokens() As String, model As ModelLevel) As String
  
  Dim tableRaw As String, returnValue As String
  
  returnValue = err("Sorry, error including this table")
  On Error GoTo err1
  
  ' get the raw page data
  tableRaw = model.getWikiAnnotatedDataStore.getRawPageData(tokens(0))
  
  returnValue = err("Sorry, problem TableIncluding " + tokens(0))
  
  Dim t As New table, t2 As New table
  ' parse page data into table
    
  If t.isValidTable(tableRaw) Then
  
    Call t.parseFromDoubleCommaString(tableRaw)
      
    Dim at As New ArrayTool
    If UBound(tokens) = 0 Then
      ' this has no further arguments, show whole thing
      Set t2 = t
    Else
      ' collect arguments 2 +
      ReDim colIndexes(UBound(tokens)) As String
      Call at.copyStringArray(tokens, colIndexes(), 1, UBound(tokens), 0)
  
      ' now get columns out
      Call t2.project(t, Join(colIndexes, " "))
    End If
  
    returnValue = t2.toWikiFormat
  Else
    returnValue = err("Error trying to TableInclude [[#raw " + tokens(0) + "]]")
  End If
  
err1:
  tableInclude = returnValue
End Function

Public Function qq(s As String) As String
  qq = Chr(34) + s + Chr(34)
End Function

Public Function imageInclude(tokens() As String, model As ModelLevel) As String
  Dim imageName As String, linkTarget As String, returnValue As String, d As String, iTag As String
  returnValue = err("Sorry, couldn't include image, no name")
  On Error GoTo err1
  imageName = tokens(0)
  
  returnValue = err("Sorry, couldn't include image called " + imageName)
  
  d = model.getWikiAnnotatedDataStore.store.pictureLocality & imageName
  iTag = "<img src='" + d + "' border=0>"
  
  ' if second argument, it's a link
  If UBound(tokens) = 1 Then
    If tokens(1) <> "" Then ' make sure they didn't just leave a trailing commas
      Dim around As String
      around = ""
      around = around + "<a href=" + qq("about:blank") + " id=" + qq(tokens(1)) + ">" + iTag + "</a>"
      iTag = around
    End If
  End If
  
  ' add a NoWiki so that picture-paths don't get interpretted as wiki-words
  iTag = "#NoWiki" + vbCrLf + iTag + vbCrLf + "#Wiki" + vbCrLf
  returnValue = iTag
err1:

  imageInclude = returnValue
End Function



Public Function monthChanges(tokens() As String, model As ModelLevel) As String
  Dim s As String, returnValue As String
  returnValue = err("Sorry, couldn't show a month. Did you specify month and year correctly? ")
  On Error GoTo err1
    s = "=== " + ti.monthName(CInt(tokens(0))) + ", " + tokens(1) + " === " + vbCrLf
    s = s + ti.monthName(tokens(0) - 1) + " : "
    s = s + ti.monthName(tokens(0) + 1) + vbCrLf
    returnValue = s + model.getWikiAnnotatedDataStore.store.timeIndexAsWikiFormat(CInt(tokens(0)), CInt(tokens(1)), False)
err1:
  monthChanges = returnValue
End Function

Public Function CalendarMonth(tokens() As String, model As ModelLevel)
  Dim s As String, returnValue As String
  Dim cm As New CalendarMonth
  returnValue = err("Sorry, couldn't show a month. Did you specify month and year correctly? ")
  On Error GoTo err1
    returnValue = cm.monthAsString(CInt(tokens(0)), CInt(tokens(1)), False, model)
err1:
  CalendarMonth = returnValue
End Function

Public Function IncludingCalendarMonth(tokens() As String, model As ModelLevel)
  Dim s As String, returnValue As String
  Dim cm As New CalendarMonth
  returnValue = err("Sorry, couldn't show a month. Did you specify month and year correctly? ")
  On Error GoTo err1
    returnValue = cm.monthAsString(CInt(tokens(0)), CInt(tokens(1)), True, model)
err1:
  IncludingCalendarMonth = returnValue
End Function

Public Function CalendarEntries(tokens() As String, model As ModelLevel)
    Dim s As String, returnValue As String
    Dim cm As New CalendarMonth
    returnValue = err("Sorry, couldn't show these entries. Did you specify months and years correctly? ")
    On Error GoTo err1
        If UBound(tokens) > 2 Then
            If UBound(tokens) > 3 Then
                If tokens(4) = "back" Then
                    returnValue = cm.includeAllBetween(CInt(tokens(0)), CInt(tokens(1)), CInt(tokens(2)), CInt(tokens(3)), 0, model.getWikiAnnotatedDataStore.store)
                Else
                    returnValue = cm.includeAllBetween(CInt(tokens(0)), CInt(tokens(1)), CInt(tokens(2)), CInt(tokens(3)), 1, model.getWikiAnnotatedDataStore.store)
                End If
            Else
                returnValue = cm.includeAllBetween(CInt(tokens(0)), CInt(tokens(1)), CInt(tokens(2)), CInt(tokens(3)), 1, model.getWikiAnnotatedDataStore.store)
            End If
        End If
err1:
  CalendarEntries = returnValue
End Function


Public Function recentChanges(model As ModelLevel) As String
  Dim d As Date, monthNumber As Integer, yearNumber As Integer, mArg As String
  d = Date
  monthNumber = month(d)
  yearNumber = year(d)
  
  Dim m As String, s As String
  m = ti.monthName(monthNumber)
  
  s = "=== " & m & ", " & yearNumber & " === " + vbCrLf
  s = s + model.getWikiAnnotatedDataStore.store.timeIndexAsWikiFormat(monthNumber, yearNumber, False)
     
  ' now one month earlier
  monthNumber = monthNumber - 1
  
  If monthNumber < 1 Then ' wrap around for before January
    monthNumber = 12
    yearNumber = yearNumber - 1
  End If
  
  mArg = CStr(monthNumber)
  
  m = ti.monthName(monthNumber)
  s = s & vbCrLf & "=== " & m & "-" & yearNumber & " === " + vbCrLf
  s = s + model.getWikiAnnotatedDataStore.store.timeIndexAsWikiFormat(monthNumber, yearNumber, False)
  
  recentChanges = s
End Function

Public Function find(searchString As String, model As ModelLevel) As String
  Dim ps As PageSet
  Set ps = model.getWikiAnnotatedDataStore.store.getPageSetContaining(searchString)
  find = ps.toWikiMarkup
  Set ps = Nothing
End Function


Public Function localResource(linkText As String, path As String, model As ModelLevel) As String
  Dim build As String
  build = "#NoWiki" + vbCrLf
  build = build + "<a target= 'new' id='external' href='file:///" + path + "'>"
  build = build + linkText + "</a>" + vbCrLf + "#Wiki" + vbCrLf
  localResource = build
End Function

Public Function localDir(path As String, model As ModelLevel) As String
  localDir = model.getLocalFileSystem.makeDirectoryPage(path).raw
End Function


Function varsInLine(l As String, p As Page, model As ModelLevel) As String
  Dim parts() As String, build As String
  Dim varName As String, pName As String
  
  build = ""
  parts = Split(l, " ")
  Dim v As Variant, s As String
  For Each v In parts
    s = CStr(v)
    If Left(s, 2) = "$$" Then
      ' it's a variable, either remote ie. $$PageName:VarName
      ' or local ie. $$VarName
      
      If InStr(s, ":") Then
            Dim subParts() As String
            subParts = Split(s, ":")
            
            pName = Right(subParts(0), Len(subParts(0)) - 2)
            varName = subParts(1)
      
            build = build & " " & model.getWikiAnnotatedDataStore.getPageVar(pName, varName)
      Else
            varName = Right(s, Len(s) - 2)
            build = build & " " & p.getVal(varName)
        ' not a var include after all, just some random two $$ thing
      End If
      
    Else
      build = build + " " + s
    End If
      
  Next v
  
  varsInLine = build
End Function


Public Function makeButton(destination, text) As String
    Dim x As String
    x = "#NoWiki" + vbCrLf + "<form action='' method=get><a href='about:blank' id='" + destination + "'>" + _
    "<input type='button' value='" + text + "'></a></form>" + vbCrLf + "#Wiki"
    'MsgBox (x)
    makeButton = x
End Function

Private Sub Class_Initialize()
    Set ti = New TimeIndex
End Sub

Private Sub Class_Terminate()
    Set ti = Nothing
End Sub
