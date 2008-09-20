VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimeIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' The purpose of the TimeIndex is to manage a time-based index
' of all pages.
' This is stored in a set of files named after months, in
' directories named years
' ie. 2004\Febuary.mnp

' The format of the month is just a double comma separated
' list ... day number,, page

' Gets read into this time-based index

' ----------------------

Private monthNames

Dim oneMonth As VCollection
Dim store As PageStore

Public Sub init(ps As PageStore)
   Set store = ps
End Sub

Public Sub loadMonthFromStore(year As Integer, month As Integer)
    Dim s As String
    s = store.loadMonth(month, year)
    Call parseMonthFromString(s)
End Sub

Public Sub loadMonthByDate(d As Date)
   Call loadMonthFromStore(year(d), month(d))
End Sub

Public Sub parseMonthFromString(s As String)
  ' expected format is
  ' day,, PageName
  ' day,, PageName etc.
  ' may have two days the same, not a problem
  Set oneMonth = New VCollection
  Dim lines() As String
  lines = Split(s, vbCrLf)
  Dim i As Integer, l As Variant
  i = 0
  For Each l In lines
    If CStr(l) <> "" Then
      Call oneMonth.Add(CStr(l), CStr(l))
    End If
  Next l
End Sub


Public Sub saveMonth(year As Integer, month As Integer)
   Call store.saveMonth(month, year, Me.toRawString)
End Sub


Public Sub saveMonthByDate(d As Date)
  Call saveMonth(CStr(year(d)), CStr(month(d)))
End Sub

Public Function oneDayToWikiString(dayNum As Integer) As String
   Dim build As String
   build = ""
   Dim r As Variant
   For Each r In oneMonth.toCollection
      ' NB : stepping through record entries, NOT days
      If CStr(r) <> "" Then
        Dim parts() As String
        parts = Split(CStr(r), ",, ")
        If CInt(parts(0)) = dayNum Then
           build = build + "** [[" + parts(1) + "]]" + vbCrLf '+ ", "
        End If
      End If
   Next r
   oneDayToWikiString = build
End Function

Public Function monthName(month As Integer) As String
    monthName = monthNames(month - 1)
End Function

Public Function toDate(aYear As Integer, aMonth As Integer, aDay As Integer) As String
      Dim cm As New CalendarMonth
      toDate = "[[" & aDay & "-" & cm.monthName(aMonth - 1) & "-" & aYear & "]]"
      Set cm = Nothing
End Function

Public Function toWikiString(month As Integer, year As Integer, order As Boolean) As String
  ' order = true = 1->31, order=false = 31 -> 1
  Dim build As String, i As Integer, ds As String
  
  build = ""
  If order = True Then
    For i = 1 To 31
      ds = oneDayToWikiString(i)
      If ds <> "" Then
        build = build & "* " & toDate(year, month, i) & " : " & vbCrLf & ds + vbCrLf
      End If
    Next i
  Else
    For i = 31 To 1 Step -1
      ds = oneDayToWikiString(i)
      If ds <> "" Then
        build = build & "* " & toDate(year, month, i) & " : " & vbCrLf & ds + vbCrLf
      End If
    Next i
  End If
  toWikiString = build
End Function

Public Function toRawString() As String
  Dim build As String, i As Variant
  build = ""
  For Each i In oneMonth.toCollection
    build = build + CStr(i) + vbCrLf
  Next i
  toRawString = build
End Function

Public Sub addWord(word As String, dayNum As Integer)
  Dim i As Variant, flag As Boolean
  Dim c As New VCollection
  Dim s As String, dummy As Integer
  
  flag = False
  dummy = 0
  
  For Each i In oneMonth.toCollection
    Dim parts() As String
    parts = Split(CStr(i), ",, ")
    If CInt(parts(0)) < dayNum Then
       Call c.Add(CStr(i), CStr(dummy))
       dummy = dummy + 1
    Else
       If parts(0) = CStr(dayNum) And parts(1) = word Then
         ' do nothing
       Else
       
         If flag = False Then
           ' first on this date
           Call c.Add(dayNum & ",, " & word, CStr(dummy))
           Call c.Add(CStr(i), CStr(dummy + 1))
           flag = True
           dummy = dummy + 2
         Else
           ' now rest
           Call c.Add(CStr(i), CStr(dummy))
           dummy = dummy + 1
         End If
       End If
    End If
  Next i
  If flag = False Then
    Call c.Add(dayNum & ",, " & word, CStr(dummy))
  End If
  Set oneMonth = c
End Sub

Public Sub removeWord(word As String)
  Dim i As Variant
  Dim c As New VCollection
  For Each i In oneMonth.toCollection
    Dim parts() As String
    parts = Split(CStr(i), ",, ")
    If parts(1) <> word Then
      Call c.Add(CStr(i), CStr(i))
    End If
  Next i
  Set oneMonth = c
End Sub

Public Sub updateWord(word As String, oldDate As Date, newDate As Date)
  ' OK, this is what we actually call
  ' we take the word out of the oldDate and put it into newDate
  ' load old date
  Me.loadMonthByDate (oldDate)
  ' remove word
  Me.removeWord (word)
  ' save old date
  Me.saveMonthByDate (oldDate)
  ' load new date
  Me.loadMonthByDate (newDate)
  ' add word
  Call Me.addWord(word, day(newDate))
  ' save new date
  Me.saveMonthByDate (newDate)
End Sub

Private Sub Class_Initialize()
   monthNames = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
   Set oneMonth = New VCollection
End Sub

Private Sub Class_Terminate()
   Set oneMonth = Nothing
   ReDim monthNames(0)
End Sub
