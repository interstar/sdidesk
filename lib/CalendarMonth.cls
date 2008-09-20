VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CalendarMonth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Calendar class month

Dim monthNames
Dim noDays(12) As Integer



Public Function getFirst(year As Integer, month As Integer) As Date
    getFirst = CDate(year & "-" & month & "-1")
End Function

Public Function datePageName(year As Integer, month As Integer, day As Integer) As String
    datePageName = day & "-" & monthName(month - 1) & "-" & year
End Function

Public Function monthName(m As Integer) As String
    monthName = CStr(monthNames(m))
End Function

Private Function nativeLink(link As String, day As Integer) As String
    nativeLink = "[[calendar>" & link & "|" & day & "]]"
End Function


Public Function monthAsString(year As Integer, month As Integer, include As Boolean, model As ModelLevel) As String
    Dim first As Integer
    first = Weekday(getFirst(year, month))
    
    Dim cal(5, 7) As Integer
    Dim i As Integer, j As Integer, dCount As Integer
    Dim s As String, dd As String
    Dim nL As String
    dCount = 1

    s = "== " & monthNames(month - 1) & ", " & year & " ==" & vbCrLf
    s = s & "<font size='2'><table border='1'>"
    
    If include = False Then
        s = s & "<tr bgcolor='#eeffaa'><td>S</td><td>M</td><td>Tu</td><td>We</td><td>Th</td><td>F</td><td>S</td></tr>" & vbCrLf
    Else
        s = s & "<tr bgcolor='#eeffaa'><td>Sunday</td><td>Monday</td><td>Tuesday</td><td>Wedenesday</td><td>Thursday</td><td>Friday</td><td>Saturday</td></tr>" & vbCrLf
    End If
    j = 0
    Do While j < noDays(month) + 1
        j = j + 1
        s = s & "<tr>"
        For i = 1 To 7
            If dCount > 1 Or i >= first Then
                nL = nativeLink(datePageName(year, month, dCount), dCount)
                dd = model.getWikiAnnotatedDataStore.getRawPageData(datePageName(year, month, dCount))
            
                If dd = "new page" Then
                    s = s & "<td bgcolor='#eeeeee'>"
                    dd = ""
                Else
                    s = s & "<td>"
                End If
                s = s & nL
                If include Then
                    s = s & "<br>" & Left(dd, 25)
                End If
                s = s & "</td>" & vbCrLf
                dCount = dCount + 1
             
             Else
                 s = s & "<td bgcolor='#999999'></td>" & vbCrLf
             
             End If
             
             If dCount > noDays(month) Then
                   Exit Do
             End If
             
        Next i
        s = s & "</tr>" & vbCrLf
    Loop
    s = s & vbCrLf & "</table></font>" & vbCrLf
    
    monthAsString = s
End Function


Public Function includedDay(p As String, store As PageStore) As String
    includedDay = "BOX<" & vbCrLf & _
    "===" & p & "===" & _
    vbCrLf & "##Include " & _
    p & vbCrLf & ">BOX" & vbCrLf & "----" & vbCrLf
    
End Function

Public Function includeAllBetween(year1 As Integer, month1 As Integer, year2 As Integer, month2 As Integer, direction As Integer, store As PageStore) As String
    Dim d As Date, s As String, p As String
    s = ""
    If direction = 1 Then
        d = getFirst(year1, month1)
        While d < getFirst(year2, month2)
            p = datePageName(year(d), month(d), day(d))
            If store.pageExists(p) Then
                s = s & includedDay(p, store)
            End If
            d = d + 1
        Wend
    Else
        d = getFirst(year2, month2)
        While d >= getFirst(year1, month1)
            p = datePageName(year(d), month(d), day(d))
            If store.pageExists(p) Then
                s = s & includedDay(p, store)
            End If
            d = d - 1
        Wend
    End If
    includeAllBetween = s
End Function

Private Sub Class_Initialize()
    noDays(1) = 31
    noDays(2) = 28
    noDays(3) = 31
    noDays(4) = 30
    noDays(5) = 31
    noDays(6) = 30
    noDays(7) = 31
    noDays(8) = 31
    noDays(9) = 30
    noDays(10) = 31
    noDays(11) = 30
    noDays(12) = 31
    
    monthNames = Array("Jan", "Feb", "Mar", "Apr", "May", "June", _
    "July", "Aug", "Sep", "Oct", "Nov", "Dec")
End Sub
