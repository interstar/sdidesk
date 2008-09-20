VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' table class (includes column names)
' and can do some tricks like totalling numeric columns etc

Private headers() As String
Private body() As Variant
Private sums() As String
Private means() As String

Public noRows As Integer
Public noCols As Integer
Public comment As String

Private st As StringTool

Private hasHeaders As Boolean

Public Sub setUp(r As Integer, c As Integer)
  noRows = r
  noCols = c
  ReDim headers(c + 1)
  ReDim body(r + 1, c + 1)
  ReDim sums(c + 1)
  ReDim means(c + 1)
  hasHeaders = False
End Sub

Public Sub putIn(r As Integer, c As Integer, v As Variant)
  body(r, c) = v
End Sub

Public Function at(r As Integer, c As Integer) As Variant
  at = body(r, c)
End Function

Public Sub setHeader(c As Integer, h As String)
  headers(c) = h
End Sub

Public Function atHeader(c As Integer) As String
  atHeader = headers(c)
End Function

Public Function isValidTable(s As String) As Boolean
  Dim b As Boolean
  Dim t As New Table
  isValidTable = t.parseFromDoubleCommaString(s)
  Set t = Nothing
End Function

Public Function parseFromDoubleCommaString(t2 As String) As Boolean

' format is like this

' head,, head,, head
' ____
' body,, body,, body
' body,, body,, body
'
' optional comments

' header is optional, and inferred from the following ____ line
' note the comments must be separated from the end by at least
' one blank line

' returns true if succesful, false if not
  
  Dim lines() As String
  Dim parts() As String
  Dim t As String
  Dim success As Boolean
  Dim i As Integer
  
  t = t2 ' makes sure what we're processing isn't the argument
    
  If InStr(t, (vbCrLf & vbCrLf)) > 0 Then
    ' strip off the comment at the bottom
    i = InStr(t, (vbCrLf & vbCrLf))
    
    comment = Right(t, (Len(t) - i) + 1)
    comment = st.strip(comment)
    comment = st.trimLeft(comment)
    comment = st.trimLeft(comment)
    
    t = st.strip(Left(t, i))
    
  End If
  
  If InStr(t, vbCrLf) Then
  
    lines = Split(t, vbCrLf)
  
    ' first guess at number of rows
    ' though we'll correct if there's a header

    noRows = UBound(lines) + 1
  
    ' now see if the first line contains headers by seeing if the
    ' second line is composed of ____
  
    Dim startRow As Integer
    startRow = 0
    Dim rowCount As Integer
    rowCount = 0
  
  
    If InStr(CStr(lines(1)), "____") > 0 Then
      ' has headers in line 0
      hasHeaders = True
      
      parts = Split(st.strip(CStr(lines(0))), ",,")
      noCols = UBound(parts) + 1
      startRow = 2 ' skip past header lines
      noRows = noRows - 2 ' lose headers and ====
    
      ' now we know enough to redim the arrays
      Call setUp(noRows, noCols)
     
      ' now we can fill the headers
      For i = 0 To noCols - 1
        headers(i) = st.strip(CStr(parts(i)))
      Next i
   
    Else
      startRow = 0 ' no header, so start from top
      ' but still must count cols
      parts = Split(CStr(lines(0)), ",,")
      noCols = UBound(parts) + 1
      Call setUp(noRows, noCols) ' and redim the arrays
    End If
   

    ' now let's read the table body
    For i = 0 To noRows - 1
      On Error GoTo failedRow
          
          parts = Split(CStr(lines(i + startRow)), ",,")
          Dim j As Integer
          For j = 0 To noCols - 1
            If j <= UBound(parts) Then
                body(i, j) = st.strip(CStr(parts(j)))
            Else
                body(i, j) = ""
            End If
          Next j
          rowCount = rowCount + 1
          
failedRow:
    Next i
    
    success = True
  Else
    success = False
  End If
  
  parseFromDoubleCommaString = success
End Function



Public Function isNumeric(col As Integer) As Boolean
' returns true if the column only contains numbers
  Dim i As Integer, flag As Boolean
  flag = False
  On Error GoTo notNumeric
  For i = 0 To noRows
    Dim d As Double
    d = CDbl(at(i, col))
  Next i
  ' if we got here, all in column could be
  ' turned into double, ie. were numeric,
  ' so
  flag = True

notNumeric:
  isNumeric = flag
End Function

Public Function allNumeric() As Boolean
  Dim i As Integer
  Dim flag As Boolean
  flag = True
  For i = 0 To noCols - 1
    If isNumeric(i) Then
      flag = False
    End If
  Next i
  allNumeric = flag
End Function

Public Function calc()
  Dim i As Integer, j As Integer
  For j = 0 To noCols - 1
    If isNumeric(j) Then
      Dim t As Double
      t = 0
      For i = 0 To noRows
         t = t + CDbl(at(i, j))
      Next i
      sums(j) = CStr(t)
      means(j) = CStr(t / noRows)
    Else
      sums(j) = ""
      means(j) = ""
    End If
  Next j

End Function

Public Function rows() As Integer
  rows = noRows
End Function


Public Sub project(t As Table, query As String)
  ' this table becomes a copy of some cols from another table
  ' query = "colNo colNo colNo"
  Dim parts() As String
  parts = Split(query, " ")
  Dim c As Integer
  c = UBound(parts) + 1
  
  ' dimension self as appropriate
  Call setUp(t.rows, c)
  
  Dim p As Variant, cn As Integer
  Dim cc As Integer
  cc = 0
  On Error GoTo endOfQueryLine
  For Each p In parts
    cn = CInt(p)
    headers(cc) = t.atHeader(cn)
    Dim i As Integer
    For i = 0 To noRows
      Call putIn(i, cc, t.at(i, cn))
    Next i
    cc = cc + 1
  Next p
  
endOfQueryLine:
  ' got here when we ran out of columns
  ' is this a good way of handling error?
  
End Sub

Public Function toWikiFormat()
  Dim i As Integer
  Dim j As Integer
  Dim s As String
  Call calc
  s = " ,,"
  For j = 0 To noCols - 1
    s = s + "'''" + CStr(headers(j)) + "''',, "
  Next j
  s = s + vbCrLf + "____" + vbCrLf
  For i = 0 To noRows - 1
    s = s + " ,,"
    For j = 0 To noCols - 1
      If body(i, j) < 0 Then
        s = s + "<font color=#660000>" + body(i, j) + "</font>,, "
      Else
        s = s + body(i, j) + ",, "
      End If
    Next j
    s = s + vbCrLf
  Next i
  s = s + "tot,, "
  For j = 0 To noCols - 1
    s = s + "<font color=#009900>" + sums(j) + "</font>,, "
  Next j
  s = s + vbCrLf + "av.,, "
  For j = 0 To noCols - 1
    s = s + "<font color=#000099>" + means(j) + "</font>,, "
  Next j
  
  
  toWikiFormat = s
End Function


Public Sub inspect()
  MsgBox (toWikiFormat())
End Sub


Public Function spitAsPrettyPersist() As String
    Dim i As Integer, j As Integer
    Dim s As String
    
    s = ""

    For j = 0 To noCols
        If st.strip(CStr(atHeader(j))) <> "" Then
            s = s & atHeader(j) & ",, "
        End If
    Next j
    
    s = st.stripRight(s, " ")
    If Right(s, 2) = ",," Then
            s = st.trimRight(s)
            s = st.trimRight(s)
    End If
    s = s & vbCrLf & "____" & vbCrLf
        
    For i = 0 To noRows
        For j = 0 To noCols
            If st.strip(CStr(at(i, j))) <> "" Then
                s = s & at(i, j) & ",, "
            End If
        Next j
                
        s = st.stripRight(s, " ")
        If Right(s, 2) = ",," Then
            s = st.trimRight(s)
            s = st.trimRight(s)
            s = st.stripRight(s, " ")
        End If
        s = s & vbCrLf
    Next i
    
    s = st.strip(s)
        
    s = s & st.stripLeft(comment, vbCrLf)

    spitAsPrettyPersist = s
End Function

Private Sub Class_Initialize()
    Set st = New StringTool
End Sub

Private Sub Class_Terminate()
    Set st = Nothing
End Sub
