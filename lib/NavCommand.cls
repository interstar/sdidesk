VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NavCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' don't entirely know what this is for yet,
' but intuitively it's going to help me sort out my
' problems with history and stuff

' a command can be one of several things
' 1) PageName
' 2) #raw PageName
' 3) #history PageName
' 4) #delete PageName etc.


Public full As String ' the full text of the command
Public command As String ' the command word
Public pageName As String  ' the argument
Public tail As String ' everything after the command
Private parts() As String ' all parts
Private st As StringTool  ' useful

Public Sub init(s As String)
  Set st = New StringTool
  full = s

  If Left(full, 1) = "#" Then
    ' it's a command
    ' this allows us to use + instead of space as separator, when
    ' useful
        
    full = Replace(full, "+", " ", 1, 1) ' only the first
    
    If InStr(full, " ") > 0 Then
      command = st.strip(st.leftsa(full, " ", 1))
      pageName = st.strip(st.star(full, " ", 1, 1))
      tail = st.strip(st.stripHead(full, " ", 1))
    Else
      command = full
    End If
  Else
    command = "#load"
    pageName = full
  End If
  
End Sub

Public Function getPageName() As String
  getPageName = pageName
End Function

Public Function getCommand() As String
  getCommand = command
End Function

Public Function getFull() As String
  getFull = full
End Function

Public Function getTail() As String
  getTail = tail
End Function


Public Function afterFirstSpace() As String
   ' when searching for things with spaces, need
   ' to return *everything* after first space
   afterFirstSpace = st.stripHead(full, " ", 1)
End Function
