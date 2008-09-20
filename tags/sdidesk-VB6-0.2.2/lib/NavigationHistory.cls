VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NavigationHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' where the user browsed, basically a pair of stacks :
' past and future
' when we go to a new page, add the last
' the history
' when we go back, pop it off the history and push it on the future
' when we go forward again, pop it off the future and push it back on the history

Private history As Collection
Private future As Collection
Private buildIndex As Integer ' building up the history
Private futureIndex As Integer ' counts through future
Private walker As Integer
Private lb As ComboBox

Public Sub clear()
  walker = 0
  buildIndex = 0
  futureIndex = 0
  Set history = New Collection
  Set future = New Collection
End Sub


Public Sub setComboBox(cb As ComboBox)
    Set lb = cb
End Sub

Public Sub wipeFuture()
  Set future = New Collection
  futureIndex = 0
End Sub

Public Sub append(pageName As String)
  If buildIndex > 1 Then
    If pageName <> Me.getAtIndex() Then
      Call history.Add(pageName, CStr(buildIndex))
      buildIndex = buildIndex + 1
      Me.inspectInList
    End If
  Else
    Call history.Add(pageName, CStr(buildIndex))
    buildIndex = buildIndex + 1
    Me.inspectInList
  End If
End Sub

Public Function getAtIndex() As String
   Dim i As Integer
   If buildIndex = 0 Then
     i = 0
   Else
     i = buildIndex - 1
   End If
   getAtIndex = history.Item(CStr(i))
End Function



Public Sub back()
   
   If buildIndex > 1 Then
     Dim s As String
     s = history.Item(CStr(buildIndex - 1))
     Call future.Add(s, CStr(futureIndex))
     futureIndex = futureIndex + 1
     history.Remove (CStr(buildIndex - 1))
     buildIndex = buildIndex - 1
   End If
   Me.inspectInList
End Sub


Public Sub forward()
   If future.Count > 0 Then
     Dim s As String
     s = future.Item(CStr(futureIndex - 1))
     Call future.Remove(CStr(futureIndex - 1))
     futureIndex = futureIndex - 1
     Call history.Add(s, CStr(buildIndex))
     buildIndex = buildIndex + 1
   End If
   Me.inspectInList
End Sub

Public Sub printOn(c As Form)
  c.Print "History"
  Dim i As Integer
  For i = 0 To history.Count - 1
    c.Print i, history.Item(CStr(i))
  Next i
  c.Print "Future"
  For i = 0 To future.Count - 1
    c.Print i, future.Item(CStr(i))
  Next i
End Sub

Public Function inspectToString() As String
  Dim s As String
  s = "History" & vbCrLf
  Dim i As Integer
  For i = 0 To history.Count - 1
    s = s + CStr(i) & "  " & history.Item(CStr(i)) & vbCrLf
  Next i
  s = s + "Future" & vbCrLf
  For i = 0 To future.Count - 1
    s = s + CStr(i) & ", " & future.Item(CStr(i)) & vbCrLf
  Next i
  inspectToString = s
End Function

Public Function inspectInList()
    If Not lb Is Nothing Then
        Call lb.clear
        Call lb.AddItem("History")
  
        Dim i As Integer
        For i = 0 To history.Count - 1
            Call lb.AddItem(history.Item(CStr(i)))
        Next i
        Call lb.AddItem("Future")
        For i = future.Count - 1 To 0 Step -1
            Call lb.AddItem(future.Item(CStr(i)))
        Next i
    End If
End Function


Private Sub Class_Initialize()
  Call clear
  Set lb = Nothing
End Sub

Private Sub Class_Terminate()
    Set history = Nothing
    Set future = Nothing
End Sub
