VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Node"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' data about a node object for networks

Public name As String
Public x As Long
Public y As Long
Public nodeType As String

Public boxLeft As Long
Public boxRight As Long
Public boxTop As Long
Public boxBottom As Long

Public Sub init(ex As Long, wy As Long, n As String)
  x = ex
  y = wy
  nodeType = ""
  name = n
End Sub

Public Sub setType(s As String)
  nodeType = s
End Sub

Public Sub setHitBox(l As Long, t As Long, r As Long, b As Long)
    boxLeft = l
    boxTop = t
    boxBottom = b
    boxRight = r
End Sub

Public Function boxToString() As String
    boxToString = "(" & boxLeft & "," & boxTop & ")-" & _
    "(" & boxRight & "," & boxBottom & ")"
End Function


Public Function hit(x As Long, y As Long)
   hit = False
   If x > boxLeft And x < boxRight And y > boxTop And y < boxBottom Then
        hit = True
   End If
End Function

