VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Arc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' represents an Arc in a network diagram

' the direction an arc goes
Public Enum ArcDirectionality
  noDirection ' undirected
  one ' one-way link
  two ' two-way link
End Enum


Public exists As Boolean ' if there is an arc, true else false
Public label As String ' label
Public direction As ArcDirectionality ' is there a notion of directionality?
Public angle As Single ' what's angle of this arc?
Public x As Long ' arcs need a notional location
Public y As Long ' as a target for hit detection

Public n1 As Integer ' index of from node
Public n2 As Integer ' index of to node


Public Function hit(ex As Long, wy As Long) As Boolean
   ' test if a point is in the active target of an arc
   hit = False
   Dim drawSize As Long, drawAspect As Integer
   drawAspect = 1
   drawSize = 100
   If ex > x - drawSize And ex < x + drawSize Then
     If wy > y - (drawSize * drawAspect) And wy < y + (drawSize * drawAspect) Then
        hit = True
     End If
   End If
End Function



