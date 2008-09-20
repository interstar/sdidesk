VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NetworkInfoManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' this class is part of the Visual Structure Editor
' along with the VseCanvas (the class that handles the interactions
' with the diagram of the network) and the two pop-ups : ArcInfo and
' NodeInfo.
' This class is responsible for passing data to and from ArcInfo and
' NodeInfo and arc and node objects themselves. It's used when
' the user wants to edit information about either of these
' network elements.

Public ArcInfoForm As Form
Public NodeInfoForm As Form

Public vc As VseCanvas
Public currentPage As Page

Public net As Network
Public currentArc As Arc
Public currentNode As Node

Public Sub init(nif As Form, aif As Form, vse As VseCanvas)
    Set NodeInfoForm = nif
    Set ArcInfoForm = aif
    Set vc = vse
    
    Set NodeInfoForm.manager = Me
    NodeInfoForm.hide
    
    Set ArcInfoForm.manager = Me
    ArcInfoForm.hide
    
End Sub

Public Sub editAnArc(theArc As Arc, n As Network, p As Page)
    Set currentArc = theArc
    Set net = n
    Set currentPage = p
    
    ArcInfoForm.ArcName.text = currentArc.label
    If currentArc.direction = one Then
        ArcInfoForm.ArcDirectionality.Value = 1
    Else
        ArcInfoForm.ArcDirectionality.Value = 0
    End If
    
    ArcInfoForm.DeleteCheckBox.Value = 0
    
    ArcInfoForm.Visible = True
    ArcInfoForm.Show
    ArcInfoForm.ArcName.SetFocus
End Sub

Public Sub editANode(theNode As Node, n As Network, p As Page)
    Set currentNode = theNode
    Set net = n
    Set currentPage = p
    
    NodeInfoForm.NodeName.text = currentNode.name
    NodeInfoForm.DeleteCheckBox.Value = 0
    
    NodeInfoForm.Visible = True
    NodeInfoForm.Show
    NodeInfoForm.NodeName.SetFocus
End Sub

Public Sub confirmChangesToArc()
    
    If ArcInfoForm.ArcName.text <> "" Then
        currentArc.label = ArcInfoForm.ArcName.text
    End If
    
    If ArcInfoForm.ArcDirectionality.Value = 1 Then
        currentArc.direction = one
    Else
        currentArc.direction = noDirection
    End If
    
    If ArcInfoForm.DeleteCheckBox.Value = 1 Then
        Call net.removeArcByArc(currentArc)
    End If
    
    Call vc.changed(net)
    ArcInfoForm.hide
End Sub

Public Sub confirmChangesToNode()

    If NodeInfoForm.NodeName.text <> "" Then
        currentNode.name = NodeInfoForm.NodeName.text
    End If
    
    If NodeInfoForm.DeleteCheckBox.Value = 1 Then
        Call net.removeNodeByNode(currentNode)
    End If

    Call vc.changed(net)
    NodeInfoForm.hide
End Sub






