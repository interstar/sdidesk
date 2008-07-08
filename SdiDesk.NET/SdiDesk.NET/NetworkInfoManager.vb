Option Strict Off
Option Explicit On
Friend Class NetworkInfoManager
	
	' this class is part of the Visual Structure Editor
	' along with the VseCanvas (the class that handles the interactions
	' with the diagram of the network) and the two pop-ups : ArcInfo and
	' NodeInfo.
	' This class is responsible for passing data to and from ArcInfo and
	' NodeInfo and arc and node objects themselves. It's used when
	' the user wants to edit information about either of these
	' network elements.
	
    Public ArcInfoForm As ArcInfo ' System.Windows.Forms.Form
    Public NodeInfoForm As NodeInfo ' System.Windows.Forms.Form
	
	Public vc As VseCanvas
	Public currentPage As _Page
	
	Public net As Network
	Public currentArc As Arc
	Public currentNode As Node
	
	Public Sub init(ByRef nif As System.Windows.Forms.Form, ByRef aif As System.Windows.Forms.Form, ByRef vse As VseCanvas)
		NodeInfoForm = nif
		ArcInfoForm = aif
		vc = vse
		
		'UPGRADE_ISSUE: Control manager could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		NodeInfoForm.manager = Me
		NodeInfoForm.Hide()
		
		'UPGRADE_ISSUE: Control manager could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		ArcInfoForm.manager = Me
		ArcInfoForm.Hide()
		
	End Sub
	
	Public Sub editAnArc(ByRef theArc As Arc, ByRef n As Network, ByRef p As _Page)
		currentArc = theArc
		net = n
		currentPage = p
		
		'UPGRADE_ISSUE: Control ArcName could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		ArcInfoForm.ArcName.text = currentArc.label
		If currentArc.direction = Arc.ArcDirectionality.one Then
			'UPGRADE_ISSUE: Control ArcDirectionality could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
            ArcInfoForm.ArcDirectionality.Checked = True '.Value = 1
		Else
			'UPGRADE_ISSUE: Control ArcDirectionality could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
            ArcInfoForm.ArcDirectionality.Checked = False '.Value = 0
		End If
		
		'UPGRADE_ISSUE: Control DeleteCheckBox could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        ArcInfoForm.DeleteCheckBox.Checked = False '.Value = 0
		
		ArcInfoForm.Visible = True
		ArcInfoForm.Show()
		'UPGRADE_ISSUE: Control ArcName could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        ArcInfoForm.ArcName.Focus()
	End Sub
	
	Public Sub editANode(ByRef theNode As Node, ByRef n As Network, ByRef p As _Page)
		currentNode = theNode
		net = n
		currentPage = p
		
		'UPGRADE_ISSUE: Control NodeName could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		NodeInfoForm.NodeName.text = currentNode.name
		'UPGRADE_ISSUE: Control DeleteCheckBox could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        NodeInfoForm.DeleteCheckBox.Checked = False '.Value = 0
		
		NodeInfoForm.Visible = True
		NodeInfoForm.Show()
		'UPGRADE_ISSUE: Control NodeName could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        NodeInfoForm.NodeName.Focus()
	End Sub
	
	Public Sub confirmChangesToArc()
		
		'UPGRADE_ISSUE: Control ArcName could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		If ArcInfoForm.ArcName.text <> "" Then
			'UPGRADE_ISSUE: Control ArcName could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			currentArc.label = ArcInfoForm.ArcName.text
		End If
		
		'UPGRADE_ISSUE: Control ArcDirectionality could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        If ArcInfoForm.ArcDirectionality.Checked = True Then '.Value = 1 Then
            currentArc.direction = Arc.ArcDirectionality.one
        Else
            currentArc.direction = Arc.ArcDirectionality.noDirection
        End If
		
		'UPGRADE_ISSUE: Control DeleteCheckBox could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        If ArcInfoForm.DeleteCheckBox.Checked = True Then '.Value = 1 Then
            Call net.removeArcByArc(currentArc)
        End If
		
		Call vc.changed(net)
		ArcInfoForm.Hide()
	End Sub
	
	Public Sub confirmChangesToNode()
		
		'UPGRADE_ISSUE: Control NodeName could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		If NodeInfoForm.NodeName.text <> "" Then
			'UPGRADE_ISSUE: Control NodeName could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			currentNode.name = NodeInfoForm.NodeName.text
		End If
		
		'UPGRADE_ISSUE: Control DeleteCheckBox could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        If NodeInfoForm.DeleteCheckBox.Checked = True Then '.Value = 1 Then
            Call net.removeNodeByNode(currentNode)
        End If
		
		Call vc.changed(net)
		NodeInfoForm.Hide()
	End Sub
End Class