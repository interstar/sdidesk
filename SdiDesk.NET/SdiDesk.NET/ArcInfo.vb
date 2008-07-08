Option Strict Off
Option Explicit On
Friend Class ArcInfo
	Inherits System.Windows.Forms.Form
	Public manager As NetworkInfoManager
	
	Private Sub CancelButton_Renamed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CancelButton_Renamed.Click
		Me.Hide()
	End Sub
	
	Private Sub OkButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OkButton.Click
		manager.confirmChangesToArc()
	End Sub
End Class