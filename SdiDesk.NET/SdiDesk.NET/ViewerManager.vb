Option Strict Off
Option Explicit On
Imports MSFlexGridLib
Imports Microsoft.VisualBasic.Compatibility
Friend Class ViewerManager
	
	' this object manages the visibility of the various viewers available,
	' and arranges which will be visible, invisible.
	
    Public TableEditor As AxMSFlexGridLib.AxMSFlexGrid '.AxMSFlexGrid
	Public NetworkCanvas As System.Windows.Forms.PictureBox
	Public HtmlView As System.Windows.Forms.WebBrowser
	Public RawText As System.Windows.Forms.RichTextBox
	
	Public Enum ViewerManagerMode
		vmmNetwork
		vmmTable
		vmmHtml
		vmmRaw
	End Enum
	
	Public mode As ViewerManagerMode
	
    Public Sub init(ByRef raw As System.Windows.Forms.RichTextBox, ByRef html As System.Windows.Forms.WebBrowser, ByRef net As System.Windows.Forms.PictureBox, ByRef table As AxMSFlexGridLib.AxMSFlexGrid)
        RawText = raw
        HtmlView = html
        NetworkCanvas = net
        TableEditor = table
    End Sub
	
	Public Sub hideAll()
		RawText.Visible = False
		HtmlView.Visible = False
		NetworkCanvas.Visible = False
		TableEditor.Visible = False
	End Sub
	
	Public Sub showRaw()
		Call hideAll()
		RawText.Visible = True
		RawText.Focus()
		mode = ViewerManagerMode.vmmRaw
	End Sub
	
	Public Sub showHtml()
		Call hideAll()
		HtmlView.Visible = True
		'HtmlView.SetFocus (doesn't seem to work)
		mode = ViewerManagerMode.vmmHtml
	End Sub
	
	Public Sub showVse()
		Call hideAll()
		NetworkCanvas.Visible = True
		NetworkCanvas.Focus()
		mode = ViewerManagerMode.vmmNetwork
	End Sub
	
	Public Sub showTable()
		Call hideAll()
		TableEditor.Visible = True
		TableEditor.Focus()
		mode = ViewerManagerMode.vmmTable
	End Sub
	
	Public Sub resize(ByRef vfWidth As Short, ByRef vfHeight As Short, ByRef dfb As Short)
		
		HtmlView.Left = VB6.TwipsToPixelsX(60)
		HtmlView.Width = VB6.TwipsToPixelsX((vfWidth - 140))
		HtmlView.Height = VB6.TwipsToPixelsY(vfHeight - dfb)
		
		RawText.Left = VB6.TwipsToPixelsX(60)
		RawText.Width = VB6.TwipsToPixelsX(vfWidth - 140)
		RawText.Height = VB6.TwipsToPixelsY(vfHeight - dfb)
		
		NetworkCanvas.Left = VB6.TwipsToPixelsX(60)
		NetworkCanvas.Width = VB6.TwipsToPixelsX(vfWidth - 140)
		NetworkCanvas.Height = VB6.TwipsToPixelsY(vfHeight - dfb)
		
		TableEditor.Left = VB6.TwipsToPixelsX(60)
		TableEditor.Width = VB6.TwipsToPixelsX(vfWidth - 140)
		TableEditor.Height = VB6.TwipsToPixelsY(vfHeight - dfb)
		
	End Sub
End Class