Option Strict Off
Option Explicit On
Friend Class VseCanvas
	
	' This is the code which manages a Visual Structure Editor
	' Canvas. Doesn't actually create the canvas (that's given by
	' the main form, but does know about and manipulate it
	
	' has all the handlers for things you click on etc.
	
	Public Canvas As System.Windows.Forms.PictureBox
	Public RawText As System.Windows.Forms.RichTextBox
	Public chef As _PageCooker
    Public mainForm As WADSMainForm ' System.Windows.Forms.Form
	Public nim As NetworkInfoManager
	
	Public Enum vseMode
		View
		Edit
	End Enum
	
	Public mode As vseMode ' is the canvas in view or edit mode
	
	Private Structure vsePoint
		Dim x As Integer
		Dim y As Integer
	End Structure
	
	Private dragFlag As Boolean ' true when drawing an arc
	Private dragStartNode As Short ' which node started from
	Private dragStartPoint As vsePoint ' used when drawing
	Private dragEndPoint As vsePoint ' used when drawing
	
	
	Public Sub init(ByRef p As System.Windows.Forms.PictureBox, ByRef rt As System.Windows.Forms.RichTextBox, ByRef pc As _PageCooker, ByRef mf As System.Windows.Forms.Form)
		Canvas = p
		
		RawText = rt
		chef = pc
		mainForm = mf
		nim = New NetworkInfoManager
		Call nim.init(NodeInfo, ArcInfo, Me)
		mode = vseMode.View
		dragFlag = False
	End Sub
	
	Public Sub setMode(ByRef m As vseMode)
		mode = m
		If m = vseMode.View Then
            Canvas.BackColor = System.Drawing.ColorTranslator.FromOle(RGB(255, 255, 255))
		Else
            Canvas.BackColor = System.Drawing.ColorTranslator.FromOle(RGB(200, 200, 255))
		End If
	End Sub
	
	Public Sub changed(ByRef n As Network)
		n.innerPage.raw = n.spitAsPrettyPersist
		n.parseFromPrettyPersist((n.innerPage.raw))
		Call draw(n, mode)
		RawText.Text = n.spitAsPrettyPersist
		'UPGRADE_ISSUE: Control MagicNotebook could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
		mainForm.MagicNotebook.getSingleUserState.changesSaved = False
	End Sub
	
	Public Sub MouseDownOnCanvas(ByRef n As Network, ByRef Button As Short, ByRef shift As Short, ByRef x As Single, ByRef y As Single)
		Dim p As _Page
		p = n
		
		Dim pn As String ' page name if we're going anywhere else
		Dim newName As String
		
		' detect hitting nodes
		Dim nodeId As Short
		nodeId = n.hitNodeDetect(x, y)
		Dim arcId As Point
		arcId = n.hitArcDetect(x, y)
		
		Dim theNode As Node
		
		Dim theArc As Arc
		If mode = vseMode.View Then
			' in view mode
			
			If Button = 1 Then
				
				If nodeId <> -1 Then
					' clicked on a node
					theNode = n.getNode(nodeId)
					pn = theNode.name
					'UPGRADE_ISSUE: Control controller could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
					Call mainForm.controller.actionLoad(pn, False)
                End If

                If arcId.x <> -1 Then
                    ' clicked on an arc

                    theArc = n.getArc((arcId.x), (arcId.y))

                    pn = theArc.label
                    If pn <> "" Then
                        'UPGRADE_ISSUE: Control controller could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
                        Call mainForm.controller.actionLoad(pn, False)
                    End If
                End If

            Else
                ' switch into edit mode if right click.
                ' this is counter-intuitive but convenient. Which is better?
                'UPGRADE_ISSUE: Control MagicNotebook could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
                'UPGRADE_ISSUE: Control controller could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
                Call mainForm.controller.actionEdit(mainForm.MagicNotebook.getControllableModel.getCurrentPage.pageName, False)
            End If

        Else
            ' edit mode

            If nodeId = -1 Then
                ' haven't hit an existing node,

                ' have we hit an arc?
                If arcId.x = -1 Then
                    ' didn't hit an arc either,
                    'must want to add a new node
                    Call n.addNode(CInt(x), CInt(y), "new node " & CStr(n.nextId))
                    Call changed(n)
                Else
                    ' hit an arc
                    theArc = n.getArc((arcId.x), (arcId.y))
                    If Button = 2 Then
                        Call nim.editAnArc(theArc, n, p)
                    End If
                End If

            Else
                ' hit an existing node
                theNode = n.getNode(nodeId)

                If Button = 1 Then
                    Call startDrag(nodeId)
                Else
                    Call nim.editANode(theNode, n, p)
                End If

            End If
        End If
	End Sub
	
	Public Sub startDrag(ByRef nodeId As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object nodeId. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		dragStartNode = nodeId
		dragFlag = True
	End Sub
	
	Public Sub endDrag(ByRef n As Network, ByRef Button As Short, ByRef shift As Short, ByRef x As Integer, ByRef y As Integer)
		
		Dim p As _Page
		Dim theNode As Node
		Dim nodeIndex As Short
		If dragFlag = True Then
			dragFlag = False
			p = n
			
			
			nodeIndex = n.hitNodeDetect(CInt(x), CInt(y))
			If nodeIndex > -1 Then
				' dragged to another node, so connect them
				If nodeIndex <> dragStartNode Then
					' only connect nodes if they aren't the same
					Call n.connectNodes(dragStartNode, nodeIndex, "", Arc.ArcDirectionality.noDirection)
					Call changed(n)
				End If
			Else
				' dragged to nowhere, let's see if we're moving the node
				theNode = n.getNode(dragStartNode)
				theNode.x = x
				theNode.y = y
				Call changed(n)
			End If
		End If
	End Sub
	
	Public Sub drawArc(ByRef n1 As Node, ByRef n2 As Node, ByRef a As Arc, ByRef m As vseMode)
		'UPGRADE_ISSUE: PictureBox method Canvas.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Canvas.Line (n1.x, n1.y) - (n2.x, n2.y)
        Dim brush1 As Brush = New SolidBrush(Color.Black)
        Dim g As Graphics = Canvas.CreateGraphics
        g.DrawLine(New Pen(brush1), n1.x, n1.y, n2.x, n2.y)



        If m = vseMode.Edit Or a.label <> "" Then

            brush1 = New SolidBrush(System.Drawing.ColorTranslator.FromOle(RGB(200, 255, 255))) 'Color.FromArgb(200, 255, 255))
            'g = Canvas.CreateGraphics
            g.FillEllipse(brush1, a.x, a.y, 10, 10)


            brush1 = New SolidBrush(System.Drawing.ColorTranslator.FromOle(RGB(0, 0, 0)))
            'g = Canvas.CreateGraphics
            g.DrawEllipse(New Pen(brush1), a.x, a.y, 10, 10)




            Dim drawFont As New Font("MS Sans Serif", 8)
            Dim drawBrush As New SolidBrush(Color.Black)
            ' '' Create point for upper-left corner of drawing.
            Dim drawPoint As New PointF(a.x + 5, a.y)
            g.DrawString(a.label, drawFont, drawBrush, drawPoint)
            '' draw label
            ''UPGRADE_ISSUE: PictureBox property Canvas.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '         'Canvas.FillColor = RGB(200, 255, 255)
            '         ''UPGRADE_ISSUE: PictureBox property Canvas.FillStyle was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '         'Canvas.FillStyle = 0
            '         ''UPGRADE_ISSUE: PictureBox method Canvas.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '         'Canvas.Circle (a.x, a.y), 100, RGB(255, 255, 255), 0, 0, 1

            '         brush1 = New SolidBrush(Color.FromArgb(RGB(200, 255, 255)))
            '         g = Canvas.CreateGraphics
            '         g.DrawEllipse(New Pen(brush1), a.x, a.y, 100 * 2, 100 * 2)


            '         'UPGRADE_ISSUE: PictureBox property Canvas.FillStyle was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '         'Canvas.FillStyle = 1
            '         ''UPGRADE_ISSUE: PictureBox method Canvas.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '         'Canvas.Circle (a.x, a.y), 100, RGB(0, 0, 0), 0, 0, 1

            '         brush1 = New SolidBrush(Color.FromArgb(RGB(0, 0, 0)))
            '         g = Canvas.CreateGraphics
            '         g.FillEllipse(brush1, a.x, a.y, 100, 100)



            ''UPGRADE_ISSUE: PictureBox method Canvas.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '         'Canvas.Line (a.x - 180, a.y + 50) - (a.x - 180, a.y + 50)

            '         brush1 = New SolidBrush(Color.FromArgb(RGB(200, 255, 255)))
            '         g = Canvas.CreateGraphics
            '         g.DrawLine(New Pen(brush1), a.x - 180, a.y + 50, a.x - 180, a.y + 50)



            ''UPGRADE_ISSUE: PictureBox method Canvas.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '         'Canvas.Print(a.label)

            '         Dim drawFont As New Font("MS Sans Serif", 8)
            '         Dim drawBrush As New SolidBrush(Color.Black)
            '         ' Create point for upper-left corner of drawing.
            '         Dim drawPoint As New PointF(10, 10)
            '         g.DrawString(a.label, drawFont, drawBrush, drawPoint)



        End If
		
		Dim turt As New Turtle
		Dim i As Short
		Dim apx, apy As Single
		If a.direction = Arc.ArcDirectionality.one Then
			' draw one arrow
			Call turt.setUp(Canvas)
			turt.state = Turtle.penState.down
			turt.collide = Turtle.collideStyle.Ignore
			turt.velocity = 5
			
			apx = (n2.x - a.x) / 2 + a.x
			apy = (n2.y - a.y) / 2 + a.y
			
			If n2.x > n1.x Then
				
				Call turt.setPos(apx, apy)
				Call turt.setAngle(a.angle + turt.pi - turt.pi / 8)
				Call turt.moveForward(40)
				
				Call turt.setPos(apx, apy)
				Call turt.setAngle(a.angle + turt.pi + turt.pi / 8)
				Call turt.moveForward(40)
				
			Else
				
				Call turt.setPos(apx, apy)
				Call turt.setAngle(a.angle - turt.pi / 8)
				Call turt.moveForward(40)
				
				Call turt.setPos(apx, apy)
				Call turt.setAngle(a.angle + turt.pi / 8)
				Call turt.moveForward(40)
				
				
			End If
			
		End If
		
		
	End Sub
	
    Public Sub drawNode(ByRef n As Node, ByRef nt As Network)

        Dim w, h As Integer
        Dim g As Graphics = Canvas.CreateGraphics
        Dim TextFont As New System.Drawing.Font("MS Sans Serif", 8)
        Dim TextSize As New System.Drawing.SizeF
        TextSize = g.MeasureString(n.name, TextFont)
        w = TextSize.Width
        h = TextSize.Height
        Dim brush1 As Brush = New SolidBrush(System.Drawing.ColorTranslator.FromOle(RGB(255, 200, 255)))
        'g.DrawLine(New Pen(brush1), CInt(n.x - w / 2), n.y, CInt(n.x + w / 2), n.y + h)
        g.FillRectangle(brush1, n.x, n.y, w, h)
        Call n.setHitBox(n.x, (n.y), n.x + w, n.y + h)

        brush1 = New SolidBrush(System.Drawing.ColorTranslator.FromOle(RGB(0, 0, 0)))
        'g.DrawLine(New Pen(brush1), CInt(n.x - w / 2), n.y, CInt(n.x + w / 2), n.y + h)
        'g.FillRectangle(brush1, CInt(n.x - w / 2), n.y, CInt(n.x + w / 2), n.y + h)
        g.DrawRectangle(New Pen(brush1), n.x, n.y, w, h)



        Dim drawFont As New Font("MS Sans Serif", 8)
        Dim drawBrush As New SolidBrush(Color.Black)
        ' Create point for upper-left corner of drawing.
        Dim drawPoint As New PointF(n.x, n.y)
        g.DrawString(n.name, drawFont, drawBrush, drawPoint)

        'Dim w, h As Integer
        ''UPGRADE_ISSUE: PictureBox method Canvas.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        '      'w = Canvas.TextWidth(n.name) + nt.drawSize
        '      ''UPGRADE_ISSUE: PictureBox method Canvas.TextHeight was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        '      'h = Canvas.TextHeight(n.name) + nt.drawSize * nt.drawAspect
        '      'Call n.setHitBox(n.x - w / 2, (n.y), n.x + w / 2, n.y + h)
        '      Dim g As Graphics = Canvas.CreateGraphics
        '      Dim TextFont As New System.Drawing.Font("MS Sans Serif", 8)
        '      Dim TextSize As New System.Drawing.SizeF
        '      TextSize = g.MeasureString(n.name, TextFont)
        '      w = TextSize.Width
        '      h = TextSize.Height


        ''UPGRADE_ISSUE: PictureBox property Canvas.FillStyle was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        '      'Canvas.FillStyle = 1
        ''UPGRADE_ISSUE: PictureBox property Canvas.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'

        'Canvas.ForeColor = System.Drawing.ColorTranslator.FromOle(RGB(0, 0, 0))
        '      'UPGRADE_ISSUE: PictureBox property Canvas.FillStyle was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        '      'Canvas.FillColor = RGB(255, 200, 255)
        '      'Canvas.FillStyle = 0
        '      ''Canvas.Circle (n.x, n.y), nt.drawSize, RGB(0, 0, 0), , , nt.drawAspect
        '      ''UPGRADE_ISSUE: PictureBox method Canvas.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        '      'Canvas.Line (n.x - w / 2, n.y) - (n.x + w / 2, n.y + h), RGB(255, 255, 255), BF

        '      Dim brush1 As Brush = New SolidBrush(Color.FromArgb(RGB(255, 255, 255)))
        '      g.DrawLine(New Pen(brush1), CInt(n.x - w / 2), n.y, CInt(n.x + w / 2), n.y + h)



        ''UPGRADE_ISSUE: PictureBox property Canvas.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        '      'Canvas.FillColor = RGB(255, 255, 255)
        '      ''UPGRADE_ISSUE: PictureBox method Canvas.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        '      'Canvas.Line (n.x - w / 2, n.y) - (n.x + w / 2, n.y + h), RGB(0, 0, 0), B

        '      brush1 = New SolidBrush(Color.FromArgb(RGB(0, 0, 0)))
        '      g.DrawLine(New Pen(brush1), CInt(n.x - w / 2), n.y, CInt(n.x + w / 2), n.y + h)

        'Canvas.ForeColor = System.Drawing.ColorTranslator.FromOle(RGB(255, 255, 255))
        ''UPGRADE_ISSUE: PictureBox method Canvas.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        '      'Canvas.Line (n.x - w / 2 + nt.drawSize / 2, n.y + nt.drawSize * nt.drawAspect * 0.5) - (n.x - w / 2 + nt.drawSize / 2, n.y + nt.drawSize * nt.drawAspect * 0.5), RGB(255, 255, 255), B
        '      brush1 = New SolidBrush(Color.FromArgb(RGB(255, 255, 255)))
        '      g.DrawLine(New Pen(brush1), CInt(n.x - w / 2 + nt.drawSize / 2), CInt(n.y + nt.drawSize * nt.drawAspect * 0.5), CInt(n.x - w / 2 + nt.drawSize / 2), CInt(n.y + nt.drawSize * nt.drawAspect * 0.5))



        'Canvas.ForeColor = System.Drawing.ColorTranslator.FromOle(RGB(0, 0, 0))
        ''UPGRADE_ISSUE: PictureBox method Canvas.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        '      'Canvas.Print(n.name)

        '      Dim drawFont As New Font("MS Sans Serif", 8)
        '      Dim drawBrush As New SolidBrush(Color.Black)
        '      ' Create point for upper-left corner of drawing.
        '      Dim drawPoint As New PointF(10, 10)
        '      g.DrawString(n.name, drawFont, drawBrush, drawPoint)

    End Sub
	
	Public Sub draw(ByRef n As Network, ByRef m As vseMode)
		'UPGRADE_ISSUE: PictureBox method Canvas.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Canvas.Cls()
        '  Canvas.CreateGraphics.Clear(Canvas.DefaultBackColor)
        Canvas.CreateGraphics.Clear(System.Drawing.ColorTranslator.FromOle(RGB(255, 255, 255)))
        'Dim drawFont As New Font("MS Sans Serif", 8)
        'Dim drawBrush As New SolidBrush(Color.Black)
        '' Create point for upper-left corner of drawing.
        'Dim drawPoint As New PointF(0, 10)
        'Dim g As Graphics = Canvas.CreateGraphics
        'Dim brush1 As Brush = New SolidBrush(System.Drawing.ColorTranslator.FromOle(RGB(255, 200, 255)))
        ''g.DrawLine(New Pen(brush1), CInt(n.x - w / 2), n.y, CInt(n.x + w / 2), n.y + h)
        'g.FillRectangle(brush1, 0, 0, 100, 30)
        'g.DrawString(DateTime.Now.ToLongTimeString(), drawFont, drawBrush, drawPoint)
		Dim i, j As Short
		For i = 0 To n.noNodes - 1
			For j = 0 To n.noNodes - 1
				If Not n.getArc(i, j) Is Nothing And n.getArc(i, j).exists = True Then
					Call drawArc(n.getNode(i), n.getNode(j), n.getArc(i, j), m)
				End If
			Next j
		Next i
		For i = 0 To n.noNodes - 1
			If Not n.getNode(i) Is Nothing Then
				Call drawNode(n.getNode(i), n)
			End If
		Next i
	End Sub
End Class