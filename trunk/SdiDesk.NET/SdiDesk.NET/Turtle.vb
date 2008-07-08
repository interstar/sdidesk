Option Strict Off
Option Explicit On
Friend Class Turtle
	
	Const twoPi As Double = 3.14159265 * 2
	
	Public Enum penState
		up = 0
		down = 1
	End Enum
	
	Public Enum collideStyle
		ignore = 0
		die = 1
		rotate = 2
	End Enum
	
	Public x As Single ' x co-ord
	Public y As Single ' y co-ord
	
	Public angle As Single ' angle of movement
	Public velocity As Single ' velocity
	
	Public dx As Single ' velocity in x direction
	Public dy As Single ' velocity in y direction
	
	Public vx As Single ' distance to look ahead x
	Public vy As Single ' distance to look ahead y
	
	Public r As Single ' brushRadius
	
	Public maxAge As Short ' how many time steps before expires
	Public currentAge As Short ' current time steps
	
	Public state As penState ' is pen up or down
	Public colour As Integer ' colour of pen
	
	Public collide As collideStyle ' what to do when something in way
	Public turnAngle As Single ' turn angle
	Public turnState As Boolean ' am I currently turning?
	Public turnStart As Single ' remember which angle the turn started
	
	Public Canvas As System.Windows.Forms.PictureBox ' canvas
	
	Public Sub setUp(ByRef c As System.Windows.Forms.PictureBox)
		Canvas = c
		turnState = False
	End Sub
	
	Public Sub setPos(ByRef ex As Single, ByRef wy As Single)
		x = ex
		y = wy
	End Sub
	
	Public Sub setAngle(ByRef a As Single)
		angle = a
		Call setVel(angleToDx(a, velocity), angleToDy(a, velocity))
		Dim s As String
		s = CStr(dx) & ", " & CStr(dy)
		'  MsgBox (s)
	End Sub
	
	Public Sub turn(ByRef a As Single)
		' MsgBox (angle)
		Call setAngle(angle + a)
	End Sub
	
	Public Sub setVel(ByRef ex As Single, ByRef wy As Single)
		dx = ex
		dy = wy
		vx = ex * 2.3
		vy = wy * 2.3
	End Sub
	
	' trig stuff
	
	Public Function angleToDx(ByRef a As Single, ByRef v As Single) As Single
		angleToDx = System.Math.Cos(a) * v
	End Function
	
	Public Function angleToDy(ByRef a As Single, ByRef v As Single) As Single
		angleToDy = System.Math.Sin(a) * v
	End Function
	
	Public Function oneTouch() As Boolean
		'UPGRADE_ISSUE: PictureBox property Canvas.FillStyle was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Canvas.FillStyle = 0
        ''UPGRADE_ISSUE: PictureBox property Canvas.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Canvas.FillColor = colour
        ''UPGRADE_ISSUE: PictureBox method Canvas.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Canvas.Circle (x, y), r, colour


        Dim brush1 As Brush = New SolidBrush(Color.FromArgb(colour))
        Dim g As Graphics = Canvas.CreateGraphics
        g.FillEllipse(brush1, x, y, r * 2, r * 2)
    End Function

    Public Declare Function GetPixel Lib "gdi32.dll" ( _
       ByVal hdc As IntPtr, _
       ByVal nXPos As Int32, _
       ByVal nYPos As Int32 _
       ) As Int32

    Private Declare Function GetDC Lib "user32.dll" ( _
    ByVal hWnd As IntPtr _
    ) As IntPtr

    Private Declare Function ReleaseDC Lib "user32.dll" ( _
    ByVal hWnd As IntPtr, _
    ByVal hdc As IntPtr _
    ) As Int32

	Public Function testCollide() As Boolean
		' see if a collision
		Dim j As Short
		Dim sx As Single
		Dim sy As Single
		Dim hit As Boolean
		Dim PCol As Integer
		
		hit = False
		For j = 1 To 10
			'     sx = r / 2 * vx + j * vx
			'     sy = r / 2 * vy + j * vy
			'UPGRADE_ISSUE: PictureBox method Canvas.Point was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            'PCol = Canvas.Point(x + vx, y + vy)
            Dim hdc As IntPtr = GetDC(Canvas.Handle)
            PCol = GetPixel(hdc, x + vx, y + vy)
            ReleaseDC(Canvas.Handle, hdc)
			If PCol <> RGB(255, 255, 255) Then
				hit = True
			End If
		Next j
		testCollide = hit
	End Function
	
	Public Function onCollide() As Boolean
		' what to do if collision occurs
		Select Case collide
			Case collideStyle.die
				currentAge = 0
			Case collideStyle.rotate
				If turnState = True Then
					If turnAngle > 0 And angle > turnStart + twoPi Then
						' turned over 360 degrees, no way out so die
						MsgBox(CStr(turnAngle) & ", " & CStr(angle))
						currentAge = 0
					End If
					If turnAngle < 0 And angle < turnStart - twoPi Then
						' turned over 360 degrees, no way out so die
						currentAge = 0
					End If
				Else
					' not currently turning, so start us off
					turnState = True
					turnStart = angle
				End If
				
				Call turn(turnAngle)
				
			Case Else
				turnState = False ' we ain't turning no more, reset flag
				Call normalForward()
		End Select
	End Function
	
	Public Sub normalForward()
		' mark and move on
		If state = penState.down Then
			Call oneTouch()
		End If
		' move on
		x = x + dx
		y = y + dy
	End Sub
	
	Public Sub moveForward(ByRef dist As Short)
		'UPGRADE_ISSUE: PictureBox property Canvas.FillStyle was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Canvas.FillStyle = 0
        ''UPGRADE_ISSUE: PictureBox property Canvas.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Canvas.FillColor = colour
        ''UPGRADE_ISSUE: PictureBox method Canvas.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Canvas.Line (x, y) - (x + (dx * dist), y + dy * dist)

        Dim brush1 As Brush = New SolidBrush(Color.FromArgb(colour))
        Dim g As Graphics = Canvas.CreateGraphics
        g.DrawLine(New Pen(brush1), x, y, x + (dx * dist), y + dy * dist)
		
	End Sub
	
	Public Function nextStep(ByRef finger As Boolean) As Boolean
		' finger is whether mouse button is down or not, true = yes
		' return false if dies
		
		' collide test
		If testCollide() Then
			Call onCollide()
		Else
			Call normalForward()
		End If
		
		If finger = False Then currentAge = currentAge - 1
		
		If currentAge <= 0 Then
			nextStep = False
		Else
			nextStep = True
		End If
		
	End Function
	
	Public Function pi() As Double
		pi = twoPi / 2
	End Function
End Class