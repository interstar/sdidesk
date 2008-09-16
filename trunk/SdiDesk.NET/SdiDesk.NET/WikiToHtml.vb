Option Strict Off
Option Explicit On
Friend Class WikiToHtml
	
	' More refactoring of page-cooking
	
	' This is a class which knows how to turn WikiMarkup lines into HTML
	
	' here are a couple of things it needs to keep track of
	Private myMg As WikiMarkupGopher ' for doing all the bits and pieces
	
	Public Function isImage(ByRef url As String) As Boolean
		isImage = myMg.isImage(url)
	End Function
	
	
    Public Function lineOfTable(ByRef l2 As String) As String
        Try

        ' wrap this line as a table line
            Dim parts() As String
            Dim b, l As String
            l = l2

            'trim end ||
            If Right(l, 2) = "||" Then
                l = Left(l, Len(l) - 2)
            End If

            ' trim front
            If Left(l, 2) = "||" Then
                l = Right(l, Len(l) - 2)
            End If

            Dim commas As Boolean
            ' is this a comma table or a double piped?
            ' if commas then commas = true
            If InStr(l, ",,") > 0 Then
                commas = True
                parts = Split(l, ",,")
            Else
                commas = False
                parts = Split(l, "||")
            End If

            b = "<tr>"
            Dim s2 As Object
            Dim s As String
            For Each s2 In parts
                'UPGRADE_WARNING: Couldn't resolve default property of object s2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                s = CStr(s2)
                b = b & "<td>" & s & "</td>"
            Next s2
            b = b & "</tr>"
            ' lineOfTable = b
            Return b
        Catch ex As Exception
            MessageBox.Show("Error in LineOfTable function. The error is " & ex.Message)
        End Try
        Return ""
    End Function
	
	Private Function noWikiLine(ByRef l As String) As String
		If l = "" Then
			l = "<br/>" & vbCrLf
		End If
		noWikiLine = l
	End Function
	
	
	Private Function simpleBox(ByRef l As String) As String
		' box ... maybe this notation ain't so hot!
		If Mid(l, 1, 4) = "BOX<" Then
			simpleBox = "<table border=2 cellpadding=3><tr bgcolor=#ffffee><td valign=top>"
			Exit Function
		End If
		
		If Mid(l, 1, 5) = ">BOX<" Then
			simpleBox = "</td><td valign=top>"
			Exit Function
		End If
		
		If Mid(l, 1, 5) = ">BOX>" Then
			simpleBox = "</td></tr><tr><td valign=top>"
			Exit Function
		End If
		
		If Mid(l, 1, 4) = ">BOX" Then
			simpleBox = "</td></tr></table>"
			Exit Function
		End If
		
		simpleBox = l
	End Function
	
	Private Function oneWikiLine(ByRef l2 As String, ByRef tableFlag As Boolean, ByRef preFlag As Boolean, ByRef bulletCount As Short, ByRef lp As _LinkProcessor, ByRef lw As _LinkWrapper) As Object
		Dim l As String
		l = l2
		
		' emphasis
		l = myMg.wrapTags(l, "'''", "<b>", "</b>") ' bold
		l = myMg.wrapTags(l, "''", "<i>", "</i>") ' italic
		
		' headers
		If Left(l, 1) = "=" Then
			l = myMg.wrapTags(l, "======", "<h6>", "</h6>")
			l = myMg.wrapTags(l, "=====", "<h5>", "</h5>")
			l = myMg.wrapTags(l, "====", "<h4>", "</h4>")
			l = myMg.wrapTags(l, "===", "<h3>", "</h3>")
			l = myMg.wrapTags(l, "==", "<h2>", "</h2>")
			l = myMg.wrapTags(l, "=", "<h1>", "</h1>")
		End If
		
		
		l = lp.wrapAllLinks(l, lw)
		
		
		' horizontal lines
		If Mid(l, 1, 4) = "----" Then
			l = "<hr>"
		End If
		
		
		
		'UPGRADE_WARNING: Couldn't resolve default property of object oneWikiLine. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oneWikiLine = l
	End Function
	
	
	Public Function mainTransform(ByRef raw As String, ByRef lp As _LinkProcessor, ByRef lw As _LinkWrapper) As String
		
		' nb : at this point raw is not the raw of a page,
        ' it should have been preprocessed to handle all inlines
        Dim sMainTransform As String = "<html><body>"

        Try

        
            ' MessageBox.Show("Stage 1.01")
            Dim lines() As String
            Dim cooked As String
            Dim l As String
            'Dim l2 As Object

            Dim preFlag, wiki, tableFlag, hide As Boolean

            Dim st As New StringTool

            cooked = "<html><body>"

            'UPGRADE_WARNING: Couldn't resolve default property of object processFootnotes(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            l = processFootnotes(raw, True) ' true is dummy, needed because of decorator. FIX THIS
            ' MessageBox.Show("Stage 1.02 (the value of l is =" & l & ")")

            ''Zebo Code changed to remove last vbcrlf and get 
            '' split of lines
            'lines = Split(l, vbCrLf)
            While Mid(l, l.Length - 1) = vbCrLf
                l = Mid(l, 1, l.Length - 2)
            End While
            'lines = Split(l, vbCrLf)
            If l.IndexOf(vbCrLf) > 0 Then
                'MessageBox.Show(l.IndexOf(vbCrLf))
                lines = Split(l, vbCrLf)
            ElseIf l.IndexOf(vbLf) > 0 Then
                lines = Split(l, vbLf)
            End If


            wiki = True ' in wiki currently
            tableFlag = False ' not in table currently
            preFlag = False ' not in pre mode currently
            hide = False ' not in hide mode

            Dim bulletCount As Short
            bulletCount = 0
            ' used to count indentations of bullets

            'MessageBox.Show("Stage 1.03 (the value of l is =" & l & ")")

            Dim iTotalLine As Integer = 0
            If Not lines Is Nothing Then
                iTotalLine = lines.Count
            End If
            Dim iCurLine As Integer = 0
            Dim noBullets As Short
            Dim bb As Short
            If Not lines Is Nothing Then
                For Each l2 As String In lines
                    'MessageBox.Show("Stage 1.03 (inside loop - the value of l2 is =" & l2 & ") and executing " & iCurLine & " loop of total " & iCurLine)

                    ' If l2 Is Nothing Then Exit For
                    'UPGRADE_WARNING: Couldn't resolve default property of object l2. 
                    'Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    l = CStr(l2) ' change from a variant to real string

                    ' turn off/on wiki
                    If Mid(l, 1, 7) = "#NoWiki" Then
                        wiki = False
                        l = "<pre>"
                    End If

                    If Mid(l, 1, 5) = "#Wiki" Then
                        l = "</pre>"
                        wiki = True
                    End If

                    ' turn off/on hide
                    If Mid(l, 1, 5) = "#Hide" Then
                        hide = True
                        l = ""
                    End If

                    If Mid(l, 1, 7) = "#Unhide" Then
                        l = ""
                        hide = False
                    End If


                    If wiki = True Then

                        ' replace blank lines only in wiki mode
                        If l = "" Then
                            l = "<p />" & vbCrLf
                        End If

                        ' --------------------------
                        ' handle table
                        If tableFlag = True Then
                            ' we are in table
                            ' if the next line is also a table, keep it up,
                            If InStr(l, "____") > 0 Then
                                ' this is a ''table'' with titles,
                                ' ignore this line
                                l = ""
                            Else
                                If (Left(l, 2) = "||" Or (InStr(l, ",,") > 0)) Then
                                    l = lineOfTable(l)
                                Else
                                    ' it's the end of the table
                                    ' so close the table tag
                                    cooked = cooked & "</table>" & vbCrLf
                                    tableFlag = False
                                End If
                            End If
                        Else
                            ' tableFlag is currently false, but
                            ' we need to open a table tag if this is one
                            If (Left(l, 2) = "||" Or (InStr(l, ",,") > 0)) Then
                                ' this is the beginning of a table
                                cooked = cooked & "<table border=1 cellpadding=2 cellspacing=1>" & vbCrLf
                                tableFlag = True
                                l = lineOfTable(l)
                            End If
                        End If


                        ' ----------===========================

                        ' handle pre-mode

                        If preFlag = True Then
                            If Left(l, 1) <> " " Then
                                ' turn off pre-mode
                                preFlag = False
                                cooked = cooked & "</pre>" & vbCrLf
                            End If
                        End If

                        If Left(l, 1) = " " Then
                            ' pre mode, but are we in it already?
                            If preFlag = False Then
                                preFlag = True
                                l = "<pre>" & l & vbCrLf
                            Else
                                l = l & vbCrLf
                                ' do nothing
                            End If
                        End If


                        ' bullets
                        If Left(l, 1) = "*" Then
                            ' the realm of bullets

                            noBullets = 1
                            While Mid(l, noBullets, 1) = "*"
                                noBullets = noBullets + 1
                            End While
                            ' now noBullets should be the character after the bullets
                            ' and equal to the number of bullets
                            ' if this is the same as bulletCount then, fine
                            ' if it is one more, indent
                            ' if one less, outdent

                            If noBullets > bulletCount Then
                                cooked = cooked & "<ul>" & vbCrLf
                                bulletCount = noBullets
                            End If
                            If noBullets < bulletCount Then
                                cooked = cooked & vbCrLf & "</ul>" & vbCrLf
                                bulletCount = noBullets
                            End If
                            If Right(l, 1) <> "*" Then
                                l = "<li>" & Right(l, Len(l) - (noBullets - 1)) & "</li>"
                            Else
                                l = "<li><span style='background-color:#ddddff'>" & st.trimRight(Right(l, Len(l) - (noBullets - 1))) & "</span></li>"
                            End If
                        Else
                            If bulletCount > 0 Then
                                ' we've hit the end of some bullets
                                For bb = 1 To bulletCount - 1
                                    cooked = cooked & vbCrLf & "</ul>"
                                Next bb
                                cooked = cooked & vbCrLf
                                bulletCount = 0
                            End If
                        End If

                        If Mid(l, 1, 1) = ":" Then
                            l = "<dd>" & Right(l, Len(l) - 1) & "</dd>"
                        End If

                        l = simpleBox(l)
                        'UPGRADE_WARNING: Couldn't resolve default property of object oneWikiLine(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        l = oneWikiLine(l, tableFlag, preFlag, bulletCount, lp, lw)

                    Else
                        l = noWikiLine(l)
                    End If

                    ' append to cooked
                    If hide = False Then
                        cooked = cooked & l
                    End If
                Next
            End If
            sMainTransform = cooked & vbCrLf

        Catch ex As Exception
            MessageBox.Show(ex.Message & vbCrLf & vbCrLf & ex.Source & vbCrLf & vbCrLf & ex.ToString())
        End Try
        sMainTransform = sMainTransform & "</body></html>"

        Return sMainTransform
    End Function
	
	
    Public Function processFootnotes(ByRef raw As String, ByRef native As Boolean) As String
        Dim sRetVal As String = ""
        Try
            
            Dim foots As New FootnoteManager
            Dim l As String
            Dim l2 As Object
            Dim lines() As String
            Dim cooked As String
            If InStr(raw, "{{") Then

                Call foots.init()
                cooked = ""

                lines = Split(raw, vbCrLf)
                For Each l2 In lines
                    'UPGRADE_WARNING: Couldn't resolve default property of object l2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    l = CStr(l2)
                    l = foots.extractFootnotes(l, native)
                    cooked = cooked & vbCrLf & l
                Next l2

                cooked = cooked & vbCrLf & "----" & vbCrLf
                cooked = cooked & "==== Footnotes ====" & vbCrLf & "<font size='-1'>" & vbCrLf & foots.getFootnotesAsHtmlString() & "</font>"

                'UPGRADE_NOTE: Object foots may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                foots = Nothing
                'UPGRADE_WARNING: Couldn't resolve default property of object processFootnotes. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                sRetVal = cooked
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object processFootnotes. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                sRetVal = raw
            End If

        Catch ex As Exception
            MessageBox.Show("Error in procesfootnote fucntion and the error is  " & ex.Message)
        End Try
        Return sRetVal
    End Function
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		myMg = New WikiMarkupGopher
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object myMg may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		myMg = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class