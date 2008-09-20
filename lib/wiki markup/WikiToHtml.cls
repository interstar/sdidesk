VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WikiToHtml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' More refactoring of page-cooking

' This is a class which knows how to turn WikiMarkup lines into HTML

' here are a couple of things it needs to keep track of
Private myMg As WikiMarkupGopher ' for doing all the bits and pieces

Public Function isImage(url As String) As Boolean
    isImage = myMg.isImage(url)
End Function


Public Function lineOfTable(l2 As String) As String
  ' wrap this line as a table line
  Dim parts() As String
  Dim b As String, l As String
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
  Dim s2
  For Each s2 In parts
    Dim s As String
    s = CStr(s2)
    b = b + "<td>" + s + "</td>"
  Next s2
  b = b + "</tr>"
  lineOfTable = b
End Function

Private Function noWikiLine(l As String) As String
    If l = "" Then
        l = "<br/>" + vbCrLf
    End If
    noWikiLine = l
End Function


Private Function simpleBox(l As String) As String
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

Private Function oneWikiLine(l2 As String, tableFlag As Boolean, preFlag As Boolean, bulletCount As Integer, lp As LinkProcessor, lw As LinkWrapper)
    Dim l As String
    l = l2
    
    ' emphasis
    l = myMg.wrapTags(l, "'''", "<b>", "</b>")  ' bold
    l = myMg.wrapTags(l, "''", "<i>", "</i>")  ' italic
      
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

    
    
    oneWikiLine = l
End Function


Public Function mainTransform(raw As String, lp As LinkProcessor, lw As LinkWrapper) As String
   
   ' nb : at this point raw is not the raw of a page,
   ' it should have been preprocessed to handle all inlines
   
   Dim lines() As String, cooked As String
   Dim l As String, l2 As Variant
   
   Dim wiki As Boolean, tableFlag As Boolean, preFlag As Boolean, hide As Boolean
   
   Dim st As New StringTool
   
   cooked = "<html><body>"
   
   l = processFootnotes(raw, True)  ' true is dummy, needed because of decorator. FIX THIS
   
   lines = Split(l, vbCrLf)
   
   wiki = True ' in wiki currently
   tableFlag = False ' not in table currently
   preFlag = False ' not in pre mode currently
   hide = False ' not in hide mode
   
   Dim bulletCount As Integer
   bulletCount = 0
   ' used to count indentations of bullets
   
   
   For Each l2 In lines
   
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
           l = "<p />" + vbCrLf
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
                  cooked = cooked + "</table>" + vbCrLf
                  tableFlag = False
               End If
            End If
         Else
            ' tableFlag is currently false, but
            ' we need to open a table tag if this is one
            If (Left(l, 2) = "||" Or (InStr(l, ",,") > 0)) Then
              ' this is the beginning of a table
              cooked = cooked + "<table border=1 cellpadding=2 cellspacing=1>" + vbCrLf
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
              cooked = cooked + "</pre>" + vbCrLf
            End If
         End If
         
         If Left(l, 1) = " " Then
            ' pre mode, but are we in it already?
            If preFlag = False Then
              preFlag = True
              l = "<pre>" + l + vbCrLf
            Else
              l = l + vbCrLf
              ' do nothing
            End If
         End If
         
        
         ' bullets
         If Left(l, 1) = "*" Then
           ' the realm of bullets
           
           Dim noBullets As Integer
           noBullets = 1
           While Mid(l, noBullets, 1) = "*"
              noBullets = noBullets + 1
           Wend
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
                l = "<li>" + Right(l, Len(l) - (noBullets - 1)) + "</li>"
           Else
                l = "<li><span style='background-color:#ddddff'>" & _
                st.trimRight(Right(l, Len(l) - (noBullets - 1))) + "</span></li>"
           End If
         Else
           If bulletCount > 0 Then
             ' we've hit the end of some bullets
             Dim bb As Integer
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
         l = oneWikiLine(l, tableFlag, preFlag, bulletCount, lp, lw)
         
      Else
            l = noWikiLine(l)
      End If
      
      ' append to cooked
      If hide = False Then
        cooked = cooked + l
      End If
   Next l2
   
   mainTransform = cooked & vbCrLf & "</body></html>"
End Function


Public Function processFootnotes(raw As String, native As Boolean)
   If InStr(raw, "{{") Then
        Dim lines() As String, cooked As String
        Dim l As String, l2 As Variant
     
        Dim foots As New FootnoteManager
        Call foots.init
        cooked = ""
        
        lines = Split(raw, vbCrLf)
        For Each l2 In lines
            l = CStr(l2)
            l = foots.extractFootnotes(l, native)
            cooked = cooked + vbCrLf + l
        Next l2

        cooked = cooked + vbCrLf + "----" + vbCrLf
        cooked = cooked + "==== Footnotes ====" + _
        vbCrLf + "<font size='-1'>" + vbCrLf + _
        foots.getFootnotesAsHtmlString() + _
        "</font>"
        
        Set foots = Nothing
        processFootnotes = cooked
   Else
      processFootnotes = raw
   End If
End Function

Private Sub Class_Initialize()
  Set myMg = New WikiMarkupGopher
End Sub

Private Sub Class_Terminate()
  Set myMg = Nothing
End Sub

