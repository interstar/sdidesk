VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PagePreparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' this module now takes the pre-processing / preparation of a page from
' the PageCooker.

Public wads As WikiAnnotatedDataStore ' needs to know about the model


Public Function processInlines(raw As String, aPage As Page) As String
   Dim lines() As String, l2 As Variant, l As String
   Dim build As String, current As String  ' what we build, and current state
   Dim ScriptEngine As New ScriptEngine
      
   build = ""
   current = raw ' this time, at end of loop it will be set to build
   
   ' we do this in a loop because it's recursive
   Dim finished As Boolean
   
   Do
      build = ""
      finished = True
      lines() = Split(current, vbCrLf)
      
      ' now through each line of the page, substituting the inlines
      
      For Each l2 In lines
        l = CStr(l2) ' ensure it's a string
        
        ' interpret #= assignments
        If InStr(l, "#=") > 0 Then
          Dim parts() As String
          parts = Split(l, "#=")
          Call aPage.setVal(parts(0), parts(1))
          l = vbCrLf + "<font size=+1 color=#339999> " + parts(0) + " #''''''= " + parts(1) + "</font>"
        End If
        
        ' interpret $$ variables
        If InStr(l, "$$") > 0 Then
          l = ScriptEngine.varsInLine(l, aPage, wads)
        End If
        
        ' interpret ##Inlines
        If Left(l, 2) = "##" Then
          l = ScriptEngine.perform(l, wads)
          finished = False ' this is dirty, try it again
        End If
               
        build = build + l + vbCrLf
      Next l2
        
      ' at this point we did all the inlines in the current state
      ' but we want to make sure that if any inlines brought in NEW inlines
      ' they get processed too.
      ' so set current = build and maybe go round again
      
      current = build
   Loop Until finished = True ' loop until there are no more ##inlines left
      
   ' now current should have all inlines fully substituted
   processInlines = current
   
   ' clean up
   Set ScriptEngine = Nothing

End Function


Public Function prepare(aPage As Page, backlinks As Boolean) As String
    Dim s As String
    s = Me.processInlines(aPage.raw, aPage)
    
    If backlinks = True Then
        Dim ps As PageSet
        Set ps = wads.getPageSetContaining(aPage.pageName)
        s = s & "----" & vbCrLf & "<table bgcolor=#ccffcc width=100%><tr>" & vbCrLf
        s = s & "<td> <h3>Backlinks</h3> " & vbCrLf & ps.toWikiMarkup & "</td></tr></table>" + vbCrLf
    End If
    aPage.prepared = Me.processInlines(s, aPage)
    prepare = aPage.prepared
End Function


