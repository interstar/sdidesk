VERSION 5.00
Begin VB.Form TestLinkForm 
   Caption         =   "Form1"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "TestLinkForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim l As New Link
    l.external = True
    l.linkType = "normal"
    l.nameSpace = ""
    l.target = "StartPage"
    l.text = "back to start"
    MsgBox (l.toString)
End Sub
