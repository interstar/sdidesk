VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExportRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Record of one type of export that the system
' can do

Public name As String ' the name of the export
Public program ' the name of the exporter program
Public paramPage ' the page where the properties of this export are defined

' name,, program,, paramPage
Public Sub init(aName As String, prog As String, pPage As String)
    name = aName
    program = prog
    paramPage = pPage
 End Sub


Public Function toString() As String
   toString = name + ", " + program + ", " + paramPage
End Function

