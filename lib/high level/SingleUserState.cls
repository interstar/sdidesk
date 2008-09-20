VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SingleUserState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This interface class represents that part of the ModelLevel
' that has to keep track of the interaction state with a user

' Includes access to a NavigationHistory, CurrentPage etc.

' implemented by ModelLevel

Public Enum PageEditState
  LoadedState
  RawState
  EditedState
  PreviewState
  SavedState
End Enum

Public currentPageName As String ' the name of the current page
Public oldPageName As String ' the name of the previous page

Public currentPage As Page ' current page

Public backlinks As Boolean ' do we automatically show backlinks?

Public isLoading As Boolean ' is the page loading (so ignore onChange)
Public changesSaved As Boolean ' record if the changes were saved
Public editState As PageEditState ' the state of editing of this page

Public history As NavigationHistory ' user's nav history

