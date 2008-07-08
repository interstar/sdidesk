Option Strict Off
Option Explicit On
Interface _ModelLevel
    Function getWikiAnnotatedDataStore() As _WikiAnnotatedDataStore
    Function getSingleUserState() As _SingleUserState
    Function getSystemConfigurations() As _SystemConfigurations
    Function getCrawlerSubsystem() As _CrawlerSubsystem
    Function getExportSubsystem() As _ExportSubsystem
    Function getControllableModel() As _ControllableModel
    Function getLocalFileSystem() As _LocalFileSystem
    Sub setCallBackForm(ByRef f As System.Windows.Forms.Form)
    Sub setPagePreparer(ByRef pp As PagePreparer)
    Sub setPageCooker(ByRef pc As _PageCooker)
    Sub setForm(ByRef f As System.Windows.Forms.Form)
End Interface
Friend Class ModelLevel
	Implements _ModelLevel
	
	' Now the ModelLevel is the interface to that part of
	' the program that implements the model part (in the MVC sense)
	' of SdiDesk
	
	' now (as of March 2005) the model level contains several sub-components
	
	' Currently
	' WikiAnnotatedDataStore
	' SingleUserState
	' SystemConfiguration
	' CrawlerSubsystem
	' ExportSubsystem
	' ControllableModel
	
	' how are we going to refactor this?
	
	' we are going to try to break out the separate modules
	' for testing in programs *without* ModelLayer,
	' and we are going to add whatever functionality we need
	' to test them, to the WikiAnnotatedDataStore etc and we will have
	' fake WADS, SysConfs etc. for testing
	
	
	Public Function getWikiAnnotatedDataStore() As _WikiAnnotatedDataStore Implements _ModelLevel.getWikiAnnotatedDataStore
	End Function
	
	Public Function getSingleUserState() As _SingleUserState Implements _ModelLevel.getSingleUserState
	End Function
	
	Public Function getSystemConfigurations() As _SystemConfigurations Implements _ModelLevel.getSystemConfigurations
	End Function
	
	Public Function getCrawlerSubsystem() As _CrawlerSubsystem Implements _ModelLevel.getCrawlerSubsystem
	End Function
	
	Public Function getExportSubsystem() As _ExportSubsystem Implements _ModelLevel.getExportSubsystem
	End Function
	
	Public Function getControllableModel() As _ControllableModel Implements _ModelLevel.getControllableModel
	End Function
	
	Public Function getLocalFileSystem() As _LocalFileSystem Implements _ModelLevel.getLocalFileSystem
	End Function
	
	Public Sub setCallBackForm(ByRef f As System.Windows.Forms.Form) Implements _ModelLevel.setCallBackForm
	End Sub
	
	Public Sub setPagePreparer(ByRef pp As PagePreparer) Implements _ModelLevel.setPagePreparer
	End Sub
	
	Public Sub setPageCooker(ByRef pc As _PageCooker) Implements _ModelLevel.setPageCooker
	End Sub
	
	Public Sub setForm(ByRef f As System.Windows.Forms.Form) Implements _ModelLevel.setForm
	End Sub
End Class