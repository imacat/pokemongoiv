' 7Registry: Utilities used from other modules to access to PokemonGoIV private configuration
'   Taken from TextToColumns, 2016-12-07

Option Explicit

Const BASE_KEY As String = "/org.openoffice.Office.Addons.PokemonGoIV.AddonConfiguration/"

' fnGetImageUrl: Returns the image URL for the UNO image controls.
Function fnGetImageUrl (sName As String) As String
	BasicLibraries.loadLibrary "Tools"
	Dim oRegKey As Object
	
	oRegKey = GetRegistryKeyContent (BASE_KEY & "FileResources/" & sName)
	fnGetImageUrl = fnExpandMacroFieldExpression (oRegKey.Url)
End Function

' fnGetResString: Returns the localized text string.
Function fnGetResString (sID As String) As String
	BasicLibraries.loadLibrary "Tools"
	Dim oRegKey As Object
	
	oRegKey = GetRegistryKeyContent (BASE_KEY & "Messages/" & sID)
	fnGetResString = oRegKey.Text
End Function

' fnExpandMacroFieldExpression
Function fnExpandMacroFieldExpression (sURL As String) As String
	Dim sTemp As String
	Dim oSM As Object
	Dim oMacroExpander As Object
	
	' Gets the service manager
	oSM = getProcessServiceManager
	' Gets the macro expander
	oMacroExpander = oSM.DefaultContext.getValueByName ( _
		"/singletons/com.sun.star.util.theMacroExpander")
	
	'cut the vnd.sun.star.expand: part
	sTemp = Join (Split (sURL, "vnd.sun.star.expand:"))
	
	'Expand the macrofield expression
	sTemp = oMacroExpander.ExpandMacros (sTemp)
	sTemp = Trim (sTemp)
	fnExpandMacroFieldExpression = sTemp
End Function
