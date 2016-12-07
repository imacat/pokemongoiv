' 8Registry: Utilities used from other modules to access to PokemonGoIV private configuration
'   Taken from TextToColumns, 2016-12-07

Option Explicit

Const BASE_KEY As String = "/org.openoffice.Office.Addons.PokemonGoIV.AddonConfiguration/"

' fnGetImageUrl: Returns the image URL for the UNO image controls.
Function fnGetImageUrl (sName As String) As String
    Dim oRegKey As Object
	
	oRegKey = fnGetRegistryKeyContent (BASE_KEY & "FileResources/" & sName)
	fnGetImageUrl = fnExpandMacroFieldExpression (oRegKey.Url)
End Function

' fnGetResString: Returns the localized text string.
Function fnGetResString (sID As String) As String
	Dim oRegKey As Object
	
	oRegKey = fnGetRegistryKeyContent (BASE_KEY & "Messages/" & sID)
	fnGetResString = oRegKey.Text
End Function

' fnGetRegistryKeyContent: Returns the registry key content
Function fnGetRegistryKeyContent (sKeyName as string, Optional bforUpdate as Boolean)
    Dim oConfigProvider As Object
    Dim aNodePath (0) As New com.sun.star.beans.PropertyValue
    
	oConfigProvider = createUnoService ( _
	    "com.sun.star.configuration.ConfigurationProvider")
	aNodePath(0).Name = "nodepath"
	aNodePath(0).Value = sKeyName
	If IsMissing (bForUpdate) Then
		fnGetRegistryKeyContent = oConfigProvider.createInstanceWithArguments ( _
			"com.sun.star.configuration.ConfigurationAccess", _
			aNodePath())
	Else
		If bForUpdate Then
			fnGetRegistryKeyContent = oConfigProvider.createInstanceWithArguments ( _
				"com.sun.star.configuration.ConfigurationUpdateAccess", _
				aNodePath())
		Else
			fnGetRegistryKeyContent = oConfigProvider.createInstanceWithArguments ( _
				"com.sun.star.configuration.ConfigurationAccess", _
				aNodePath())
		End If
	End If
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
