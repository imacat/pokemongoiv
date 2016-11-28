' _3Load: The Pokémon Go IV data
'   by imacat <imacat@mail.imacat.idv.tw>, 2016-11-28

Option Explicit

' subReadDataSheets: Reads the data sheets and shows the data as
'                    OpenOffice Basic arrays
Sub subReadDataSheets
	Dim sOutput as String, mData As Variant
	
	sOutput = "" _
		& "' _2Data: The Pokémon Go IV data" & Chr (10) _
		& "'   by imacat <imacat@mail.imacat.idv.tw>, " & Format (Date (), "yyyy-mm-dd") & Chr (10) _
		& Chr (10) _
		& "Option Explicit"
	sOutput = sOutput & Chr (10) & Chr (10) & fnReadBaseStatsSheet
	sOutput = sOutput & Chr (10) & Chr (10) & fnReadCPMSheet
	sOutput = sOutput & Chr (10) & Chr (10) & fnReadStarDustSheet
	subShowBasicData (sOutput)
End Sub

' subShowBasicData: Shows the data table as Basic arrays
Sub subShowBasicData (sContent As String)
	Dim oDialog As Object, oDialogModel As Object
	Dim oEditModel As Object, oButtonModel As Object
	
	' Creates a dialog
	oDialogModel = CreateUnoService ( _
		"com.sun.star.awt.UnoControlDialogModel")
	oDialogModel.setPropertyValue ("PositionX", 100)
	oDialogModel.setPropertyValue ("PositionY", 100)
	oDialogModel.setPropertyValue ("Height", 130)
	oDialogModel.setPropertyValue ("Width", 200)
	oDialogModel.setPropertyValue ("Title", "Pokémon Go Data")
	
	' Adds the content area
	oEditModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlEditModel")
	oEditModel.setPropertyValue ("PositionX", 5)
	oEditModel.setPropertyValue ("PositionY", 5)
	oEditModel.setPropertyValue ("Height", 100)
	oEditModel.setPropertyValue ("Width", 190)
	oEditModel.setPropertyValue ("MultiLine", True)
	oEditModel.setPropertyValue ("Text", sContent)
	oEditModel.setPropertyValue ("ReadOnly", True)
	oDialogModel.insertByName ("edtContent", oEditModel)
	
	' Adds the OK button.
	oButtonModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlButtonModel")
	oButtonModel.setPropertyValue ("PositionX", 70)
	oButtonModel.setPropertyValue ("PositionY", 110)
	oButtonModel.setPropertyValue ("Height", 15)
	oButtonModel.setPropertyValue ("Width", 60)
	oButtonModel.setPropertyValue ("PushButtonType", _
		com.sun.star.awt.PushButtonType.OK)
	oButtonModel.setPropertyValue ("DefaultButton", True)
	oDialogModel.insertByName ("btnOK", oButtonModel)
	
	' Adds the dialog model to the control and runs it.
	oDialog = CreateUnoService ("com.sun.star.awt.UnoControlDialog")
	oDialog.setModel (oDialogModel)
	oDialog.setVisible (True)
	oDialog.execute
End Sub

' fnReadBaseStatsSheet: Reads the base stats sheet.
Function fnReadBaseStatsSheet As String
	Dim oSheet As Object, oRange As Object, mData As Variant
	Dim nI As Integer, sOutput As String
	Dim nJ As Integer, sEvolveInto As String
	
	oSheet = ThisComponent.getSheets.getByName ("basestat")
	oRange = oSheet.getCellRangeByName ("BaseStats")
	mData = oRange.getDataArray
	
	sOutput = "" _
		& "' fnGetBaseStatsData: Returns the base stats table" & Chr (10) _
		& "Function fnGetBaseStatsData As Variant" & Chr (10) _
		& Chr (9) & "fnGetBaseStatsData = Array( _" & Chr (10)
	For nI = 1 To UBound (mData) - 1
		For nJ = 9 To 7 Step -1
			If mData (nI) (nJ) <> "" Then
				sEvolveInto = mData (nI) (nJ)
				nJ = 6
			End If
		Next nJ
		sOutput = sOutput _
			& Chr (9) & Chr (9) & "Array (""" & mData (nI) (0) _
				& """, """ & mData (nI) (1) _
				& """, " & mData (nI) (3) _
				& ", " & mData (nI) (4) _
				& ", " & mData (nI) (5) _
				& ", """ & sEvolveInto & """), _" & Chr (10)
	Next nI
	nI = UBound (mData)
	For nJ = 9 To 7 Step -1
		If mData (nI) (nJ) <> "" Then
			sEvolveInto = mData (nI) (nJ)
			nJ = 6
		End If
	Next nJ
	sOutput = sOutput _
		& Chr (9) & Chr (9) & "Array (""" & mData (nI) (0) _
			& """, """ & mData (nI) (1) _
			& """, " & mData (nI) (3) _
			& ", " & mData (nI) (4) _
			& ", " & mData (nI) (5) _
			& ", """ & sEvolveInto & """))" & Chr (10) _
		& "End Function"
	fnReadBaseStatsSheet = sOutput
End Function

' fnReadCPMSheet: Reads the combat power multiplier sheet.
Function fnReadCPMSheet As String
	Dim oSheet As Object, oRange As Object, mData As Variant
	Dim nI As Integer, sOutput As String
	
	oSheet = ThisComponent.getSheets.getByName ("cpm")
	oRange = oSheet.getCellRangeByName ("CPM")
	mData = oRange.getDataArray
	
	sOutput = "" _
		& "' fnGetCPMData: Returns the combat power multiplier table" & Chr (10) _
		& "Function fnGetCPMData As Variant" & Chr (10) _
		& Chr (9) & "fnGetCPMData = Array( _" & Chr (10) _
		& Chr (9) & Chr (9) & "-1, _" & Chr (10)
	For nI = 1 To UBound (mData) - 2 Step 2
		sOutput = sOutput _
			& Chr (9) & Chr (9) & mData (nI) (1) & ", _" & Chr (10)
	Next nI
	nI = UBound (mData) - 2
	sOutput = sOutput _
		& Chr (9) & Chr (9) & mData (nI) (1) & ")" & Chr (10) _
		& "End Function"
	fnReadCPMSheet = sOutput
End Function

' fnReadStarDustSheet: Reads the star dust sheet.
Function fnReadStarDustSheet As String
	Dim oSheet As Object, oRange As Object, mData As Variant
	Dim nI As Integer, sOutput As String
	
	oSheet = ThisComponent.getSheets.getByName ("lvup")
	oRange = oSheet.getCellRangeByName ("A2:D81")
	mData = oRange.getDataArray
	
	sOutput = "" _
		& "' fnGetStarDustData: Returns the star dust table" & Chr (10) _
		& "Function fnGetStarDustData As Variant" & Chr (10) _
		& Chr (9) & "fnGetStarDustData = Array( _" & Chr (10) _
		& Chr (9) & Chr (9) & "-1, _" & Chr (10)
	For nI = 1 To UBound (mData) - 1 Step 2
		sOutput = sOutput _
			& Chr (9) & Chr (9) & mData (nI) (2) & ", _" & Chr (10)
	Next nI
	nI = UBound (mData)
	sOutput = sOutput _
		& Chr (9) & Chr (9) & mData (nI) (2) & ")" & Chr (10) _
		& "End Function"
	fnReadStarDustSheet = sOutput
End Function
