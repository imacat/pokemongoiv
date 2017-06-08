' Copyright (c) 2016-2017 imacat.
' 
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
' 
'     http://www.apache.org/licenses/LICENSE-2.0
' 
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.

' 9Load: The Pokémon GO data sheets loader
'   by imacat <imacat@mail.imacat.idv.tw>, 2016-11-28

Option Explicit

' subReadDataSheets: Reads the data sheets and shows the data as
'                    OpenOffice Basic arrays
Sub subReadDataSheets
	Dim sOutput as String, mData As Variant
	
	sOutput = "" _
		& "' Copyright (c) 2016-" & Year (Now) & " imacat." & Chr (10) _
		& "' " & Chr (10) _
		& "' Licensed under the Apache License, Version 2.0 (the ""License"");" & Chr (10) _
		& "' you may not use this file except in compliance with the License." & Chr (10) _
		& "' You may obtain a copy of the License at" & Chr (10) _
		& "' " & Chr (10) _
		& "'     http://www.apache.org/licenses/LICENSE-2.0" & Chr (10) _
		& "' " & Chr (10) _
		& "' Unless required by applicable law or agreed to in writing, software" & Chr (10) _
		& "' distributed under the License is distributed on an ""AS IS"" BASIS," & Chr (10) _
		& "' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied." & Chr (10) _
		& "' See the License for the specific language governing permissions and" & Chr (10) _
		& "' limitations under the License." & Chr (10) _
		& Chr (10) _
		& "' 3Data: The Pokémon GO data for IV calculation" & Chr (10) _
		& "'   by imacat <imacat@mail.imacat.idv.tw>, " & Format (Date (), "yyyy-mm-dd") & Chr (10) _
		& "'   Generated with 9Load.subReadDataSheets ()" & Chr (10) _
		& Chr (10) _
		& "Option Explicit"
	sOutput = sOutput & Chr (10) & Chr (10) & fnReadBaseStatsSheet
	sOutput = sOutput & Chr (10) & Chr (10) & fnReadCPMSheet
	sOutput = sOutput & Chr (10) & Chr (10) & fnReadStardustSheet
	subShowText (sOutput)
End Sub

' subShowChinesePokemonNames: Shows the Chinese names of Pokémons as
'                             resource properties.
Sub subShowChinesePokemonNames
	Dim oSheet As Object, oRange As Object
	Dim mData As Variant, nI As Integer
	Dim sNo As String, sName As String, sNewName As String
	Dim nJ As Integer, sChar As String, nCharCode As Long
	Dim sResult As String

	oSheet = ThisComponent.getSheets.getByName ("basestat")
	oRange = oSheet.getCellRangeByName ("BaseStats")
	mData = oRange.getDataArray
	sResult = ""
	For nI = 1 To UBound (mData)
		sNo = mData (nI) (1)
		sName = mData (nI) (2)
		If sName = "" Then
			sName = mData (nI) (0)
		Else
			sNewName = ""
			For nJ = 1 To Len (sName)
				sChar = Mid (sName, nJ, 1)
				nCharCode = Asc (sChar)
				If nCharCode > 255 Then
					sNewName = sNewName & "\u" & LCase (Hex (nCharCode))
				Else
					sNewName = sNewName & sChar
				End If
			Next nJ
			sName = sNewName
		End If
		sResult = sResult & "1" & sNo & ".lstPokemon.StringItemList=" _
			& sName & Chr (10)
	Next nI
	subShowText (sResult)
End Sub

' subShowText: Shows the text in a text box for copy and paste.
Sub subShowText (sContent As String)
	Dim oDialog As Object, oDialogModel As Object
	Dim oEditModel As Object, oButtonModel As Object
	
	' Creates a dialog
	oDialogModel = CreateUnoService ( _
		"com.sun.star.awt.UnoControlDialogModel")
	oDialogModel.setPropertyValue ("PositionX", 100)
	oDialogModel.setPropertyValue ("PositionY", 100)
	oDialogModel.setPropertyValue ("Height", 130)
	oDialogModel.setPropertyValue ("Width", 200)
	oDialogModel.setPropertyValue ("Title", "Pokémon GO Data")
	
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
	Dim nJ As Integer, nStart As Integer, nEnd As Integer
	Dim sEvolveForms As String
	
	oSheet = ThisComponent.getSheets.getByName ("basestat")
	oRange = oSheet.getCellRangeByName ("BaseStats")
	mData = oRange.getDataArray
	
	sOutput = "" _
		& "' fnGetBaseStatsData: Returns the base stats data." & Chr (10) _
		& "Function fnGetBaseStatsData As Variant" & Chr (10) _
		& Chr (9) & "fnGetBaseStatsData = Array( _" & Chr (10)
	For nI = 1 To UBound (mData) - 1
		sEvolveForms = fnFindEvolveForms (mData (nI))
		sOutput = sOutput _
			& Chr (9) & Chr (9) & "Array (""" _
				& fnMapPokemonNameToId (mData (nI) (0)) _
				& """, """ & mData (nI) (1) _
				& """, " & mData (nI) (3) _
				& ", " & mData (nI) (4) _
				& ", " & mData (nI) (5) _
				& ", " & sEvolveForms & "), _" & Chr (10)
	Next nI
	nI = UBound (mData)
	sEvolveForms = fnFindEvolveForms (mData (nI))
	sOutput = sOutput _
		& Chr (9) & Chr (9) & "Array (""" _
			& fnMapPokemonNameToId (mData (nI) (0)) _
			& """, """ & mData (nI) (1) _
			& """, " & mData (nI) (3) _
			& ", " & mData (nI) (4) _
			& ", " & mData (nI) (5) _
			& ", " & sEvolveForms & "))" & Chr (10) _
		& "End Function"
	fnReadBaseStatsSheet = sOutput
End Function

' fnFindEvolveForms: Finds the evolved forms of the Pokémons.
Function fnFindEvolveForms (mData () As Variant) As String
	Dim nJ As Integer, nStart As Integer, nEnd As Integer
	Dim sEvolveForms As String
	
	If mData (0) = "Eevee" Then
		sEvolveForms = "Array (""Vaporeon"", ""Jolteon"", ""Flareon"")"
	Else
		For nJ = 6 To 8
			If mData (nJ) = mData (0) Then
				nStart = nJ + 1
				nJ = 9
			End If
		Next nJ
		If nStart = 9 Then
			nEnd = 8
		Else
			For nJ = nStart To 8
				If mData (nJ) = "" Then
					nEnd = nJ - 1
					nJ = 9
				Else
					If nJ = 8 Then
						nEnd = 8
						nJ = 9
					End If
				End If
			Next nJ
		End If
		If nEnd = nStart - 1 Then
			sEvolveForms = "Array ()"
		Else
			sEvolveForms = """" _
				& fnMapPokemonNameToId (mData (nStart)) & """"
			For nJ = nStart + 1 To nEnd
				sEvolveForms = sEvolveForms _
					& ", """ _
					& fnMapPokemonNameToId (mData (nJ)) & """"
			Next nJ
			sEvolveForms = "Array (" & sEvolveForms & ")"
		End If
	End If
	fnFindEvolveForms = sEvolveForms
End Function

' fnMapPokemonNameToId: Maps the English Pokémon names to their IDs.
Function fnMapPokemonNameToId (sName As String) As String
	Dim sId As String
	
	sId = ""
	If sName = "Farfetch'd" Then
		sId = "Farfetchd"
	End If
	If sName = "Nidoran♀" Then
		sId = "NidoranFemale"
	End If
	If sName = "Nidoran♂" Then
		sId = "NidoranMale"
	End If
	If sName = "Mr. Mime" Then
		sId = "MrMime"
	End If
	If sName = "Ho-Oh" Then
		sId = "HoOh"
	End If
	If sId = "" Then
		sId = sName
	End If
	fnMapPokemonNameToId = sId
End Function

' fnReadCPMSheet: Reads the combat power multiplier sheet.
Function fnReadCPMSheet As String
	Dim oSheet As Object, oRange As Object, mData As Variant
	Dim nI As Integer, sOutput As String
	
	oSheet = ThisComponent.getSheets.getByName ("cpm")
	oRange = oSheet.getCellRangeByName ("CPM")
	mData = oRange.getDataArray
	
	sOutput = "" _
		& "' fnGetCPMData: Returns the combat power multiplier data." & Chr (10) _
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

' fnReadStardustSheet: Reads the stardust sheet.
Function fnReadStardustSheet As String
	Dim oSheet As Object, oRange As Object, mData As Variant
	Dim nI As Integer, sOutput As String
	
	oSheet = ThisComponent.getSheets.getByName ("lvup")
	oRange = oSheet.getCellRangeByName ("A2:D81")
	mData = oRange.getDataArray
	
	sOutput = "" _
		& "' fnGetStardustData: Returns the stardust data." & Chr (10) _
		& "Function fnGetStardustData As Variant" & Chr (10) _
		& Chr (9) & "fnGetStardustData = Array( _" & Chr (10) _
		& Chr (9) & Chr (9) & "-1, _" & Chr (10)
	For nI = 1 To UBound (mData) - 1 Step 2
		sOutput = sOutput _
			& Chr (9) & Chr (9) & mData (nI) (2) & ", _" & Chr (10)
	Next nI
	nI = UBound (mData)
	sOutput = sOutput _
		& Chr (9) & Chr (9) & mData (nI) (2) & ")" & Chr (10) _
		& "End Function"
	fnReadStardustSheet = sOutput
End Function
