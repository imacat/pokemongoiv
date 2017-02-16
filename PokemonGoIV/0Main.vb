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

' 0Main: The main module for the Pokémon GO IV calculator
'   by imacat <imacat@mail.imacat.idv.tw>, 2016-11-27

Option Explicit

' The base stats of a Pokémon.
Type aStats
	sNo As String
	sPokemon As String
	nStamina As Integer
	nAttack As Integer
	nDefense As Integer
	mEvolved () As String
End Type

' The individual values of a Pokémon.
Type aIV
	fLevel As Double
	nStamina As Integer
	nAttack As Integer
	nDefense As Integer
	nTotal As Integer
	nMaxCP As Integer
	maEvolved () As aEvolvedStats
End Type

' The calculated evolved stats of a Pokémon.
Type aEvolvedStats
	nCP As Integer
	nMaxCP As Integer
End Type

' The parameters to find the individual values.
Type aFindIVParam
	sPokemon As String
	nCP As Integer
	nHP As Integer
	nStardust As Integer
	nPlayerLevel As Integer
	bIsNew As Boolean
	nTotal As Integer
	sBest As String
	nMax As Integer
	bIsCancelled As Boolean
End Type

Private maBaseStats () As New aStats
Private mCPM () As Double, mStardust () As Integer

' subMain: The main program
Sub subMain
	Dim aBaseStats As New aStats, maIVs As Variant, nI As Integer
	Dim aQuery As New aFindIVParam
	
	aQuery = fnAskParam
	If aQuery.bIsCancelled Then
		Exit Sub
	End If
	aBaseStats = fnGetBaseStats (aQuery.sPokemon)
	maIVs = fnFindIV (aBaseStats, aQuery)
	If UBound (maIVs) = -1 Then
		MsgBox fnGetResString ("ErrorNotFound")
	Else
		subSaveIV (aBaseStats, aQuery, maIVs)
	End If
End Sub

' fnAskParam: Asks the users for the parameters for the Pokémon.
Function fnAskParam As aFindIVParam
	Dim oDialog As Object
	Dim oList As Object, mPokemons () As String, nI As Integer
	Dim bIsBestAttack As Boolean, bIsBestDefense As Boolean
	Dim bIsBestHP As Boolean
	Dim aQuery As New aFindIVParam
	
	DialogLibraries.loadLibrary "PokemonGoIV"
	oDialog = CreateUnoDialog (DialogLibraries.PokemonGoIV.DlgMain)
	oDialog.getControl ("lstTotal").setVisible (False)
	oDialog.getControl ("txtBestBefore").setVisible (False)
	oDialog.getControl ("lstBest").setVisible (False)
	oDialog.getControl ("txtBestAfter").setVisible (False)
	oDialog.getControl ("cbxBest2").setVisible (False)
	oDialog.getControl ("cbxBest3").setVisible (False)
	oDialog.getControl ("lstMax").setVisible (False)
	
	oDialog.getControl ("imgPokemon").getModel.setPropertyValue ( _
		"ImageURL", fnGetImageUrl ("Unknown"))
	oDialog.getControl ("imgTeamLogo").getModel.setPropertyValue ( _
		"ImageURL", fnGetImageUrl ("Unknown"))
	
	If oDialog.execute = 0 Then
		aQuery.bIsCancelled = True
		fnAskParam = aQuery
		Exit Function
	End If
	Xray oDialog.getControl ("lstPokemon")
	
	With aQuery
		.sPokemon = oDialog.getControl ("lstPokemon").getSelectedItem
		.nCP = oDialog.getControl ("numCP").getValue
		.nHP = oDialog.getControl ("numHP").getValue
		.nStardust = CInt (oDialog.getControl ("lstStardust").getSelectedItem)
		.nPlayerLevel = CInt (oDialog.getControl ("lstPlayerLevel").getSelectedItem)
		.nTotal = oDialog.getControl ("lstTotal").getSelectedItemPos + 1
		.nMax = oDialog.getControl ("lstMax").getSelectedItemPos + 1
		.bIsCancelled = False
	End With
	If oDialog.getControl ("cbxIsNew").getState = 1 Then
		aQuery.bIsNew = True
	Else
		aQuery.bIsNew = False
	End If
	
	' The best stats
	bIsBestAttack = False
	bIsBestDefense = False
	bIsBestHP = False
	If oDialog.getControl ("lstBest").getSelectedItemPos = 0 Then
		bIsBestAttack = True
		If oDialog.getControl ("cbxBest2").getState = 1 Then
			bIsBestDefense = True
		End If
		If oDialog.getControl ("cbxBest3").getState = 1 Then
			bIsBestHP = True
		End If
	End If
	If oDialog.getControl ("lstBest").getSelectedItemPos = 1 Then
		bIsBestDefense = True
		If oDialog.getControl ("cbxBest2").getState = 1 Then
			bIsBestAttack = True
		End If
		If oDialog.getControl ("cbxBest3").getState = 1 Then
			bIsBestHP = True
		End If
	End If
	If oDialog.getControl ("lstBest").getSelectedItemPos = 2 Then
		bIsBestHP = True
		If oDialog.getControl ("cbxBest2").getState = 1 Then
			bIsBestAttack = True
		End If
		If oDialog.getControl ("cbxBest3").getState = 1 Then
			bIsBestDefense = True
		End If
	End If
	aQuery.sBest = ""
	If bIsBestAttack Then
		aQuery.sBest = aQuery.sBest & "Atk "
	End If
	If bIsBestDefense Then
		aQuery.sBest = aQuery.sBest & "Def "
	End If
	If bIsBestHP Then
		aQuery.sBest = aQuery.sBest & "Sta "
	End If
	
	fnAskParam = aQuery
End Function

' subBtnOKCheck: Checks whether the required columns are filled.
Sub subBtnOKCheck (oEvent As object)
	Dim oDialog As Object
	Dim oPokemon As Object, oCP As Object
	Dim oHP As Object, oStardust As Object, oOK As Object
	
	oDialog = oEvent.Source.getContext
	oPokemon = oDialog.getControl ("lstPokemon")
	oCP = oDialog.getControl ("numCP")
	oHP = oDialog.getControl ("numHP")
	oStardust = oDialog.getControl ("lstStardust")
	oOK = oDialog.getControl ("btnOK")
	
	If oPokemon.getSelectedItemPos <> -1 _
			And oCP.getText <> "" _
			And oHP.getText <> "" _
			And oStardust.getSelectedItemPos <> -1 Then
		oOK.setEnable (True)
	Else
		oOK.setEnable (False)
	End If
End Sub

' subLstPokemonSelected: When the Pokémon is selected.
Sub subLstPokemonSelected (oEvent As object)
	Dim oDialog As Object, nSelected As Integer
	Dim oImageModel As Object, sImageId As String
	
	oDialog = oEvent.Source.getContext
	
	' Updates the Pokémon image.
	nSelected = oDialog.getControl ("lstPokemon").getSelectedItemPos
	subReadBaseStats
	sImageId = "Pokemon" & maBaseStats (nSelected).sPokemon
	oImageModel = oDialog.getControl ("imgPokemon").getModel
	oImageModel.setPropertyValue ("ImageURL", _
		fnGetImageUrl (sImageId))
	
	' Updates the text of the stats total appraisal.
	subUpdateTotalAppraisal (oDialog, True)
	' Checks if the required columns are filled.
	subBtnOKCheck (oEvent)
End Sub

' subRdoTeamValorItemChanged: When Team Valor is selected.
Sub subRdoTeamValorItemChanged (oEvent As object)
	Dim oDialog As Object, oList As Object, oText As Object
	Dim oImageModel As Object
	Dim mItems () As String
	
	oDialog = oEvent.Source.getContext
	
	oImageModel = oDialog.getControl ("imgTeamLogo").getModel
	oImageModel.setPropertyValue ("ImageURL", _
		fnGetImageUrl ("TeamLogoValor"))
	oImageModel = oDialog.getControl ("imgTeamLeader").getModel
	oImageModel.setPropertyValue ("ImageURL", _
		fnGetImageUrl ("TeamLeaderCandela"))
	
	oText = oDialog.getControl ("txtLeaderAppraise")
	oText.setVisible (True)
	oText.setText (fnGetResString ("AppraiseFromCandela"))
	
	' Updates the text of the stats total appraisal.
	subUpdateTotalAppraisal (oDialog, False)
	
	' Updates the text of the best stat appraisal.
	subUpdateBestStatAppraisal (oDialog, _
		fnGetResString ("AppraisalValorBest"))
	
	mItems = Array ( _
		fnGetResString ("AppraisalValorMax15"), _
		fnGetResString ("AppraisalValorMax13Or14"), _
		fnGetResString ("AppraisalValorMax8To12"), _
		fnGetResString ("AppraisalValorMaxUpTo7"))
	oList = oDialog.getControl ("lstMax")
	oList.removeItems (0, oList.getItemCount())
	oList.addItems (mItems, 0)
	oList.setVisible (True)
End Sub

' subRdoTeamMysticItemChanged: When Team Mystic is selected.
Sub subRdoTeamMysticItemChanged (oEvent As object)
	Dim oDialog As Object, oList As Object, oText As Object
	Dim oImageModel As Object
	Dim mItems () As String
	
	oDialog = oEvent.Source.getContext
	
	oImageModel = oDialog.getControl ("imgTeamLogo").getModel
	oImageModel.setPropertyValue ("ImageURL", _
		fnGetImageUrl ("TeamLogoMystic"))
	oImageModel = oDialog.getControl ("imgTeamLeader").getModel
	oImageModel.setPropertyValue ("ImageURL", _
		fnGetImageUrl ("TeamLeaderBlanche"))
	
	oText = oDialog.getControl ("txtLeaderAppraise")
	oText.setVisible (True)
	oText.setText (fnGetResString ("AppraiseFromBlanche"))
	
	' Updates the text of the stats total appraisal.
	subUpdateTotalAppraisal (oDialog, False)
	
	' Updates the text of the best stat appraisal.
	subUpdateBestStatAppraisal (oDialog, _
		fnGetResString ("AppraisalMysticBest"))
	
	mItems = Array ( _
		fnGetResString ("AppraisalMysticMax15"), _
		fnGetResString ("AppraisalMysticMax13Or14"), _
		fnGetResString ("AppraisalMysticMax8To12"), _
		fnGetResString ("AppraisalMysticMaxUpTo7"))
	oList = oDialog.getControl ("lstMax")
	oList.removeItems (0, oList.getItemCount())
	oList.addItems (mItems, 0)
	oList.setVisible (True)
End Sub

' subRdoTeamInstinctItemChanged: When Team Instinct is selected.
Sub subRdoTeamInstinctItemChanged (oEvent As object)
	Dim oDialog As Object, oList As Object, oText As Object
	Dim oImageModel As Object
	Dim mItems () As String
	
	oDialog = oEvent.Source.getContext
	
	oImageModel = oDialog.getControl ("imgTeamLogo").getModel
	oImageModel.setPropertyValue ("ImageURL", _
		fnGetImageUrl ("TeamLogoInstinct"))
	oImageModel = oDialog.getControl ("imgTeamLeader").getModel
	oImageModel.setPropertyValue ("ImageURL", _
		fnGetImageUrl ("TeamLeaderSpark"))
	
	oText = oDialog.getControl ("txtLeaderAppraise")
	oText.setVisible (True)
	oText.setText (fnGetResString ("AppraiseFromSpark"))
	
	' Updates the text of the stats total appraisal.
	subUpdateTotalAppraisal (oDialog, False)
	
	' Updates the text of the best stat appraisal.
	subUpdateBestStatAppraisal (oDialog, _
		fnGetResString ("AppraisalInstinctBest"))
	
	mItems = Array ( _
		fnGetResString ("AppraisalInstinctMax15"), _
		fnGetResString ("AppraisalInstinctMax13Or14"), _
		fnGetResString ("AppraisalInstinctMax8To12"), _
		fnGetResString ("AppraisalInstinctMaxUpTo7"))
	oList = oDialog.getControl ("lstMax")
	oList.removeItems (0, oList.getItemCount())
	oList.addItems (mItems, 0)
	oList.setVisible (True)
End Sub

' subUpdateBestStatAppraisal: Updates the text of the best stat appraisal.
Sub subUpdateBestStatAppraisal (oDialog As Object, sAppraisal)
	Dim oText As Object, oList As Object, nX As Integer
	Dim sBefore As String, nBeforeWidth As Integer
	Dim sAfter As String, nAfterWidth As Integer
	Dim nPos As Integer
	Dim mItems () As String
	
	nPos = InStr (sAppraisal, "[Stat]")
	sBefore = Left (sAppraisal, nPos - 1)
	If Right (sBefore, 1) <> " " Then
	    sBefore = sBefore & " "
	End If
	nBeforeWidth = CInt (Len (sBefore) * 2.3)
	sAfter = Right (sAppraisal, _
	    Len (sAppraisal) - nPos - Len ("[Stat]") + 1)
	If Left (sAfter, 1) <> " " Then
	    sAfter = " " & sAfter
	End If
	nAfterWidth = CInt (Len (sAfter) * 2.3)
	
	oText = oDialog.getControl ("txtBestBefore")
	oText.getModel.setPropertyValue ("Width", nBeforeWidth)
	oText.setVisible (True)
	oText.setText (sBefore)
	nX = oText.getModel.getPropertyValue ("PositionX") + nBeforeWidth
	
	mItems = Array ( _
		fnGetResString ("StatAttack"), _
		fnGetResString ("StatDefense"), _
		fnGetResString ("StatHP"))
	oList = oDialog.getControl ("lstBest")
	oList.removeItems (0, oList.getItemCount())
	oList.addItems (mItems, 0)
	oList.getModel.setPropertyValue ("PositionX", nX)
	oList.setVisible (True)
	nX = nX + oList.getModel.getPropertyValue ("Width") + 2
	
	oText = oDialog.getControl ("txtBestAfter")
	oText.getModel.setPropertyValue ("PositionX", nX)
	oText.getModel.setPropertyValue ("Width", nAfterWidth)
	oText.setVisible (True)
	oText.setText (sAfter)
	
	oList = oDialog.getControl ("cbxBest2")
	oList.setVisible (False)
	
	oList = oDialog.getControl ("cbxBest3")
	oList.setVisible (False)
End Sub

' subLstBestItemChanged: When the best stat is selected.
Sub subLstBestItemChanged (oEvent As object)
	Dim oDialog As Object, oCheckBox As Object, sBestToo As String
	
	oDialog = oEvent.Source.getContext
	If oDialog.getControl ("rdoTeamValor").getState Then
		sBestToo = fnGetResString ("AppraisalValorBestToo")
		If oDialog.getControl ("lstBest").getSelectedItemPos = 0 Then
			oCheckBox = oDialog.getControl ("cbxBest2")
			oCheckBox.setLabel (fnReplace ( _
				sBestToo, "[Stat]", fnGetResString ("StatDefense")))
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
			oCheckBox = oDialog.getControl ("cbxBest3")
			oCheckBox.setLabel (fnReplace ( _
				sBestToo, "[Stat]", fnGetResString ("StatHP")))
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
		End If
		If oDialog.getControl ("lstBest").getSelectedItemPos = 1 Then
			oCheckBox = oDialog.getControl ("cbxBest2")
			oCheckBox.setLabel (fnReplace ( _
				sBestToo, "[Stat]", fnGetResString ("StatAttack")))
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
			oCheckBox = oDialog.getControl ("cbxBest3")
			oCheckBox.setLabel (fnReplace ( _
				sBestToo, "[Stat]", fnGetResString ("StatHP")))
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
		End If
		If oDialog.getControl ("lstBest").getSelectedItemPos = 2 Then
			oCheckBox = oDialog.getControl ("cbxBest2")
			oCheckBox.setLabel (fnReplace ( _
				sBestToo, "[Stat]", fnGetResString ("StatAttack")))
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
			oCheckBox = oDialog.getControl ("cbxBest3")
			oCheckBox.setLabel (fnReplace ( _
				sBestToo, "[Stat]", fnGetResString ("StatDefense")))
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
		End If
	End If
	If oDialog.getControl ("rdoTeamMystic").getState Then
		sBestToo = fnGetResString ("AppraisalMysticBestToo")
		If oDialog.getControl ("lstBest").getSelectedItemPos = 0 Then
			oCheckBox = oDialog.getControl ("cbxBest2")
			oCheckBox.setLabel (fnReplace ( _
				sBestToo, "[Stat]", fnGetResString ("StatDefense")))
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
			oCheckBox = oDialog.getControl ("cbxBest3")
			oCheckBox.setLabel (fnReplace ( _
				sBestToo, "[Stat]", fnGetResString ("StatHP")))
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
		End If
		If oDialog.getControl ("lstBest").getSelectedItemPos = 1 Then
			oCheckBox = oDialog.getControl ("cbxBest2")
			oCheckBox.setLabel (fnReplace ( _
				sBestToo, "[Stat]", fnGetResString ("StatAttack")))
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
			oCheckBox = oDialog.getControl ("cbxBest3")
			oCheckBox.setLabel (fnReplace ( _
				sBestToo, "[Stat]", fnGetResString ("StatHP")))
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
		End If
		If oDialog.getControl ("lstBest").getSelectedItemPos = 2 Then
			oCheckBox = oDialog.getControl ("cbxBest2")
			oCheckBox.setLabel (fnReplace ( _
				sBestToo, "[Stat]", fnGetResString ("StatAttack")))
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
			oCheckBox = oDialog.getControl ("cbxBest3")
			oCheckBox.setLabel (fnReplace ( _
				sBestToo, "[Stat]", fnGetResString ("StatDefense")))
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
		End If
	End If
	If oDialog.getControl ("rdoTeamInstinct").getState Then
		sBestToo = fnGetResString ("AppraisalInstinctBestToo")
		If oDialog.getControl ("lstBest").getSelectedItemPos = 0 Then
			oCheckBox = oDialog.getControl ("cbxBest2")
			oCheckBox.setLabel (fnReplace ( _
				sBestToo, "[Stat]", fnGetResString ("StatDefense")))
			oCheckBox.setVisible (True)
			oCheckBox = oDialog.getControl ("cbxBest3")
			oCheckBox.setLabel (fnReplace ( _
				sBestToo, "[Stat]", fnGetResString ("StatHP")))
			oCheckBox.setVisible (True)
		End If
		If oDialog.getControl ("lstBest").getSelectedItemPos = 1 Then
			oCheckBox = oDialog.getControl ("cbxBest2")
			oCheckBox.setLabel (fnReplace ( _
				sBestToo, "[Stat]", fnGetResString ("StatAttack")))
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
			oCheckBox = oDialog.getControl ("cbxBest3")
			oCheckBox.setLabel (fnReplace ( _
				sBestToo, "[Stat]", fnGetResString ("StatHP")))
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
		End If
		If oDialog.getControl ("lstBest").getSelectedItemPos = 2 Then
			oCheckBox = oDialog.getControl ("cbxBest2")
			oCheckBox.setLabel (fnReplace ( _
				sBestToo, "[Stat]", fnGetResString ("StatAttack")))
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
			oCheckBox = oDialog.getControl ("cbxBest3")
			oCheckBox.setLabel (fnReplace ( _
				sBestToo, "[Stat]", fnGetResString ("StatDefense")))
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
		End If
	End If
End Sub

' subUpdateTotalAppraisal: Updates the text of the stats total
'						   appraisal.
Sub subUpdateTotalAppraisal ( _
		oDialog As Object, bIsKeepSelected As Boolean)
	Dim sPokemon As String, oList As Object, nSelected As Integer
	Dim mItems () As String, nI As Integer
	
	If oDialog.getControl ("rdoTeamValor").getState Then
		mItems = Array ( _
			fnGetResString ("AppraisalValorTotal37OrHigher"), _
			fnGetResString ("AppraisalValorTotal30To36"), _
			fnGetResString ("AppraisalValorTotal23To29"), _
			fnGetResString ("AppraisalValorTotalUpTo22"))
	End If
	If oDialog.getControl ("rdoTeamMystic").getState Then
		mItems = Array ( _
			fnGetResString ("AppraisalMysticTotal37OrHigher"), _
			fnGetResString ("AppraisalMysticTotal30To36"), _
			fnGetResString ("AppraisalMysticTotal23To29"), _
			fnGetResString ("AppraisalMysticTotalUpTo22"))
	End If
	If oDialog.getControl ("rdoTeamInstinct").getState Then
		mItems = Array ( _
			fnGetResString ("AppraisalInstinctTotal37OrHigher"), _
			fnGetResString ("AppraisalInstinctTotal30To36"), _
			fnGetResString ("AppraisalInstinctTotal23To29"), _
			fnGetResString ("AppraisalInstinctTotalUpTo22"))
	End If
	' The team was not selected yet.
	If UBound (mItems) = -1 Then
		Exit sub
	End If
	
	sPokemon = oDialog.getControl ("lstPokemon").getSelectedItem
	If sPokemon <> "" Then
		For nI = 0 To UBound (mItems)
			mItems (nI) = fnReplace (mItems (nI), _
				"[Pokémon]", sPokemon)
		Next nI
	End If
	
	oList = oDialog.getControl ("lstTotal")
	If bIsKeepSelected Then
		nSelected = oList.getSelectedItemPos
	End If
	oList.removeItems (0, oList.getItemCount())
	oList.addItems (mItems, 0)
	If bIsKeepSelected Then
		oList.selectItemPos (nSelected, True)
	End If
	oList.setVisible (True)
End Sub

' fnFindIV: Finds the possible individual values of the Pokémon
Function fnFindIV ( _
		aBaseStats As aStats, aQuery As aFindIVParam) As Variant
	Dim nEvolved As Integer
	Dim maEvBaseStats () As New aStats, aTempStats As New aStats
	Dim maIV () As New aIV, aTempIV As New aIV
	Dim fLevel As Double, nStamina As Integer
	Dim nAttack As Integer, nDefense As integer
	Dim nI As Integer, nJ As Integer
	Dim fStep As Double, nN As Integer
	
	If aQuery.sPokemon = "" Then
		fnFindIV = maIV
		Exit Function
	End If
	If aQuery.bIsNew Then
		fStep = 1
	Else
		fStep = 0.5
	End If
	subReadStardust
	nEvolved = UBound (aBaseStats.mEvolved)
	If nEvolved > -1 Then
		ReDim Preserve maEvBaseStats (nEvolved) As New aStats
		For nI = 0 To nEvolved
			aTempStats = fnGetBaseStats (aBaseStats.mEvolved (nI))
			With maEvBaseStats (nI)
				.nAttack = aTempStats.nAttack
				.nDefense = aTempStats.nDefense
				.nStamina = aTempStats.nStamina
			End With
		Next nI
	End If
	nN = -1
	For fLevel = 1 To UBound (mStardust) Step fStep
		If mStardust (CInt (fLevel - 0.5)) = aQuery.nStardust Then
			For nStamina = 0 To 15
				If fnCalcHP (aBaseStats, fLevel, nStamina) = aQuery.nHP Then
					For nAttack = 0 To 15
						For nDefense = 0 To 15
							If fnCalcCP (aBaseStats, fLevel, nAttack, nDefense, nStamina) = aQuery.nCP _
									And Not fnFilterAppraisals (aQuery, nAttack, nDefense, nStamina) Then
								nN = nN + 1
								ReDim Preserve maIV (nN) As New aIV
								With maIV (nN)
									.fLevel = fLevel
									.nAttack = nAttack
									.nDefense = nDefense
									.nStamina = nStamina
									.nTotal = nAttack _
										+ nDefense + nStamina
								End With
								If aQuery.nPlayerLevel <> 0 Then
									maIV (nN).nMaxCP = fnCalcCP ( _
										aBaseStats, _
										aQuery.nPlayerLevel + 1.5, _
										nAttack, nDefense, nStamina)
								End If
								maIV (nN).maEvolved _
									= fnGetEvolvedArray (nEvolved)
								For nI = 0 To nEvolved
									maIV (nN).maEvolved (nI).nCP _
										= fnCalcCP ( _
											maEvBaseStats (nI), _
											fLevel, nAttack, _
											nDefense, nStamina)
									If aQuery.nPlayerLevel <> 0 Then
										maIV (nN).maEvolved (nI).nMaxCP = fnCalcCP (maEvBaseStats (nI), aQuery.nPlayerLevel + 1.5, nAttack, nDefense, nStamina)
									End If
								Next nI
							End If
						Next nDefense
					Next nAttack
				End If
			Next nStamina
		End If
	Next fLevel
	' Sorts the IVs
	For nI = 0 To UBound (maIV) - 1
		For nJ = nI + 1 To UBound (maIV)
			If fnCompareIV (maIV (nI), maIV (nJ)) > 0 Then
				' This is an array of data.  The data are actually
				' allocated in sequences.  maIV (nI) is not a
				' reference.  They cannot simply be assigned.
				subCopyIV (maIV (nI), aTempIV)
				subCopyIV (maIV (nJ), maIV (nI))
				subCopyIV (aTempIV, maIV (nJ))
			End If
		Next nJ
	Next nI
	fnFindIV = maIV
End Function

' fnCompareIV: Compare two IVs for sorting
Function fnCompareIV (aIVa As aIV, aIVb As aIV) As Double
	Dim nCPa As Integer, nCPb As Integer, nI As Integer
	
	nCPa = aIVa.nMaxCP
	For nI = 0 To UBound (aIVa.maEvolved)
		If nCPa < aIVa.maEvolved (nI).nMaxCP Then
			nCPa = aIVa.maEvolved (nI).nMaxCP
		End If
	Next nI
	nCPb = aIVb.nMaxCP
	For nI = 0 To UBound (aIVb.maEvolved)
		If nCPb < aIVb.maEvolved (nI).nMaxCP Then
			nCPb = aIVb.maEvolved (nI).nMaxCP
		End If
	Next nI
	fnCompareIV = nCPb - nCPa
	If fnCompareIV <> 0 Then
		Exit Function
	End If
	
	nCPa = 0
	For nI = 0 To UBound (aIVa.maEvolved)
		If nCPa < aIVa.maEvolved (nI).nCP Then
			nCPa = aIVa.maEvolved (nI).nCP
		End If
	Next nI
	nCPb = 0
	For nI = 0 To UBound (aIVb.maEvolved)
		If nCPb < aIVb.maEvolved (nI).nCP Then
			nCPb = aIVb.maEvolved (nI).nCP
		End If
	Next nI
	fnCompareIV = nCPb - nCPa
	If fnCompareIV <> 0 Then
		Exit Function
	End If
	
	fnCompareIV = aIVb.nTotal - aIVa.nTotal
	If fnCompareIV <> 0 Then
		Exit Function
	End If
	fnCompareIV = aIVb.fLevel - aIVa.fLevel
	If fnCompareIV <> 0 Then
		Exit Function
	End If
	fnCompareIV = aIVb.nStamina - aIVa.nStamina
	If fnCompareIV <> 0 Then
		Exit Function
	End If
	fnCompareIV = aIVb.nAttack - aIVa.nAttack
	If fnCompareIV <> 0 Then
		Exit Function
	End If
	fnCompareIV = aIVb.nDefense - aIVa.nDefense
	If fnCompareIV <> 0 Then
		Exit Function
	End If
End Function

' subCopyIV: Copies one IV to another
Function subCopyIV (aFrom As aIV, aTo As aIV) As Double
	Dim nI As Integer
	
	With aTo
		.nAttack = aFrom.nAttack
		.nDefense = aFrom.nDefense
		.nStamina = aFrom.nStamina
		.nTotal = aFrom.nTotal
		.nMaxCP = aFrom.nMaxCP
	End With
	aTo.maEvolved = fnGetEvolvedArray (UBound (aFrom.maEvolved))
	For nI = 0 To UBound (aFrom.maEvolved)
		With aTo.maEvolved (nI)
			.nCP = aFrom.maEvolved (nI).nCP
			.nMaxCP = aFrom.maEvolved (nI).nMaxCP
		End With
	Next nI
End Function

' subSaveIV: Saves the found IV
Sub subSaveIV ( _
		aBaseStats As aStats, aQuery As aFindIVParam, maIVs () As aIV)
	Dim oDoc As Object, oSheet As Object
	Dim oRange As Object, oColumns As Object, oRows As Object
	Dim nI As Integer, nJ As Integer, nFront As Integer
	Dim nEvolved As Integer
	Dim mData (Ubound (maIVs) + 1) As Variant, mRow () As Variant
	Dim mProps () As New com.sun.star.beans.PropertyValue
	
	oDoc = StarDesktop.loadComponentFromURL ( _
		"private:factory/scalc", "_default", 0, mProps)
	oSheet = oDoc.getSheets.getByIndex (0)
	nEvolved = UBound (maIVs (0).maEvolved) + 1
	
	mRow = Array ( _
		"No", "Pokemon", "CP", "HP", "Stardust", _
		"Lv", "Atk", "Def", "Sta", "IV")
	nFront = UBound (mRow)
	If aQuery.sPokemon = "Eevee" Then
		If aQuery.nPlayerLevel <> 0 Then
			ReDim Preserve mRow (nFront + 6) As Variant
			mRow (nFront + 1) = "CP as " _
				& aBaseStats.mEvolved (0)
			mRow (nFront + 2) = "Powered-up as " _
				& aBaseStats.mEvolved (0)
			mRow (nFront + 3) = "CP as " _
				& aBaseStats.mEvolved (1)
			mRow (nFront + 4) = "Powered-up as " _
				& aBaseStats.mEvolved (1)
			mRow (nFront + 5) = "CP as " _
				& aBaseStats.mEvolved (2)
			mRow (nFront + 6) = "Powered-up as " _
				& aBaseStats.mEvolved (2)
		Else
			ReDim Preserve mRow (nFront + 3) As Variant
			mRow (nFront + 1) = "CP as " _
				& aBaseStats.mEvolved (0)
			mRow (nFront + 2) = "CP as " _
				& aBaseStats.mEvolved (1)
			mRow (nFront + 3) = "CP as " _
				& aBaseStats.mEvolved (2)
		End If
	Else
		If nEvolved = 0 Then
			If aQuery.nPlayerLevel <> 0 Then
				ReDim Preserve mRow (nFront + 1) As Variant
				mRow (nFront + 1) = "Powered-up"
			End If
		Else
			If aQuery.nPlayerLevel <> 0 Then
				ReDim Preserve mRow (nFront + nEvolved + 1) As Variant
				For nJ = 0 To nEvolved - 1
					mRow (nFront + nJ + 1) = "CP as " _
						& aBaseStats.mEvolved (nJ)
				Next nJ
				mRow (UBound (mRow)) = "Powered-up as " _
					& aBaseStats.mEvolved (nEvolved - 1)
			Else
				ReDim Preserve mRow (nFront + nEvolved) As Variant
				For nJ = 0 To nEvolved - 1
					mRow (nFront + nJ + 1) = "CP as " _
						& aBaseStats.mEvolved (nJ)
				Next nJ
			End If
		End If
	End If
	mData (0) = mRow
	
	For nI = 0 To UBound (maIVs)
		mRow = Array ( _
			"", "", "", "", "", _
			maIVs (nI).fLevel, maIVs (nI).nAttack, _
			maIVs (nI).nDefense, maIVs (nI).nStamina, _
			maIVs (nI).nTotal / 45)
		If aQuery.sPokemon = "Eevee" Then
			If aQuery.nPlayerLevel <> 0 Then
				ReDim Preserve mRow (nFront + 6) As Variant
				mRow (nFront + 1) = maIVs (nI).maEvolved (0).nCP
				mRow (nFront + 2) = maIVs (nI).maEvolved (0).nMaxCP
				mRow (nFront + 3) = maIVs (nI).maEvolved (1).nCP
				mRow (nFront + 4) = maIVs (nI).maEvolved (1).nMaxCP
				mRow (nFront + 5) = maIVs (nI).maEvolved (2).nCP
				mRow (nFront + 6) = maIVs (nI).maEvolved (2).nMaxCP
			Else
				ReDim Preserve mRow (nFront + 3) As Variant
				mRow (nFront + 1) = maIVs (nI).maEvolved (0).nCP
				mRow (nFront + 2) = maIVs (nI).maEvolved (1).nCP
				mRow (nFront + 3) = maIVs (nI).maEvolved (2).nCP
			End If
		Else
			If nEvolved = 0 Then
				If aQuery.nPlayerLevel <> 0 Then
					ReDim Preserve mRow (nFront + 1) As Variant
					mRow (nFront + 1) = maIVs (nI).nMaxCP
				End If
			Else
				If aQuery.nPlayerLevel <> 0 Then
					ReDim Preserve mRow (nFront + nEvolved + 1) As Variant
					For nJ = 0 To nEvolved - 1
						mRow (nFront + nJ + 1) = maIVs (nI).maEvolved (nJ).nCP
					Next nJ
					mRow (UBound (mRow)) = maIVs (nI).maEvolved (nEvolved - 1).nMaxCP
				Else
					ReDim Preserve mRow (nFront + nEvolved) As Variant
					For nJ = 0 To nEvolved - 1
						mRow (nFront + nJ + 1) = maIVs (nI).maEvolved (nJ).nCP
					Next nJ
				End If
			End If
		End If
		mData (nI + 1) = mRow
	Next nI
	
	' Fills the query information at the first row
	mData (1) (0) = aBaseStats.sNo
	mData (1) (1) = aQuery.sPokemon
	mData (1) (2) = aQuery.nCP
	mData (1) (3) = aQuery.nHP
	mData (1) (4) = aQuery.nStardust
	
	oRange = oSheet.getCellRangeByPosition ( _
		0, 0, UBound (mData (0)), UBound (mData))
	oRange.setDataArray (mData)
	oRange.setPropertyValue ("VertJustify", _
		com.sun.star.table.CellVertJustify.TOP)
	
	oRange = oSheet.getCellRangeByPosition ( _
		0, 1, 0, UBound (mData))
	oRange.merge (True)
	oRange = oSheet.getCellRangeByPosition ( _
		1, 1, 1, UBound (mData))
	oRange.merge (True)
	oRange = oSheet.getCellRangeByPosition ( _
		2, 1, 2, UBound (mData))
	oRange.merge (True)
	oRange = oSheet.getCellRangeByPosition ( _
		3, 1, 3, UBound (mData))
	oRange.merge (True)
	oRange = oSheet.getCellRangeByPosition ( _
		4, 1, 4, UBound (mData))
	oRange.merge (True)
	oRange = oSheet.getCellRangeByPosition ( _
		9, 1, 9, UBound (mData))
	oRange.setPropertyValue ("NumberFormat", 10)
	
	If aQuery.sPokemon = "Eevee" Then
		oRange = oSheet.getCellRangeByPosition ( _
			10, 0, 15, 0)
	Else
		If nEvolved = 0 Then
			oRange = oSheet.getCellRangeByPosition ( _
				10, 0, 10, 0)
		Else
			oRange = oSheet.getCellRangeByPosition ( _
				10, 0, 10 + nEvolved + 1, 0)
		End If
	End If
	oRange.setPropertyValue ("IsTextWrapped", True)
	
	oRows = oSheet.getRows
	oRows.getByIndex (0).setPropertyValue ("Height", 840)
	oColumns = oSheet.getColumns
	oColumns.getByIndex (0).setPropertyValue ("Width", 890)
	oColumns.getByIndex (1).setPropertyValue ("Width", 2310)
	oColumns.getByIndex (2).setPropertyValue ("Width", 890)
	oColumns.getByIndex (3).setPropertyValue ("Width", 890)
	oColumns.getByIndex (4).setPropertyValue ("Width", 1780)
	oColumns.getByIndex (5).setPropertyValue ("Width", 860)
	oColumns.getByIndex (6).setPropertyValue ("Width", 860)
	oColumns.getByIndex (7).setPropertyValue ("Width", 860)
	oColumns.getByIndex (8).setPropertyValue ("Width", 860)
	oColumns.getByIndex (9).setPropertyValue ("Width", 1030)
	If aQuery.sPokemon = "Eevee" Then
		If aQuery.nPlayerLevel <> 0 Then
			For nI = 0 To 5 Step 2
				oColumns.getByIndex (10 + nI).setPropertyValue ( _
					"Width", 2310)
				oColumns.getByIndex (10 + nI + 1).setPropertyValue ( _
					"Width", 2810)
			Next nI
		Else
			For nI = 0 To 2
				oColumns.getByIndex (10 + nI).setPropertyValue ( _
					"Width", 2310)
			Next nI
		End If
	Else
		If nEvolved = 0 Then
			If aQuery.nPlayerLevel <> 0 Then
				oColumns.getByIndex (10).setPropertyValue ( _
					"Width", 2200)
			End If
		Else
			For nI = 0 To nEvolved - 1
				oColumns.getByIndex (10 + nI).setPropertyValue ( _
					"Width", 2310)
			Next nI
			If aQuery.nPlayerLevel <> 0 Then
				oColumns.getByIndex ( _
					10 + nEvolved).setPropertyValue ( _
					"Width", 2810)
			End If
		End If
	End If
End Sub

' fnFilterAppraisals: Filters the IV by the appraisals.
Function fnFilterAppraisals (aQuery As aFindIVParam, _
		nAttack As Integer, nDefense As Integer, _
		nStamina As Integer) As Boolean
	Dim nTotal As Integer, nMax As Integer, sBest As String
	
	' The stats total.
	nTotal = nAttack + nDefense + nStamina
	If aQuery.nTotal = 1 And Not (nTotal >= 37) Then
		fnFilterAppraisals = True
		Exit Function
	End If
	If aQuery.nTotal = 2 And Not (nTotal >= 30 And nTotal <= 36) Then
		fnFilterAppraisals = True
		Exit Function
	End If
	If aQuery.nTotal = 3 And Not (nTotal >= 23 And nTotal <= 29) Then
		fnFilterAppraisals = True
		Exit Function
	End If
	If aQuery.nTotal = 4 And Not (nTotal <= 22) Then
		fnFilterAppraisals = True
		Exit Function
	End If
	' The best stats.
	nMax = nAttack
	If nDefense > nMax Then
		nMax = nDefense
	End If
	If nStamina > nMax Then
		nMax = nStamina
	End If
	If aQuery.sBest <> "" Then
		sBest = ""
		If nAttack = nMax Then
			sBest = sBest & "Atk "
		End If
		If nDefense = nMax Then
			sBest = sBest & "Def "
		End If
		If nStamina = nMax Then
			sBest = sBest & "Sta "
		End If
		If aQuery.sBest <> sBest Then
			fnFilterAppraisals = True
			Exit Function
		End If
	End If
	' The max stat value.
	If aQuery.nMax = 1 And Not (nMax = 15) Then
		fnFilterAppraisals = True
		Exit Function
	End If
	If aQuery.nMax = 2 And Not (nMax = 13 Or nMax = 14) Then
		fnFilterAppraisals = True
		Exit Function
	End If
	If aQuery.nMax = 3 And Not (nMax >= 8 And nMax <= 12) Then
		fnFilterAppraisals = True
		Exit Function
	End If
	If aQuery.nMax = 4 And Not (nMax <= 7) Then
		fnFilterAppraisals = True
		Exit Function
	End If
	fnFilterAppraisals = False
End Function

' fnCalcCP: Calculates the combat power of the Pokémon
Function fnCalcCP (aBaseStats As aStats, fLevel As Double, _
		nAttack As Integer, nDefense As Integer, _
		nStamina As Integer) As Integer
	Dim nCP As Integer
		
	nCP = fnFloor ((aBaseStats.nAttack + nAttack) _
		* ((aBaseStats.nDefense + nDefense) ^ 0.5) _
		* ((aBaseStats.nStamina + nStamina) ^ 0.5) _
		* (fnGetCPM (fLevel) ^ 2) / 10)
	If nCP < 10 Then
		nCP = 10
	End If
	fnCalcCP = nCP
End Function

' fnCalcHP: Calculates the hit points of the Pokémon
Function fnCalcHP (aBaseStats As aStats, _
		fLevel As Double, nStamina As Integer) As Integer
	Dim nHP As Integer
	
	nHP = fnFloor ((aBaseStats.nStamina + nStamina) _
		* fnGetCPM (fLevel))
	If nHP < 10 Then
		nHP = 10
	End If
	fnCalcHP = nHP
End Function

' fnGetBaseStats: Returns the base stats of the Pokémon.
Function fnGetBaseStats (sPokemon As String) As aStats
	Dim nI As Integer
	
	subReadBaseStats
	For nI = 0 To UBound (maBaseStats)
		If maBaseStats (nI).sPokemon = sPokemon Then
			fnGetBaseStats = maBaseStats (nI)
			Exit Function
		End If
	Next nI
End Function

' fnGetCPM: Returns the combat power multiplier.
Function fnGetCPM (fLevel As Double) As Double
	Dim nI As Integer
	
	subReadCPM
	If CInt (fLevel) = fLevel Then
		fnGetCPM = mCPM (fLevel)
	Else
		fnGetCPM = ((mCpm (fLevel - 0.5) ^ 2 _
			+ mCpm (fLevel + 0.5) ^ 2) / 2) ^ 0.5
	End If
End Function

' fnFloor: Returns the floor of the number
Function fnFloor (fNumber As Double) As Integer
	fnFloor = CInt (fNumber - 0.5)
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
	If sId = "" Then
		sId = sName
	End If
	fnMapPokemonNameToId = sId
End Function

' fnMapPokemonIdToName: Maps the Pokémon IDs to their localized names.
Function fnMapPokemonIdToName (sId As String) As String
	fnMapPokemonIdToName = fnGetResString ("Pokemon" & sId)
End Function

' subReadBaseStats: Reads the base stats table.
Sub subReadBaseStats
	Dim mData As Variant, nI As Integer, nJ As Integer
	DIm nEvolved As Integer
	
	If UBound (maBaseStats) = -1 Then
		mData = fnGetBaseStatsData
		ReDim Preserve maBaseStats (UBound (mData)) As New aStats
		For nI = 0 To UBound (mData)
			With maBaseStats (nI)
				.sNo = mData (nI) (1)
				.sPokemon = mData (nI) (0)
				.nStamina = mData (nI) (2)
				.nAttack = mData (nI) (3)
				.nDefense = mData (nI) (4)
			End With
			nEvolved = UBound (mData (nI) (5))
			maBaseStats (nI).mEvolved = fnGetStringArray (nEvolved)
			For nJ = 0 To nEvolved
				maBaseStats (nI).mEvolved (nJ) = mData (nI) (5) (nJ)
			Next nJ
		Next nI
	End If
End Sub

' fnGetStringArray: Obtains a blank string array
Function fnGetStringArray (nUBound As Integer) As Variant
	Dim mData () As String
	
	If nUBound >= 0 Then
		ReDim Preserve mData (nUBound) As String
	End If
	fnGetStringArray = mData
End Function

' fnGetEvolvedArray: Obtains a blank aEvolvedStats array
Function fnGetEvolvedArray (nUBound As Integer) As Variant
	Dim mData () As New aEvolvedStats
	
	If nUBound >= 0 Then
		ReDim Preserve mData (nUBound) As New aEvolvedStats
	End If
	fnGetEvolvedArray = mData
End Function

' fnReplace: Replaces all occurrances of a term to another.
Function fnReplace ( _
		sText As String, sFrom As String, sTo As String) As String
	Dim sResult As String, nPos As Integer
	
	sResult = sText
	nPos = InStr (sResult, sFrom)
	Do While nPos <> 0
		sResult = Left (sResult, nPos - 1) & sTo _
			& Right (sResult, Len (sResult) - nPos - Len (sFrom) + 1)
		nPos = InStr (nPos + Len (sTo), sResult, sFrom)
	Loop
	fnReplace = sResult
End Function

' subReadCPM: Reads the CPM table.
Sub subReadCPM
	If UBound (mCPM) = -1 Then
		mCPM = fnGetCPMData
	End If
End Sub

' subReadStardust: Reads the stardust table.
Sub subReadStardust
	If UBound (mStardust) = -1 Then
		mStardust = fnGetStardustData
	End If
End Sub
