' Copyright (c) 2016 imacat.
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
	nStarDust As Integer
	nPlayerLevel As Integer
	bIsNew As Boolean
	nAppraisal1 As Integer
	sBest As String
	nAppraisal2 As Integer
	bIsCancelled As Boolean
End Type

Private maBaseStats () As New aStats
Private mCPM () As Double, mStarDust () As Integer

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
	Dim oDialog As Object, oDialogModel As Object
	Dim bIsBestAttack As Boolean, bIsBestDefense As Boolean
	Dim bIsBestHP As Boolean
	Dim aQuery As New aFindIVParam
	
	DialogLibraries.loadLibrary "PokemonGoIV"
	oDialog = CreateUnoDialog (DialogLibraries.PokemonGoIV.DlgMain)
	oDialog.getControl ("lstAppraisal1").setVisible (False)
	oDialog.getControl ("txtBestBefore").setVisible (False)
	oDialog.getControl ("lstBest").setVisible (False)
	oDialog.getControl ("txtBestAfter").setVisible (False)
	oDialog.getControl ("cbxBest2").setVisible (False)
	oDialog.getControl ("cbxBest3").setVisible (False)
	oDialog.getControl ("lstAppraisal2").setVisible (False)
	
	oDialog.getControl ("imgPokemon").getModel.setPropertyValue ( _
		"ImageURL", fnGetImageUrl ("Unknown"))
	oDialog.getControl ("imgTeamLogo").getModel.setPropertyValue ( _
		"ImageURL", fnGetImageUrl ("Unknown"))
	
	If oDialog.execute = 0 Then
		aQuery.bIsCancelled = True
		fnAskParam = aQuery
		Exit Function
	End If
	
	With aQuery
		.sPokemon = oDialog.getControl ("lstPokemon").getSelectedItem
		.nCP = oDialog.getControl ("numCP").getValue
		.nHP = oDialog.getControl ("numHP").getValue
		.nStarDust = CInt (oDialog.getControl ("lstStarDust").getSelectedItem)
		.nPlayerLevel = CInt (oDialog.getControl ("lstPlayerLevel").getSelectedItem)
		.nAppraisal1 = oDialog.getControl ("lstAppraisal1").getSelectedItemPos + 1
		.nAppraisal2 = oDialog.getControl ("lstAppraisal2").getSelectedItemPos + 1
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
	If oDialog.getControl ("lstBest").getSelectedItem = "Attack" Then
		bIsBestAttack = True
		If oDialog.getControl ("cbxBest2").getState = 1 Then
			bIsBestDefense = True
		End If
		If oDialog.getControl ("cbxBest3").getState = 1 Then
			bIsBestHP = True
		End If
	End If
	If oDialog.getControl ("lstBest").getSelectedItem = "Defense" Then
		bIsBestDefense = True
		If oDialog.getControl ("cbxBest2").getState = 1 Then
			bIsBestAttack = True
		End If
		If oDialog.getControl ("cbxBest3").getState = 1 Then
			bIsBestHP = True
		End If
	End If
	If oDialog.getControl ("lstBest").getSelectedItem = "HP" Then
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
	Dim oHP As Object, oStarDust As Object, oOK As Object
	
	oDialog = oEvent.Source.getContext
	oPokemon = oDialog.getControl ("lstPokemon")
	oCP = oDialog.getControl ("numCP")
	oHP = oDialog.getControl ("numHP")
	oStarDust = oDialog.getControl ("lstStarDust")
	oOK = oDialog.getControl ("btnOK")
	
	If oPokemon.getSelectedItemPos <> -1 _
			And oCP.getText <> "" _
			And oHP.getText <> "" _
			And oStarDust.getSelectedItemPos <> -1 Then
		oOK.setEnable (True)
	Else
		oOK.setEnable (False)
	End If
End Sub

' subLstPokemonSelected: When the Pokémon is selected.
Sub subLstPokemonSelected (oEvent As object)
	Dim oDialog As Object, sPokemon As String
	Dim oImageModel As Object, sImageId As String
	
	oDialog = oEvent.Source.getContext
	
	' Updates the Pokémon image.
	sPokemon = oDialog.getControl ("lstPokemon").getSelectedItem
	sImageId = ""
	If sPokemon = "Farfetch'd" Then
		sImageId = "PokemonFarfetchd"
	End If
	If sPokemon = "Nidoran♀" Then
		sImageId = "PokemonNidoranFemale"
	End If
	If sPokemon = "Nidoran♂" Then
		sImageId = "PokemonNidoranMale"
	End If
	If sPokemon = "Mr. Mime" Then
		sImageId = "PokemonMrMime"
	End If
	If sImageId = "" Then
		sImageId = "Pokemon" & sPokemon
	End If
	oImageModel = oDialog.getControl ("imgPokemon").getModel
	oImageModel.setPropertyValue ("ImageURL", _
		fnGetImageUrl (sImageId))
	
	' Updates the text of the first appraisal.
	subUpdateAppraisal1 (oDialog, True)
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
	
	' Updates the text of the first appraisal.
	subUpdateAppraisal1 (oDialog, False)
	
	' Updates the text of the best stat appraisal.
	subUpdateBestStatAppraisal (oDialog, _
		"Its", 8, "is its strongest feature.", 65)
	
	mItems = Array ( _
		"I'm blown away by its stats. WOW!", _
		"It's got excellent stats! How exciting!", _
		"Its stats indicate that in battle, it'll get the job done.", _
		"Its stats don't point to greatness in battle.")
	oList = oDialog.getControl ("lstAppraisal2")
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
	
	' Updates the text of the first appraisal.
	subUpdateAppraisal1 (oDialog, False)
	
	' Updates the text of the best stat appraisal.
	subUpdateBestStatAppraisal (oDialog, _
		"I see that its best attribute is its", 85, ".", 5)
	
	mItems = Array ( _
		"Its stats exceed my calculations. It's incredible!", _
		"I am certainly impressed by its stats, I must say.", _
		"Its stats are noticeably trending to the positive.", _
		"Its stats are not out of the norm, in my opinion.")
	oList = oDialog.getControl ("lstAppraisal2")
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
	
	' Updates the text of the first appraisal.
	subUpdateAppraisal1 (oDialog, False)
	
	' Updates the text of the best stat appraisal.
	subUpdateBestStatAppraisal (oDialog, _
		"Its best quality is", 45, ".", 5)
	
	mItems = Array ( _
		"Its stats are the best I've ever seen! No doubt about it!", _
		"Its stats are really strong! Impressive.", _
		"It's definitely got some good stats. Definitely!", _
		"Its stats are all right, but kinda basic, as far as I can see.")
	oList = oDialog.getControl ("lstAppraisal2")
	oList.removeItems (0, oList.getItemCount())
	oList.addItems (mItems, 0)
	oList.setVisible (True)
End Sub

' subUpdateBestStatAppraisal: Updates the text of the best stat appraisal.
Sub subUpdateBestStatAppraisal (oDialog As Object, _
		sBefore As String, nBeforeWidth As Integer, _
		sAfter As String, nAfterWidth As Integer)
	Dim oText As Object, oList As Object, nX As Integer
	Dim mItems () As String
	
	oText = oDialog.getControl ("txtBestBefore")
	oText.getModel.setPropertyValue ("Width", nBeforeWidth)
	oText.setVisible (True)
	oText.setText (sBefore)
	nX = oText.getModel.getPropertyValue ("PositionX") + nBeforeWidth
	
	mItems = Array ("Attack", "Defense", "HP")
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
	Dim oDialog As Object, oCheckBox As Object
	
	oDialog = oEvent.Source.getContext
	If oDialog.getControl ("rdoTeamValor").getState Then
		If oDialog.getControl ("lstBest").getSelectedItem = "Attack" Then
			oCheckBox = oDialog.getControl ("cbxBest2")
			oCheckBox.setLabel ("I'm just as impressed with its Defense.")
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
			oCheckBox = oDialog.getControl ("cbxBest3")
			oCheckBox.setLabel ("I'm just as impressed with its HP.")
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
		End If
		If oDialog.getControl ("lstBest").getSelectedItem = "Defense" Then
			oCheckBox = oDialog.getControl ("cbxBest2")
			oCheckBox.setLabel ("I'm just as impressed with its Attack.")
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
			oCheckBox = oDialog.getControl ("cbxBest3")
			oCheckBox.setLabel ("I'm just as impressed with its HP.")
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
		End If
		If oDialog.getControl ("lstBest").getSelectedItem = "HP" Then
			oCheckBox = oDialog.getControl ("cbxBest2")
			oCheckBox.setLabel ("I'm just as impressed with its Attack.")
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
			oCheckBox = oDialog.getControl ("cbxBest3")
			oCheckBox.setLabel ("I'm just as impressed with its Defense.")
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
		End If
	End If
	If oDialog.getControl ("rdoTeamMystic").getState Then
		If oDialog.getControl ("lstBest").getSelectedItem = "Attack" Then
			oCheckBox = oDialog.getControl ("cbxBest2")
			oCheckBox.setLabel ("It is matched equally by its Defense.")
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
			oCheckBox = oDialog.getControl ("cbxBest3")
			oCheckBox.setLabel ("It is matched equally by its HP.")
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
		End If
		If oDialog.getControl ("lstBest").getSelectedItem = "Defense" Then
			oCheckBox = oDialog.getControl ("cbxBest2")
			oCheckBox.setLabel ("It is matched equally by its Attack.")
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
			oCheckBox = oDialog.getControl ("cbxBest3")
			oCheckBox.setLabel ("It is matched equally by its HP.")
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
		End If
		If oDialog.getControl ("lstBest").getSelectedItem = "HP" Then
			oCheckBox = oDialog.getControl ("cbxBest2")
			oCheckBox.setLabel ("It is matched equally by its Attack.")
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
			oCheckBox = oDialog.getControl ("cbxBest3")
			oCheckBox.setLabel ("It is matched equally by its Defense.")
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
		End If
	End If
	If oDialog.getControl ("rdoTeamInstinct").getState Then
		If oDialog.getControl ("lstBest").getSelectedItem = "Attack" Then
			oCheckBox = oDialog.getControl ("cbxBest2")
			oCheckBox.setLabel ("Its Defense is great, too!")
			oCheckBox.setVisible (True)
			oCheckBox = oDialog.getControl ("cbxBest3")
			oCheckBox.setLabel ("Its HP is great, too!")
			oCheckBox.setVisible (True)
		End If
		If oDialog.getControl ("lstBest").getSelectedItem = "Defense" Then
			oCheckBox = oDialog.getControl ("cbxBest2")
			oCheckBox.setLabel ("Its Attack is great, too!")
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
			oCheckBox = oDialog.getControl ("cbxBest3")
			oCheckBox.setLabel ("Its HP is great, too!")
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
		End If
		If oDialog.getControl ("lstBest").getSelectedItem = "HP" Then
			oCheckBox = oDialog.getControl ("cbxBest2")
			oCheckBox.setLabel ("Its Attack is great, too!")
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
			oCheckBox = oDialog.getControl ("cbxBest3")
			oCheckBox.setLabel ("Its Defense is great, too!")
			oCheckBox.setVisible (True)
			oCheckBox.setState (0)
		End If
	End If
End Sub

' subUpdateAppraisal1: Updates the text of the first appraisal.
Sub subUpdateAppraisal1 (oDialog As Object, bIsKeepSelected As Boolean)
	Dim sPokemon As String, oList As Object, nSelected As Integer
	Dim mItems () As String, nI As Integer
	
	If oDialog.getControl ("rdoTeamValor").getState Then
		mItems = Array ( _
			"Overall, your [Pokémon] simply amazes me. It can accomplish anything!", _
			"Overall, your [Pokémon] is a strong Pokémon. You should be proud!", _
			"Overall, your [Pokémon] is a decent Pokémon.", _
			"Overall, your [Pokémon] may not be great in battle, but I still like it!")
	End If
	If oDialog.getControl ("rdoTeamMystic").getState Then
		mItems = Array ( _
			"Overall, your [Pokémon] is a wonder! What a breathtaking Pokémon!", _
			"Overall, your [Pokémon] has certainly caught my attention.", _
			"Overall, your [Pokémon] is above average.", _
			"Overall, your [Pokémon] is not likely to make much headway in battle.")
	End If
	If oDialog.getControl ("rdoTeamInstinct").getState Then
		mItems = Array ( _
			"Overall, your [Pokémon] looks like it can really battle with the best of them!", _
			"Overall, your [Pokémon] is really strong!", _
			"Overall, your [Pokémon] is pretty decent!", _
			"Overall, your [Pokémon] has room for improvement as far as battling goes.")
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
	
	oList = oDialog.getControl ("lstAppraisal1")
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
	subReadStarDust
	nEvolved = UBound (aBaseStats.mEvolved)
	ReDim maEvBaseStats (nEvolved) As New aStats
	For nI = 0 To nEvolved
		aTempStats = fnGetBaseStats (aBaseStats.mEvolved (nI))
		With maEvBaseStats (nI)
			.nAttack = aTempStats.nAttack
			.nDefense = aTempStats.nDefense
			.nStamina = aTempStats.nStamina
		End With
	Next nI
	nN = -1
	For fLevel = 1 To UBound (mStarDust) Step fStep
		If mStarDust (CInt (fLevel - 0.5)) = aQuery.nStarDust Then
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
		"No", "Pokemon", "CP", "HP", "Star dust", _
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
	mData (1) (4) = aQuery.nStarDust
	
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
	
	' The first appraisal.
	nTotal = nAttack + nDefense + nStamina
	If aQuery.nAppraisal1 = 1 _
			And Not (nTotal >= 37) Then
		fnFilterAppraisals = True
		Exit Function
	End If
	If aQuery.nAppraisal1 = 2 _
			And Not (nTotal >= 30 And nTotal <= 36) Then
		fnFilterAppraisals = True
		Exit Function
	End If
	If aQuery.nAppraisal1 = 3 _
			And Not (nTotal >= 23 And nTotal <= 29) Then
		fnFilterAppraisals = True
		Exit Function
	End If
	If aQuery.nAppraisal1 = 4 _
			And Not (nTotal <= 22) Then
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
	' The second appraisal.
	If aQuery.nAppraisal2 = 1 And Not (nMax = 15) Then
		fnFilterAppraisals = True
		Exit Function
	End If
	If aQuery.nAppraisal2 = 2 And Not (nMax = 13 Or nMax = 14) Then
		fnFilterAppraisals = True
		Exit Function
	End If
	If aQuery.nAppraisal2 = 3 And Not (nMax >= 8 And nMax <= 12) Then
		fnFilterAppraisals = True
		Exit Function
	End If
	If aQuery.nAppraisal2 = 4 And Not (nMax <= 7) Then
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
	Dim nPos As Integer
	
	nPos = InStr (sText, sFrom)
	Do While nPos <> 0
		sText = Left (sText, nPos - 1) & sTo _
			& Right (sText, Len (sText) - nPos - Len (sFrom) + 1)
		nPos = InStr (nPos + Len (sTo), sText, sFrom)
	Loop
	fnReplace = sText
End Function

' subReadCPM: Reads the CPM table.
Sub subReadCPM
	If UBound (mCPM) = -1 Then
		mCPM = fnGetCPMData
	End If
End Sub

' subReadStarDust: Reads the star dust table.
Sub subReadStarDust
	If UBound (mStarDust) = -1 Then
		mStarDust = fnGetStarDustData
	End If
End Sub
