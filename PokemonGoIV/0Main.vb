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

' The stats of a Pokémon.
Type aStats
	sNo As String
	sPokemon As String
	fLevel As Double
	nStamina As Integer
	nAttack As Integer
	nDefense As Integer
	nTotal As Integer
	nMaxCP As Integer
	maEvolvedForms () As aEvolveForm
End Type

' The amount of star dust to power-up.
Type aStarDust
	fLevel As Double
	nStarDust As Integer
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

Type aEvolveForm
	sPokemon As String
	nCP As Integer
	nMaxCP As Integer
End Type

Private maBaseStats () As New aStats
Private mCPM () As Double, mStarDust () As Integer

' subMain: The main program
Sub subMain
	Dim maIVs As Variant, nI As Integer
	Dim aQuery As New aFindIVParam, aBaseStats As New aStats
	
	aQuery = fnAskParam
	If aQuery.bIsCancelled Then
		Exit Sub
	End If
	maIVs = fnFindIV (aQuery)
	If UBound (maIVs) = -1 Then
		MsgBox fnGetResString ("ErrorNotFound")
	Else
		subSaveIV (aQuery, maIVs)
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
	subUpdateAppraisal1 (oEvent, True)
	' Checks if the required columns are filled.
	subBtnOKCheck (oEvent)
End Sub

' subRdoTeamRedItemChanged: When the red team is selected.
Sub subRdoTeamRedItemChanged (oEvent As object)
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
	subUpdateAppraisal1 (oEvent, False)
	
	oText = oDialog.getControl ("txtBestBefore")
	oText.setPosSize (-1, -1, 20, -1, _
		com.sun.star.awt.PosSize.WIDTH)
	oText.setVisible (True)
	oText.setText ("Its")
	
	mItems = Array ("Attack", "Defense", "HP")
	oList = oDialog.getControl ("lstBest")
	oList.removeItems (0, oList.getItemCount())
	oList.addItems (mItems, 0)
	oList.setPosSize (140, -1, -1, -1, _
		com.sun.star.awt.PosSize.X)
	oList.setVisible (True)
	
	oText = oDialog.getControl ("txtBestAfter")
	oText.setPosSize (240, -1, 160, -1, _
		com.sun.star.awt.PosSize.X + com.sun.star.awt.PosSize.WIDTH)
	oText.setVisible (True)
	oText.setText ("is its strongest feature.")
	
	oList = oDialog.getControl ("cbxBest2")
	oList.setVisible (False)
	
	oList = oDialog.getControl ("cbxBest3")
	oList.setVisible (False)
	
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

' subRdoTeamBlueItemChanged: When the blue team is selected.
Sub subRdoTeamBlueItemChanged (oEvent As object)
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
	subUpdateAppraisal1 (oEvent, False)
	
	oText = oDialog.getControl ("txtBestBefore")
	oText.setPosSize (-1, -1, 200, -1, _
		com.sun.star.awt.PosSize.WIDTH)
	oText.setVisible (True)
	oText.setText ("I see that its best attribute is its")
	
	mItems = Array ("Attack", "Defense", "HP")
	oList = oDialog.getControl ("lstBest")
	oList.removeItems (0, oList.getItemCount())
	oList.addItems (mItems, 0)
	oList.setPosSize (320, -1, -1, -1, _
		com.sun.star.awt.PosSize.X)
	oList.setVisible (True)
	
	oText = oDialog.getControl ("txtBestAfter")
	oText.setPosSize (415, -1, 5, -1, _
		com.sun.star.awt.PosSize.X + com.sun.star.awt.PosSize.WIDTH)
	oText.setVisible (True)
	oText.setText (".")
	
	oList = oDialog.getControl ("cbxBest2")
	oList.setVisible (False)
	
	oList = oDialog.getControl ("cbxBest3")
	oList.setVisible (False)
	
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

' subRdoTeamYellowItemChanged: When the yellow team is selected.
Sub subRdoTeamYellowItemChanged (oEvent As object)
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
	subUpdateAppraisal1 (oEvent, False)
	
	oText = oDialog.getControl ("txtBestBefore")
	oText.setPosSize (-1, -1, 115, -1, _
		com.sun.star.awt.PosSize.WIDTH)
	oText.setVisible (True)
	oText.setText ("Its best quality is")
	
	mItems = Array ("Attack", "Defense", "HP")
	oList = oDialog.getControl ("lstBest")
	oList.removeItems (0, oList.getItemCount())
	oList.addItems (mItems, 0)
	oList.setPosSize (240, -1, -1, -1, _
		com.sun.star.awt.PosSize.X)
	oList.setVisible (True)
	
	oText = oDialog.getControl ("txtBestAfter")
	oText.setPosSize (335, -1, 5, -1, _
		com.sun.star.awt.PosSize.X + com.sun.star.awt.PosSize.WIDTH)
	oText.setVisible (True)
	oText.setText (".")
	
	oList = oDialog.getControl ("cbxBest2")
	oList.setVisible (False)
	
	oList = oDialog.getControl ("cbxBest3")
	oList.setVisible (False)
	
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

' subLstBestItemChanged: When the best stat is selected.
Sub subLstBestItemChanged (oEvent As object)
	Dim oDialog As Object, oCheckBox As Object
	
	oDialog = oEvent.Source.getContext
	If oDialog.getControl ("rdoTeamRed").getState Then
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
	If oDialog.getControl ("rdoTeamBlue").getState Then
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
	If oDialog.getControl ("rdoTeamYellow").getState Then
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
Sub subUpdateAppraisal1 (oEvent As Object, bIsKeepSelected As Boolean)
	Dim oDialog As Object, sPokemon As String
	Dim oList As Object, nSelected As Integer
	Dim mItems () As String, nI As Integer
	
	oDialog = oEvent.Source.getContext
	
	If oDialog.getControl ("rdoTeamRed").getState Then
	    mItems = Array ( _
		    "Overall, your [Pokémon] simply amazes me. It can accomplish anything!", _
		    "Overall, your [Pokémon] is a strong Pokémon. You should be proud!", _
		    "Overall, your [Pokémon] is a decent Pokémon.", _
		    "Overall, your [Pokémon] may not be great in battle, but I still like it!")
	End If
	If oDialog.getControl ("rdoTeamBlue").getState Then
	    mItems = Array ( _
		    "Overall, your [Pokémon] is a wonder! What a breathtaking Pokémon!", _
		    "Overall, your [Pokémon] has certainly caught my attention.", _
		    "Overall, your [Pokémon] is above average.", _
		    "Overall, your [Pokémon] is not likely to make much headway in battle.")
	End If
	If oDialog.getControl ("rdoTeamYellow").getState Then
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
Function fnFindIV (aQuery As aFindIVParam) As Variant
	Dim aBaseStats As New aStats, maIV () As New aStats
	Dim fLevel As Double, nStamina As Integer
	Dim nAttack As Integer, nDefense As integer
	Dim nI As Integer, nJ As Integer
	Dim fStep As Double, nCount As Integer
	Dim aEvBaseStats As new aStats, aTempIV As New aStats
	Dim maEvolvedForms () As New aEvolveForm
	
	If aQuery.sPokemon = "" Then
		fnFindIV = maIV
		Exit Function
	End If
	If aQuery.bIsNew Then
		fStep = 1
	Else
		fStep = 0.5
	End If
	aBaseStats = fnGetBaseStats (aQuery.sPokemon)
	subReadStarDust
	nCount = -1
	For fLevel = 1 To UBound (mStarDust) Step fStep
		If mStarDust (CInt (fLevel - 0.5)) = aQuery.nStarDust Then
			For nStamina = 0 To 15
				If fnCalcHP (aBaseStats, fLevel, nStamina) = aQuery.nHP Then
					For nAttack = 0 To 15
						For nDefense = 0 To 15
							If fnCalcCP (aBaseStats, fLevel, nAttack, nDefense, nStamina) = aQuery.nCP _
									And Not (fnFilterAppraisals (aQuery, nAttack, nDefense, nStamina)) Then
								nCount = nCount + 1
								ReDim Preserve maIV (nCount) As New aStats
								With maIV (nCount)
									.sNo = aBaseStats.sNo
									.sPokemon = aQuery.sPokemon
									.fLevel = fLevel
									.nAttack = nAttack
									.nDefense = nDefense
									.nStamina = nStamina
									.nTotal = nAttack + nDefense + nStamina
								End With
								If aQuery.nPlayerLevel <> 0 Then
									maIV (nCount).nMaxCP = fnCalcCP (aBaseStats, aQuery.nPlayerLevel + 1.5, nAttack, nDefense, nStamina)
								Else
									maIV (nCount).nMaxCP = -1
								End If
								maIV (nCount).maEvolvedForms = fnGetEvolvedFormArray (UBound (aBaseStats.maEvolvedForms))
								For nI = 0 To UBound (aBaseStats.maEvolvedForms)
									maIV (nCount).maEvolvedForms (nI).sPokemon = aBaseStats.maEvolvedForms (nI).sPokemon
									aEvBaseStats = fnGetBaseStats (aBaseStats.maEvolvedForms (nI).sPokemon)
									maIV (nCount).maEvolvedForms (nI).nCP = fnCalcCP (aEvBaseStats, fLevel, nAttack, nDefense, nStamina)
									If aQuery.nPlayerLevel <> 0 Then
										maIV (nCount).maEvolvedForms (nI).nMaxCP = fnCalcCP (aEvBaseStats, aQuery.nPlayerLevel + 1.5, nAttack, nDefense, nStamina)
									Else
										maIV (nCount).maEvolvedForms (nI).nMaxCP = -1
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
				subCopyIV (maIV (nI), aTempIV)
				subCopyIV (maIV (nJ), maIV (nI))
				subCopyIV (aTempIV, maIV (nJ))
			End If
		Next nJ
	Next nI
	fnFindIV = maIV
End Function

' fnCompareIV: Compare two IVs for sorting
Function fnCompareIV (aIVa As aStats, aIVb As aStats) As Double
	Dim nCPa As Integer, nCPb As Integer, nI As Integer
	
	nCPa = aIVa.nMaxCP
	For nI = 0 To UBound (aIVa.maEvolvedForms)
		If nCPa < aIVa.maEvolvedForms (nI).nMaxCP Then
			nCPa = aIVa.maEvolvedForms (nI).nMaxCP
		End If
	Next nI
	nCPb = aIVb.nMaxCP
	For nI = 0 To UBound (aIVb.maEvolvedForms)
		If nCPb < aIVb.maEvolvedForms (nI).nMaxCP Then
			nCPb = aIVb.maEvolvedForms (nI).nMaxCP
		End If
	Next nI
	fnCompareIV = nCPb - nCPa
	If fnCompareIV <> 0 Then
		Exit Function
	End If
	
	nCPa = 0
	For nI = 0 To UBound (aIVa.maEvolvedForms)
		If nCPa < aIVa.maEvolvedForms (nI).nCP Then
			nCPa = aIVa.maEvolvedForms (nI).nCP
		End If
	Next nI
	nCPb = 0
	For nI = 0 To UBound (aIVb.maEvolvedForms)
		If nCPb < aIVb.maEvolvedForms (nI).nCP Then
			nCPb = aIVb.maEvolvedForms (nI).nCP
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
Function subCopyIV (aFrom As aStats, aTo As aStats) As Double
	Dim nI As Integer, maEvolvedForms () As New aEvolveForm
	
	With aTo
		.sNo = aFrom.sNo
		.sPokemon = aFrom.sPokemon
		.fLevel = aFrom.fLevel
		.nAttack = aFrom.nAttack
		.nDefense = aFrom.nDefense
		.nStamina = aFrom.nStamina
		.nTotal = aFrom.nTotal
		.nMaxCP = aFrom.nMaxCP
	End With
	aTo.maEvolvedForms = fnGetEvolvedFormArray (UBound (aFrom.maEvolvedForms))
	For nI = 0 To UBound (aFrom.maEvolvedForms)
		With aTo.maEvolvedForms (nI)
			.sPokemon = aFrom.maEvolvedForms (nI).sPokemon
			.nCP = aFrom.maEvolvedForms (nI).nCP
			.nMaxCP = aFrom.maEvolvedForms (nI).nMaxCP
		End With
	Next nI
End Function

' subSaveIV: Saves the found IV
Sub subSaveIV (aQuery As aFindIVParam, maIVs () As aStats)
	Dim oDoc As Object, oSheet As Object
	Dim oRange As Object, oColumns As Object, oRows As Object
	Dim nI As Integer, nJ As Integer, nFrontCols As Integer
	Dim mData (Ubound (maIVs) + 1) As Variant, mRow () As Variant
	Dim mProps () As New com.sun.star.beans.PropertyValue
	
	oDoc = StarDesktop.loadComponentFromURL ( _
		"private:factory/scalc", "_default", 0, mProps)
	oSheet = oDoc.getSheets.getByIndex (0)
	
	mRow = Array ( _
		"No", "Pokemon", "CP", "HP", "Star dust", _
		"Lv", "Atk", "Def", "Sta", "IV")
	nFrontCols = UBound (mRow)
	If aQuery.sPokemon = "Eevee" Then
		If aQuery.nPlayerLevel <> 0 Then
			ReDim Preserve mRow (nFrontCols + 6) As Variant
			mRow (nFrontCols + 1) = "CP as " & maIVs (0).maEvolvedForms (0).sPokemon
			mRow (nFrontCols + 2) = "Powered-up as " & maIVs (0).maEvolvedForms (0).sPokemon
			mRow (nFrontCols + 3) = "CP as " & maIVs (0).maEvolvedForms (1).sPokemon
			mRow (nFrontCols + 4) = "Powered-up as " & maIVs (0).maEvolvedForms (1).sPokemon
			mRow (nFrontCols + 5) = "CP as " & maIVs (0).maEvolvedForms (2).sPokemon
			mRow (nFrontCols + 6) = "Powered-up as " & maIVs (0).maEvolvedForms (2).sPokemon
		Else
			ReDim Preserve mRow (nFrontCols + 3) As Variant
			mRow (nFrontCols + 1) = "CP as " & maIVs (0).maEvolvedForms (0).sPokemon
			mRow (nFrontCols + 2) = "CP as " & maIVs (0).maEvolvedForms (1).sPokemon
			mRow (nFrontCols + 3) = "CP as " & maIVs (0).maEvolvedForms (2).sPokemon
		End If
	Else
		If UBound (maIVs (0).maEvolvedForms) = -1 Then
			If aQuery.nPlayerLevel <> 0 Then
				ReDim Preserve mRow (nFrontCols + 1) As Variant
				mRow (nFrontCols + 1) = "Powered-up"
			End If
		Else
			If aQuery.nPlayerLevel <> 0 Then
				ReDim Preserve mRow (nFrontCols + UBound (maIVs (0).maEvolvedForms) + 2) As Variant
				For nJ = 0 To UBound (maIVs (0).maEvolvedForms)
					mRow (nFrontCols + nJ + 1) = "CP as " & maIVs (0).maEvolvedForms (nJ).sPokemon
				Next nJ
				mRow (UBound (mRow)) = "Powered-up as " & maIVs (0).maEvolvedForms (UBound (maIVs (0).maEvolvedForms)).sPokemon
			Else
				ReDim Preserve mRow (nFrontCols + UBound (maIVs (0).maEvolvedForms) + 1) As Variant
				For nJ = 0 To UBound (maIVs (0).maEvolvedForms)
					mRow (nFrontCols + nJ + 1) = "CP as " & maIVs (0).maEvolvedForms (nJ).sPokemon
				Next nJ
			End If
		End If
	End If
	mData (0) = mRow
	
	For nI = 0 To UBound (maIVs)
		mRow = Array ( _
			"", "", "", "", "", _
			maIVs (nI).fLevel, maIVs (nI).nAttack, maIVs (nI).nDefense, _
			maIVs (nI).nStamina, maIVs (nI).nTotal / 45)
		If aQuery.sPokemon = "Eevee" Then
			If aQuery.nPlayerLevel <> 0 Then
				ReDim Preserve mRow (nFrontCols + 6) As Variant
				mRow (nFrontCols + 1) = maIVs (nI).maEvolvedForms (0).nCP
				mRow (nFrontCols + 2) = maIVs (nI).maEvolvedForms (0).nMaxCP
				mRow (nFrontCols + 3) = maIVs (nI).maEvolvedForms (1).nCP
				mRow (nFrontCols + 4) = maIVs (nI).maEvolvedForms (1).nMaxCP
				mRow (nFrontCols + 5) = maIVs (nI).maEvolvedForms (2).nCP
				mRow (nFrontCols + 6) = maIVs (nI).maEvolvedForms (2).nMaxCP
			Else
				ReDim Preserve mRow (nFrontCols + 3) As Variant
				mRow (nFrontCols + 1) = maIVs (nI).maEvolvedForms (0).nCP
				mRow (nFrontCols + 2) = maIVs (nI).maEvolvedForms (1).nCP
				mRow (nFrontCols + 3) = maIVs (nI).maEvolvedForms (2).nCP
			End If
		Else
			If UBound (maIVs (nI).maEvolvedForms) = -1 Then
				If aQuery.nPlayerLevel <> 0 Then
					ReDim Preserve mRow (nFrontCols + 1) As Variant
					mRow (nFrontCols + 1) = maIVs (nI).nMaxCP
				End If
			Else
				If aQuery.nPlayerLevel <> 0 Then
					ReDim Preserve mRow (nFrontCols + UBound (maIVs (nI).maEvolvedForms) + 2) As Variant
					For nJ = 0 To UBound (maIVs (nI).maEvolvedForms)
						mRow (nFrontCols + nJ + 1) = maIVs (nI).maEvolvedForms (nJ).nCP
					Next nJ
					mRow (UBound (mRow)) = maIVs (nI).maEvolvedForms (UBound (maIVs (nI).maEvolvedForms)).nMaxCP
				Else
					ReDim Preserve mRow (nFrontCols + UBound (maIVs (nI).maEvolvedForms) + 1) As Variant
					For nJ = 0 To UBound (maIVs (nI).maEvolvedForms)
						mRow (nFrontCols + nJ + 1) = maIVs (nI).maEvolvedForms (nJ).nCP
					Next nJ
				End If
			End If
		End If
		mData (nI + 1) = mRow
	Next nI
	
	' Fills the query information at the first row
	mData (1) (0) = maIVs (0).sNo
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
		If UBound (maIVs (0).maEvolvedForms) = -1 Then
			oRange = oSheet.getCellRangeByPosition ( _
				10, 0, 10, 0)
		Else
			oRange = oSheet.getCellRangeByPosition ( _
				10, 0, 10 + UBound (maIVs (0).maEvolvedForms) + 2, 0)
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
				oColumns.getByIndex (10 + nI).setPropertyValue ("Width", 2310)
				oColumns.getByIndex (10 + nI + 1).setPropertyValue ("Width", 2810)
			Next nI
		Else
			For nI = 0 To 2
				oColumns.getByIndex (10 + nI).setPropertyValue ("Width", 2310)
			Next nI
		End If
	Else
		If UBound (maIVs (0).maEvolvedForms) = -1 Then
			If aQuery.nPlayerLevel <> 0 Then
				oColumns.getByIndex (10).setPropertyValue ("Width", 2200)
			End If
		Else
			For nI = 0 To UBound (maIVs (0).maEvolvedForms)
				oColumns.getByIndex (10 + nI).setPropertyValue ("Width", 2310)
			Next nI
			If aQuery.nPlayerLevel <> 0 Then
				oColumns.getByIndex (10 + UBound (maIVs (0).maEvolvedForms) + 1).setPropertyValue ("Width", 2810)
			End If
		End If
	End If
End Sub

' fnFilterAppraisals: Filters the IV by the appraisals.
Function fnFilterAppraisals (aQuery As aFindIVParam, nAttack As Integer, nDefense As Integer, nStamina As Integer) As Boolean
	Dim nTotal As Integer, nMax As Integer, sBest As String
	
	' The first appraisal.
	nTotal = nAttack + nDefense + nStamina
	If aQuery.nAppraisal1 = 1 And Not (nTotal >= 37) Then
		fnFilterAppraisals = True
		Exit Function
	End If
	If aQuery.nAppraisal1 = 2 And Not (nTotal >= 30 And nTotal <= 36) Then
		fnFilterAppraisals = True
		Exit Function
	End If
	If aQuery.nAppraisal1 = 3 And Not (nTotal >= 23 And nTotal <= 29) Then
		fnFilterAppraisals = True
		Exit Function
	End If
	If aQuery.nAppraisal1 = 4 And Not (nTotal <= 22) Then
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
Function fnCalcCP (aBaseStats As aStats, fLevel As Double, nAttack As Integer, nDefense As Integer, nStamina As Integer) As Integer
	fnCalcCP = fnFloor ((aBaseStats.nAttack + nAttack) _
		* ((aBaseStats.nDefense + nDefense) ^ 0.5) _
		* ((aBaseStats.nStamina + nStamina) ^ 0.5) _
		* (fnGetCPM (fLevel) ^ 2) / 10)
End Function

' fnCalcHP: Calculates the hit points of the Pokémon
Function fnCalcHP (aBaseStats As aStats, fLevel As Double, nStamina As Integer) As Integer
	fnCalcHP = fnFloor ((aBaseStats.nStamina + nStamina) _
		* fnGetCPM (fLevel))
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
			maBaseStats (nI).maEvolvedForms = fnGetEvolvedFormArray (UBound (mData (nI) (5)))
			For nJ = 0 To UBound (maBaseStats (nI).maEvolvedForms)
				With maBaseStats (nI).maEvolvedForms (nJ)
					.sPokemon = mData (nI) (5) (nJ)
					.nCP = -1
					.nMaxCP = -1
				End With
			Next nJ
		Next nI
	End If
End Sub

' fnGetEvolvedFormArray: Obtains a blank aEvolveForm array
Function fnGetEvolvedFormArray (nUBound As Integer) As Variant
	If nUBound = -1 Then
		fnGetEvolvedFormArray = fnGetEmptyEvolvedFormArray
	Else
		fnGetEvolvedFormArray = fnGetNumberedEvolvedFormArray (nUBound)
	End If
End Function

' fnGetNumberedEvolvedFormArray: Obtains a numbered aEvolveForm array
Function fnGetNumberedEvolvedFormArray (nUBound As Integer) As Variant
	Dim mData (nUBound) As New aEvolveForm
	
	fnGetNumberedEvolvedFormArray = mData
End Function

' fnGetEmptyEvolvedFormArray: Obtains an empty aEvolveForm array
Function fnGetEmptyEvolvedFormArray () As Variant
	Dim mData () As New aEvolveForm
	
	fnGetEmptyEvolvedFormArray = mData
End Function

' fnReplace: Replaces all occurrances of a term to another.
Function fnReplace (sText As String, sFrom As String, sTo As String) As String
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
