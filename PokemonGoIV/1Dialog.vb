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

' 1Dialog: The Dialog UI processor
'   by imacat <imacat@mail.imacat.idv.tw>, 2017-02-24

Option Explicit

' fnLoadParamDialog: Loads the parameter dialog.
Function fnLoadParamDialog As Object
	Dim oDialog As Object
	
	DialogLibraries.loadLibrary "PokemonGoIV"
	oDialog = CreateUnoDialog (DialogLibraries.PokemonGoIV.DlgMain)
	oDialog.getControl ("lstTotal").setVisible (False)
	oDialog.getControl ("txtBestHead").setVisible (False)
	oDialog.getControl ("lstBest").setVisible (False)
	oDialog.getControl ("txtBestTail").setVisible (False)
	oDialog.getControl ("cbxBest2").setVisible (False)
	oDialog.getControl ("cbxBest3").setVisible (False)
	oDialog.getControl ("lstMax").setVisible (False)
	
	oDialog.getControl ("imgPokemon").getModel.setPropertyValue ( _
		"ImageURL", fnGetImageUrl ("Unknown"))
	oDialog.getControl ("imgTeamLogo").getModel.setPropertyValue ( _
		"ImageURL", fnGetImageUrl ("Unknown"))
	
	fnLoadParamDialog = oDialog
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
	sImageId = "Pokemon" & maBaseStats (nSelected).sPokemonId
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
		fnGetResString ("AppraisalValorBest"), _
		CInt (fnGetResString ("AppraisalValorBestHeadWidth")))
	
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
		fnGetResString ("AppraisalMysticBest"), _
		CInt (fnGetResString ("AppraisalMysticBestHeadWidth")))
	
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
		fnGetResString ("AppraisalInstinctBest"), _
		CInt (fnGetResString ("AppraisalInstinctBestHeadWidth")))
	
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
Sub subUpdateBestStatAppraisal (oDialog As Object, _
		sAppraisal As String, nHeadWidth As Integer)
	Dim oText As Object, oList As Object, nX As Integer
	Dim sHead As String, sTail As String, nTailWidth As Integer
	Dim nDialogWidth As Integer
	Dim nPos As Integer
	Dim mItems () As String
	
	nPos = InStr (sAppraisal, "[Stat]")
	sHead = Left (sAppraisal, nPos - 1)
	sTail = Right (sAppraisal, _
		Len (sAppraisal) - nPos - Len ("[Stat]") + 1)
	nDialogWidth = oDialog.getModel.getPropertyValue ("Width")
	
	oText = oDialog.getControl ("txtBestHead")
	oText.getModel.setPropertyValue ("Width", nHeadWidth)
	oText.setVisible (True)
	oText.setText (sHead)
	nX = oText.getModel.getPropertyValue ("PositionX") + nHeadWidth
	
	mItems = Array ( _
		fnGetResString ("StatAttack"), _
		fnGetResString ("StatDefense"), _
		fnGetResString ("StatHP"))
	oList = oDialog.getControl ("lstBest")
	oList.removeItems (0, oList.getItemCount())
	oList.addItems (mItems, 0)
	oList.getModel.setPropertyValue ("PositionX", nX)
	oList.getModel.setPropertyValue ("Width", _
		CInt (fnGetResString ("BestStatWidth")))
	oList.setVisible (True)
	nX = nX + oList.getModel.getPropertyValue ("Width")
	
	nTailWidth = nDialogWidth - nX - 10
	oText = oDialog.getControl ("txtBestTail")
	oText.getModel.setPropertyValue ("PositionX", nX)
	oText.getModel.setPropertyValue ("Width", nTailWidth)
	oText.setVisible (True)
	oText.setText (sTail)
	
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
