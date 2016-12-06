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
		MsgBox "Found no matching IV."
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
	oDialog.getControl ("lstApprasal1").setVisible (False)
	oDialog.getControl ("txtBestBefore").setVisible (False)
	oDialog.getControl ("lstBest").setVisible (False)
	oDialog.getControl ("txtBestAfter").setVisible (False)
	oDialog.getControl ("cbxBest2").setVisible (False)
	oDialog.getControl ("cbxBest3").setVisible (False)
	oDialog.getControl ("lstApprasal2").setVisible (False)
	
	' TODO: To be removed.
	oDialog.getControl ("imgTeam").getModel.setPropertyValue ("ImageURL", fnGetImageUrl ("TeamValor"))
	
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
		.nAppraisal1 = oDialog.getControl ("lstApprasal1").getSelectedItemPos + 1
		.nAppraisal2 = oDialog.getControl ("lstApprasal2").getSelectedItemPos + 1
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

' fnAskParam: Asks the users for the parameters for the Pokémon.
Function fnAskParam0 As aFindIVParam
	Dim oDialog As Object, oDialogModel As Object
	Dim oTextModel As Object, oListModel As Object
	Dim oNumericModel As Object, oCheckBoxModel As Object
	Dim oGroupModel As Object, oRadioModel As Object
	Dim oButtonModel As Object, oListener As Object
	Dim mListItems () As String, sTemp As String
	Dim nI As Integer, nCount As Integer
	Dim bIsBestAttack As Boolean, bIsBestDefense As Boolean
	Dim bIsBestHP As Boolean
	Dim aQuery As New aFindIVParam
	
	' Creates a dialog
	oDialogModel = CreateUnoService ( _
		"com.sun.star.awt.UnoControlDialogModel")
	oDialogModel.setPropertyValue ("PositionX", 100)
	oDialogModel.setPropertyValue ("PositionY", 100)
	oDialogModel.setPropertyValue ("Height", 185)
	oDialogModel.setPropertyValue ("Width", 220)
	oDialogModel.setPropertyValue ("Title", "Pokémon GO IV Calculator")
	
	' Adds a text label for the Pokémon list.
	oTextModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlFixedTextModel")
	oTextModel.setPropertyValue ("PositionX", 5)
	oTextModel.setPropertyValue ("PositionY", 6)
	oTextModel.setPropertyValue ("Height", 8)
	oTextModel.setPropertyValue ("Width", 30)
	oTextModel.setPropertyValue ("Label", "~Pokémon:")
	oDialogModel.insertByName ("txtPokemon", oTextModel)
	
	' Adds the Pokémon list.
	subReadBaseStats
	ReDim mListItems (UBound (maBaseStats)) As String
	For nI = 0 To UBound (maBaseStats)
		mListItems (nI) = maBaseStats (nI).sPokemon
	Next nI
	oListModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlListBoxModel")
	oListModel.setPropertyValue ("PositionX", 35)
	oListModel.setPropertyValue ("PositionY", 4)
	oListModel.setPropertyValue ("Height", 12)
	oListModel.setPropertyValue ("Width", 50)
	oListModel.setPropertyValue ("TabIndex", 0)
	oListModel.setPropertyValue ("Dropdown", True)
	oListModel.setPropertyValue ("StringItemList", mListItems)
	oDialogModel.insertByName ("lstPokemon", oListModel)
	
	' Adds a text label for the CP field.
	oTextModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlFixedTextModel")
	oTextModel.setPropertyValue ("PositionX", 5)
	oTextModel.setPropertyValue ("PositionY", 21)
	oTextModel.setPropertyValue ("Height", 8)
	oTextModel.setPropertyValue ("Width", 15)
	oTextModel.setPropertyValue ("Label", "~CP:")
	oDialogModel.insertByName ("txtCP", oTextModel)
	
	' Adds the CP field.
	oNumericModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlNumericFieldModel")
	oNumericModel.setPropertyValue ("PositionX", 20)
	oNumericModel.setPropertyValue ("PositionY", 19)
	oNumericModel.setPropertyValue ("Height", 12)
	oNumericModel.setPropertyValue ("Width", 20)
	oNumericModel.setPropertyValue ("DecimalAccuracy", 0)
	oDialogModel.insertByName ("numCP", oNumericModel)
	
	' Adds a text label for the HP field.
	oTextModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlFixedTextModel")
	oTextModel.setPropertyValue ("PositionX", 50)
	oTextModel.setPropertyValue ("PositionY", 21)
	oTextModel.setPropertyValue ("Height", 8)
	oTextModel.setPropertyValue ("Width", 15)
	oTextModel.setPropertyValue ("Label", "~HP:")
	oDialogModel.insertByName ("txtHP", oTextModel)
	
	' Adds the HP field.
	oNumericModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlNumericFieldModel")
	oNumericModel.setPropertyValue ("PositionX", 65)
	oNumericModel.setPropertyValue ("PositionY", 19)
	oNumericModel.setPropertyValue ("Height", 12)
	oNumericModel.setPropertyValue ("Width", 15)
	oNumericModel.setPropertyValue ("DecimalAccuracy", 0)
	oDialogModel.insertByName ("numHP", oNumericModel)
	
	' Adds a text label for the star dust field.
	oTextModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlFixedTextModel")
	oTextModel.setPropertyValue ("PositionX", 90)
	oTextModel.setPropertyValue ("PositionY", 21)
	oTextModel.setPropertyValue ("Height", 8)
	oTextModel.setPropertyValue ("Width", 30)
	oTextModel.setPropertyValue ("Label", "S~tar dust:")
	oDialogModel.insertByName ("txtStarDust", oTextModel)
	
	' Adds the star dust field.
	subReadStarDust
	sTemp = " "
	ReDim mListItems () As String
	nCount = -1
	For nI = 1 To UBound (mStarDust)
		If InStr (sTemp, " " & CStr (mStarDust (nI)) & " ") = 0 Then
			nCount = nCount + 1
			ReDim Preserve mListItems (nCount) As String
			mListItems (nCount) = CStr (mStarDust (nI))
			sTemp = sTemp & CStr (mStarDust (nI)) & " "
		End If
	Next nI
	oListModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlListBoxModel")
	oListModel.setPropertyValue ("PositionX", 120)
	oListModel.setPropertyValue ("PositionY", 19)
	oListModel.setPropertyValue ("Height", 12)
	oListModel.setPropertyValue ("Width", 30)
	oListModel.setPropertyValue ("Dropdown", True)
	oListModel.setPropertyValue ("StringItemList", mListItems)
	oDialogModel.insertByName ("lstStarDust", oListModel)
	
	' Adds a text label for the player level field.
	oTextModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlFixedTextModel")
	oTextModel.setPropertyValue ("PositionX", 160)
	oTextModel.setPropertyValue ("PositionY", 21)
	oTextModel.setPropertyValue ("Height", 8)
	oTextModel.setPropertyValue ("Width", 35)
	oTextModel.setPropertyValue ("Label", "Player ~level:")
	oDialogModel.insertByName ("txtPlayerLevel", oTextModel)
	
	' Adds the player level field.
	ReDim mListItems (39) As String
	For nI = 0 To UBound (mListItems)
		mListItems (nI) = CStr (nI + 1)
	Next nI
	oListModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlListBoxModel")
	oListModel.setPropertyValue ("PositionX", 195)
	oListModel.setPropertyValue ("PositionY", 19)
	oListModel.setPropertyValue ("Height", 12)
	oListModel.setPropertyValue ("Width", 20)
	oListModel.setPropertyValue ("Dropdown", True)
	oListModel.setPropertyValue ("StringItemList", mListItems)
	oDialogModel.insertByName ("lstPlayerLevel", oListModel)
	
	' Adds the whether powered-up check box.
	oCheckBoxModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlCheckBoxModel")
	oCheckBoxModel.setPropertyValue ("PositionX", 5)
	oCheckBoxModel.setPropertyValue ("PositionY", 36)
	oCheckBoxModel.setPropertyValue ("Height", 8)
	oCheckBoxModel.setPropertyValue ("Width", 210)
	oCheckBoxModel.setPropertyValue ("Label", _
		"This Pokémon is ~newly-caught and was not powered-up yet.")
	oCheckBoxModel.setPropertyValue ("State", 1)
	oDialogModel.insertByName ("cbxIsNew", oCheckBoxModel)
	
	' Adds a group for the appraisals
	oGroupModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlGroupBoxModel")
	oGroupModel.setPropertyValue ("PositionX", 5)
	oGroupModel.setPropertyValue ("PositionY", 50)
	oGroupModel.setPropertyValue ("Height", 110)
	oGroupModel.setPropertyValue ("Width", 210)
	oGroupModel.setPropertyValue ("Label", "Team Leader Apprasal")
	oDialogModel.insertByName ("grpApprasals", oGroupModel)
	
	' Adds a text label for the team.
	oTextModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlFixedTextModel")
	oTextModel.setPropertyValue ("PositionX", 10)
	oTextModel.setPropertyValue ("PositionY", 66)
	oTextModel.setPropertyValue ("Height", 8)
	oTextModel.setPropertyValue ("Width", 20)
	oTextModel.setPropertyValue ("Label", "Team:")
	oDialogModel.insertByName ("txtTeam", oTextModel)
	
	' Adds the red team radio button.
	oRadioModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlRadioButtonModel")
	oRadioModel.setPropertyValue ("PositionX", 30)
	oRadioModel.setPropertyValue ("PositionY", 66)
	oRadioModel.setPropertyValue ("Height", 8)
	oRadioModel.setPropertyValue ("Width", 25)
	oRadioModel.setPropertyValue ("Label", "~Valor")
	oRadioModel.setPropertyValue ("TextColor", RGB (255, 255, 255))
	oRadioModel.setPropertyValue ("BackgroundColor", RGB (255, 0, 0))
	oDialogModel.insertByName ("rdoTeamRed", oRadioModel)
	
	' Adds the blue team radio button.
	oRadioModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlRadioButtonModel")
	oRadioModel.setPropertyValue ("PositionX", 60)
	oRadioModel.setPropertyValue ("PositionY", 66)
	oRadioModel.setPropertyValue ("Height", 8)
	oRadioModel.setPropertyValue ("Width", 30)
	oRadioModel.setPropertyValue ("Label", "~Mystic")
	oRadioModel.setPropertyValue ("TextColor", RGB (255, 255, 255))
	oRadioModel.setPropertyValue ("BackgroundColor", RGB (0, 0, 255))
	oDialogModel.insertByName ("rdoTeamBlue", oRadioModel)
	
	' Adds the yellow team radio button.
	oRadioModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlRadioButtonModel")
	oRadioModel.setPropertyValue ("PositionX", 95)
	oRadioModel.setPropertyValue ("PositionY", 66)
	oRadioModel.setPropertyValue ("Height", 8)
	oRadioModel.setPropertyValue ("Width", 30)
	oRadioModel.setPropertyValue ("Label", "~Instinct")
	oRadioModel.setPropertyValue ("BackgroundColor", RGB (255, 255, 0))
	oDialogModel.insertByName ("rdoTeamYellow", oRadioModel)
	
	' Adds the first appraisal list.
	oListModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlListBoxModel")
	oListModel.setPropertyValue ("PositionX", 10)
	oListModel.setPropertyValue ("PositionY", 79)
	oListModel.setPropertyValue ("Height", 12)
	oListModel.setPropertyValue ("Width", 200)
	oListModel.setPropertyValue ("Dropdown", True)
	oDialogModel.insertByName ("lstApprasal1", oListModel)
	
	' Adds a text label before the best stat.
	oTextModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlFixedTextModel")
	oTextModel.setPropertyValue ("PositionX", 10)
	oTextModel.setPropertyValue ("PositionY", 96)
	oTextModel.setPropertyValue ("Height", 8)
	oTextModel.setPropertyValue ("Width", 20)
	oDialogModel.insertByName ("txtBestBefore", oTextModel)
	
	' Adds the best stat field.
	oListModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlListBoxModel")
	oListModel.setPropertyValue ("PositionX", 30)
	oListModel.setPropertyValue ("PositionY", 94)
	oListModel.setPropertyValue ("Height", 12)
	oListModel.setPropertyValue ("Width", 35)
	oListModel.setPropertyValue ("Dropdown", True)
	oDialogModel.insertByName ("lstBest", oListModel)
	
	' Adds a text label after the best stat.
	oTextModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlFixedTextModel")
	oTextModel.setPropertyValue ("PositionX", 65)
	oTextModel.setPropertyValue ("PositionY", 96)
	oTextModel.setPropertyValue ("Height", 8)
	oTextModel.setPropertyValue ("Width", 100)
	oDialogModel.insertByName ("txtBestAfter", oTextModel)
	
	' Adds the second best stat check box.
	oCheckBoxModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlCheckBoxModel")
	oCheckBoxModel.setPropertyValue ("PositionX", 10)
	oCheckBoxModel.setPropertyValue ("PositionY", 111)
	oCheckBoxModel.setPropertyValue ("Height", 8)
	oCheckBoxModel.setPropertyValue ("Width", 200)
	oDialogModel.insertByName ("cbxBest2", oCheckBoxModel)
	
	' Adds the third best stat check box.
	oCheckBoxModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlCheckBoxModel")
	oCheckBoxModel.setPropertyValue ("PositionX", 10)
	oCheckBoxModel.setPropertyValue ("PositionY", 126)
	oCheckBoxModel.setPropertyValue ("Height", 8)
	oCheckBoxModel.setPropertyValue ("Width", 200)
	oDialogModel.insertByName ("cbxBest3", oCheckBoxModel)
	
	' Adds the second appraisal list.
	oListModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlListBoxModel")
	oListModel.setPropertyValue ("PositionX", 10)
	oListModel.setPropertyValue ("PositionY", 139)
	oListModel.setPropertyValue ("Height", 12)
	oListModel.setPropertyValue ("Width", 200)
	oListModel.setPropertyValue ("Dropdown", True)
	oDialogModel.insertByName ("lstApprasal2", oListModel)
	
	' Adds the OK button.
	oButtonModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlButtonModel")
	oButtonModel.setPropertyValue ("PositionX", 35)
	oButtonModel.setPropertyValue ("PositionY", 165)
	oButtonModel.setPropertyValue ("Height", 15)
	oButtonModel.setPropertyValue ("Width", 60)
	oButtonModel.setPropertyValue ("PushButtonType", _
		com.sun.star.awt.PushButtonType.OK)
	oButtonModel.setPropertyValue ("DefaultButton", True)
	oButtonModel.setPropertyValue ("Enabled", False)
	oDialogModel.insertByName ("btnOK", oButtonModel)
	
	' Adds the cancel button.
	oButtonModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlButtonModel")
	oButtonModel.setPropertyValue ("PositionX", 125)
	oButtonModel.setPropertyValue ("PositionY", 165)
	oButtonModel.setPropertyValue ("Height", 15)
	oButtonModel.setPropertyValue ("Width", 60)
	oButtonModel.setPropertyValue ("PushButtonType", _
		com.sun.star.awt.PushButtonType.CANCEL)
	oDialogModel.insertByName ("btnCancel", oButtonModel)
	
	' Adds the dialog model to the control and runs it.
	oDialog = CreateUnoService ("com.sun.star.awt.UnoControlDialog")
	oDialog.setModel (oDialogModel)
	oDialog.setVisible (True)
	oDialog.getControl ("lstPokemon").setFocus
	oListener = CreateUnoListener ("subBtnOKCheck_", _
		"com.sun.star.awt.XItemListener")
	oDialog.getControl ("lstPokemon").addItemListener (oListener)
	oListener = CreateUnoListener ("subBtnOKCheck_", _
		"com.sun.star.awt.XTextListener")
	oDialog.getControl ("numCP").addTextListener (oListener)
	oListener = CreateUnoListener ("subBtnOKCheck_", _
		"com.sun.star.awt.XTextListener")
	oDialog.getControl ("numHP").addTextListener (oListener)
	oListener = CreateUnoListener ("subBtnOKCheck_", _
		"com.sun.star.awt.XItemListener")
	oDialog.getControl ("lstStarDust").addItemListener (oListener)
	oListener = CreateUnoListener ("subRdoTeamRedItemChanged_", _
		"com.sun.star.awt.XItemListener")
	oDialog.getControl ("rdoTeamRed").addItemListener (oListener)
	oListener = CreateUnoListener ("subRdoTeamBlueItemChanged_", _
		"com.sun.star.awt.XItemListener")
	oDialog.getControl ("rdoTeamBlue").addItemListener (oListener)
	oListener = CreateUnoListener ("subRdoTeamYellowItemChanged_", _
		"com.sun.star.awt.XItemListener")
	oDialog.getControl ("rdoTeamYellow").addItemListener (oListener)
	oListener = CreateUnoListener ("subLstBestItemChanged_", _
		"com.sun.star.awt.XItemListener")
	oDialog.getControl ("lstBest").addItemListener (oListener)
	oDialog.getControl ("lstApprasal1").setVisible (False)
	oDialog.getControl ("txtBestBefore").setVisible (False)
	oDialog.getControl ("lstBest").setVisible (False)
	oDialog.getControl ("txtBestAfter").setVisible (False)
	oDialog.getControl ("cbxBest2").setVisible (False)
	oDialog.getControl ("cbxBest3").setVisible (False)
	oDialog.getControl ("lstApprasal2").setVisible (False)
	If oDialog.execute = 0 Then
		aQuery.bIsCancelled = True
		fnAskParam0 = aQuery
		Exit Function
	End If
	
	With aQuery
		.sPokemon = oDialog.getControl ("lstPokemon").getSelectedItem
		.nCP = oDialog.getControl ("numCP").getValue
		.nHP = oDialog.getControl ("numHP").getValue
		.nStarDust = CInt (oDialog.getControl ("lstStarDust").getSelectedItem)
		.nPlayerLevel = CInt (oDialog.getControl ("lstPlayerLevel").getSelectedItem)
		.nAppraisal1 = oDialog.getControl ("lstApprasal1").getSelectedItemPos + 1
		.nAppraisal2 = oDialog.getControl ("lstApprasal2").getSelectedItemPos + 1
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
	
	fnAskParam0 = aQuery
End Function

' subBtnOKCheck_disposing: Dummy for the listener.
Sub subBtnOKCheck_disposing (oEvent As object)
End Sub

' subBtnOKCheck_itemStateChanged: When the Pokémon or star dust is selected.
Sub subBtnOKCheck_itemStateChanged (oEvent As object)
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

' subBtnOKCheck_textChanged: When the CP or HP is filled
Sub subBtnOKCheck_textChanged (oEvent As object)
	subBtnOKCheck_itemStateChanged (oEvent)
End Sub

' subRdoTeamRedItemChanged_disposing: Dummy for the listener.
Sub subRdoTeamRedItemChanged_disposing (oEvent As object)
End Sub

' subRdoTeamRedItemChanged_itemStateChanged: When the team is selected.
Sub subRdoTeamRedItemChanged_itemStateChanged (oEvent As object)
	Dim oDialog As Object, oList As Object, oText As Object
	Dim mItems () As String
	
	oDialog = oEvent.Source.getContext
	
	mItems = Array ( _
		"Overall, your [Pokémon] simply amazes me. It can accomplish anything!", _
		"Overall, your [Pokémon] is a strong Pokémon. You should be proud!", _
		"Overall, your [Pokémon] is a decent Pokémon.", _
		"Overall, your [Pokémon] may not be great in battle, but I still like it!")
	oList = oDialog.getControl ("lstApprasal1")
	oList.removeItems (0, oList.getItemCount())
	oList.addItems (mItems, 0)
	oList.setVisible (True)
	
	oText = oDialog.getControl ("txtBestBefore")
	oText.setPosSize (30, 96, 20, 8, _
		com.sun.star.awt.PosSize.X + com.sun.star.awt.PosSize.WIDTH)
	oText.setVisible (True)
	oText.setText ("Its")
	
	mItems = Array ("Attack", "Defense", "HP")
	oList = oDialog.getControl ("lstBest")
	oList.removeItems (0, oList.getItemCount())
	oList.addItems (mItems, 0)
	oList.setPosSize (50, 96, 35, 8, _
		com.sun.star.awt.PosSize.X)
	oList.setVisible (True)
	
	oText = oDialog.getControl ("txtBestAfter")
	oText.setPosSize (145, 96, 160, 8, _
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
	oList = oDialog.getControl ("lstApprasal2")
	oList.removeItems (0, oList.getItemCount())
	oList.addItems (mItems, 0)
	oList.setVisible (True)
End Sub

' subRdoTeamBlueItemChanged_disposing: Dummy for the listener.
Sub subRdoTeamBlueItemChanged_disposing (oEvent As object)
End Sub

' subRdoTeamBlueItemChanged_itemStateChanged: When the blue team is selected.
Sub subRdoTeamBlueItemChanged_itemStateChanged (oEvent As object)
	Dim oDialog As Object, oList As Object, oText As Object
	Dim mItems () As String
	
	oDialog = oEvent.Source.getContext
	
	mItems = Array ( _
		"Overall, your [Pokémon] is a wonder! What a breathtaking Pokémon!", _
		"Overall, your [Pokémon] has certainly caught my attention.", _
		"Overall, your [Pokémon] is above average.", _
		"Overall, your [Pokémon] is not likely to make much headway in battle.")
	oList = oDialog.getControl ("lstApprasal1")
	oList.removeItems (0, oList.getItemCount())
	oList.addItems (mItems, 0)
	oList.setVisible (True)
	
	oText = oDialog.getControl ("txtBestBefore")
	oText.setPosSize (30, 96, 200, 8, _
		com.sun.star.awt.PosSize.X + com.sun.star.awt.PosSize.WIDTH)
	oText.setVisible (True)
	oText.setText ("I see that its best attribute is its")
	
	mItems = Array ("Attack", "Defense", "HP")
	oList = oDialog.getControl ("lstBest")
	oList.removeItems (0, oList.getItemCount())
	oList.addItems (mItems, 0)
	oList.setPosSize (230, 96, 35, 8, _
		com.sun.star.awt.PosSize.X)
	oList.setVisible (True)
	
	oText = oDialog.getControl ("txtBestAfter")
	oText.setPosSize (325, 96, 5, 8, _
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
	oList = oDialog.getControl ("lstApprasal2")
	oList.removeItems (0, oList.getItemCount())
	oList.addItems (mItems, 0)
	oList.setVisible (True)
End Sub

' subRdoTeamYellowItemChanged_disposing: Dummy for the listener.
Sub subRdoTeamYellowItemChanged_disposing (oEvent As object)
End Sub

' subRdoTeamYellowItemChanged_itemStateChanged: When the yellow team is selected.
Sub subRdoTeamYellowItemChanged_itemStateChanged (oEvent As object)
	Dim oDialog As Object, oList As Object, oText As Object
	Dim mItems () As String
	
	oDialog = oEvent.Source.getContext
	
	mItems = Array ( _
		"Overall, your [Pokémon] looks like it can really battle with the best of them!", _
		"Overall, your [Pokémon] is really strong!", _
		"Overall, your [Pokémon] is pretty decent!", _
		"Overall, your [Pokémon] has room for improvement as far as battling goes.")
	oList = oDialog.getControl ("lstApprasal1")
	oList.removeItems (0, oList.getItemCount())
	oList.addItems (mItems, 0)
	oList.setVisible (True)
	
	oText = oDialog.getControl ("txtBestBefore")
	oText.setPosSize (30, 96, 115, 8, _
		com.sun.star.awt.PosSize.X + com.sun.star.awt.PosSize.WIDTH)
	oText.setVisible (True)
	oText.setText ("Its best quality is")
	
	mItems = Array ("Attack", "Defense", "HP")
	oList = oDialog.getControl ("lstBest")
	oList.removeItems (0, oList.getItemCount())
	oList.addItems (mItems, 0)
	oList.setPosSize (145, 96, 35, 8, _
		com.sun.star.awt.PosSize.X)
	oList.setVisible (True)
	
	oText = oDialog.getControl ("txtBestAfter")
	oText.setPosSize (240, 96, 5, 8, _
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
	oList = oDialog.getControl ("lstApprasal2")
	oList.removeItems (0, oList.getItemCount())
	oList.addItems (mItems, 0)
	oList.setVisible (True)
End Sub

' subLstBestItemChanged_disposing: Dummy for the listener.
Sub subLstBestItemChanged_disposing (oEvent As object)
End Sub

' subLstBestItemChanged_itemStateChanged: When the best stat is selected.
Sub subLstBestItemChanged_itemStateChanged (oEvent As object)
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
