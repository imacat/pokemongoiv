' 0Main: The main module for the Pokémon Go IV calculator
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
	sEvolveInto As String
	nEvolvedCP As Integer
	nMaxCP As Integer
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

Private maBaseStats () As New aStats
Private mCPM () As Double, mStarDust () As Integer

' subMain: The main program
Sub subMain
	Dim maIVs As Variant, nI As Integer, sOutput As String
	Dim aQuery As New aFindIVParam, aBaseStats As New aStats
	
	aQuery = fnAskParam
	If aQuery.bIsCancelled Then
		Exit Sub
	End If
	maIVs = fnFindIV (aQuery)
	sOutput = ""
	For nI = 0 To UBound (maIVs)
		sOutput = sOutput _
			& " Lv=" & maIVs (nI).fLevel _
			& " Atk=" & maIVs (nI).nAttack _
			& " Def=" & maIVs (nI).nDefense _
			& " Sta=" & maIVs (nI).nStamina _
			& " IV=" & fnFloor (maIVs (nI).nTotal * 100 / 45) & "%"
		If aQuery.sPokemon <> maIVs (nI).sEvolveInto Then
			aBaseStats = fnGetBaseStats (maIVs (nI).sEvolveInto)
			sOutput = sOutput & " Ev=" & maIVs (nI).sEvolveInto _
				& " " & maIVs (nI).nEvolvedCP
		End If
		If aQuery.nPlayerLevel <> 0 Then
			sOutput = sOutput & " XCP=" & maIVs (nI).nMaxCP
		End If
		sOutput = sOutput & Chr (10)
	Next nI
	If sOutput = "" Then
		MsgBox "Found no matching IV."
	Else
		subSaveIV (aQuery, maIVs)
	End If
End Sub

' fnAskParam: Asks the users for the parameters for the Pokémon.
Function fnAskParam As aFindIVParam
	Dim oDialog As Object, oDialogModel As Object
	Dim oTextModel As Object, oListModel As Object
	Dim oNumericModel As Object, oCheckBoxModel As Object
	Dim oGroupModel As Object, oButtonModel As Object
	Dim mListItems () As String, sTemp As String
	Dim nI As Integer, nCount As Integer
	Dim aQuery As New aFindIVParam
	
	' Creates a dialog
	oDialogModel = CreateUnoService ( _
		"com.sun.star.awt.UnoControlDialogModel")
	oDialogModel.setPropertyValue ("PositionX", 100)
	oDialogModel.setPropertyValue ("PositionY", 100)
	oDialogModel.setPropertyValue ("Height", 140)
	oDialogModel.setPropertyValue ("Width", 220)
	oDialogModel.setPropertyValue ("Title", "Pokémon Go IV Calculator")
	
	' Adds a text label for the Pokémon list.
	oTextModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlFixedTextModel")
	oTextModel.setPropertyValue ("PositionX", 5)
	oTextModel.setPropertyValue ("PositionY", 5)
	oTextModel.setPropertyValue ("Height", 12)
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
	oTextModel.setPropertyValue ("PositionY", 20)
	oTextModel.setPropertyValue ("Height", 12)
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
	oTextModel.setPropertyValue ("PositionY", 20)
	oTextModel.setPropertyValue ("Height", 12)
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
	oTextModel.setPropertyValue ("PositionY", 20)
	oTextModel.setPropertyValue ("Height", 12)
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
	oTextModel.setPropertyValue ("PositionY", 20)
	oTextModel.setPropertyValue ("Height", 12)
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
	oCheckBoxModel.setPropertyValue ("PositionY", 35)
	oCheckBoxModel.setPropertyValue ("Height", 12)
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
	oGroupModel.setPropertyValue ("Height", 65)
	oGroupModel.setPropertyValue ("Width", 210)
	oGroupModel.setPropertyValue ("Label", "Apprasals")
	oDialogModel.insertByName ("grpApprasals", oGroupModel)
	
	' Adds the first appraisal list.
	mListItems = Array ( _
		"1. Amazed me/wonder/best", _
		"2. Strong/caught my attention", _
		"3. Decent/above average", _
		"4. Not great/not make headway/has room")
	oListModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlListBoxModel")
	oListModel.setPropertyValue ("PositionX", 10)
	oListModel.setPropertyValue ("PositionY", 64)
	oListModel.setPropertyValue ("Height", 12)
	oListModel.setPropertyValue ("Width", 200)
	oListModel.setPropertyValue ("Dropdown", True)
	oListModel.setPropertyValue ("StringItemList", mListItems)
	oDialogModel.insertByName ("lstApprasal1", oListModel)
	
	' Adds a text label for the HP field.
	oTextModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlFixedTextModel")
	oTextModel.setPropertyValue ("PositionX", 10)
	oTextModel.setPropertyValue ("PositionY", 80)
	oTextModel.setPropertyValue ("Height", 12)
	oTextModel.setPropertyValue ("Width", 15)
	oTextModel.setPropertyValue ("Label", "Best:")
	oDialogModel.insertByName ("txtBest", oTextModel)
	
	' Adds the attack is best check box
	oCheckBoxModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlCheckBoxModel")
	oCheckBoxModel.setPropertyValue ("PositionX", 25)
	oCheckBoxModel.setPropertyValue ("PositionY", 80)
	oCheckBoxModel.setPropertyValue ("Height", 12)
	oCheckBoxModel.setPropertyValue ("Width", 30)
	oCheckBoxModel.setPropertyValue ("Label", "~Attack")
	oDialogModel.insertByName ("cbxAttackBest", oCheckBoxModel)
	
	' Adds the defense is best check box
	oCheckBoxModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlCheckBoxModel")
	oCheckBoxModel.setPropertyValue ("PositionX", 55)
	oCheckBoxModel.setPropertyValue ("PositionY", 80)
	oCheckBoxModel.setPropertyValue ("Height", 12)
	oCheckBoxModel.setPropertyValue ("Width", 35)
	oCheckBoxModel.setPropertyValue ("Label", "~Defense")
	oDialogModel.insertByName ("cbxDefenseBest", oCheckBoxModel)
	
	' Adds the defense is best check box
	oCheckBoxModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlCheckBoxModel")
	oCheckBoxModel.setPropertyValue ("PositionX", 90)
	oCheckBoxModel.setPropertyValue ("PositionY", 80)
	oCheckBoxModel.setPropertyValue ("Height", 12)
	oCheckBoxModel.setPropertyValue ("Width", 45)
	oCheckBoxModel.setPropertyValue ("Label", "HP (~Stamina)")
	oDialogModel.insertByName ("cbxHPBest", oCheckBoxModel)
	
	' Adds the second appraisal list.
	mListItems = Array ( _
		"1. WOW/incredible/stats are best", _
		"2. Excellent/impressed/impressive", _
		"3. Get the job done/noticeable/some good stats", _
		"4. No greatness/not out of the norm/kinda basic")
	oListModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlListBoxModel")
	oListModel.setPropertyValue ("PositionX", 10)
	oListModel.setPropertyValue ("PositionY", 95)
	oListModel.setPropertyValue ("Height", 12)
	oListModel.setPropertyValue ("Width", 200)
	oListModel.setPropertyValue ("Dropdown", True)
	oListModel.setPropertyValue ("StringItemList", mListItems)
	oDialogModel.insertByName ("lstApprasal2", oListModel)
	
	' Adds the OK button.
	oButtonModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlButtonModel")
	oButtonModel.setPropertyValue ("PositionX", 35)
	oButtonModel.setPropertyValue ("PositionY", 120)
	oButtonModel.setPropertyValue ("Height", 15)
	oButtonModel.setPropertyValue ("Width", 60)
	oButtonModel.setPropertyValue ("PushButtonType", _
		com.sun.star.awt.PushButtonType.OK)
	oButtonModel.setPropertyValue ("DefaultButton", True)
	oDialogModel.insertByName ("btnOK", oButtonModel)
	
	' Adds the cancel button.
	oButtonModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlButtonModel")
	oButtonModel.setPropertyValue ("PositionX", 125)
	oButtonModel.setPropertyValue ("PositionY", 120)
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
	aQuery.sBest = ""
	If oDialog.getControl ("cbxAttackBest").getState = 1 Then
		aQuery.sBest = aQuery.sBest & "Atk "
	End If
	If oDialog.getControl ("cbxDefenseBest").getState = 1 Then
		aQuery.sBest = aQuery.sBest & "Def "
	End If
	If oDialog.getControl ("cbxHPBest").getState = 1 Then
		aQuery.sBest = aQuery.sBest & "Sta "
	End If
	fnAskParam = aQuery
End Function

' fnFindIV: Finds the possible individual values of the Pokémon
Function fnFindIV (aQuery As aFindIVParam) As Variant
	Dim aBaseStats As New aStats, maIV () As New aStats
	Dim fLevel As Double, nStamina As Integer
	Dim nAttack As Integer, nDefense As integer
	Dim nI As Integer, nJ As Integer
	Dim fStep As Double, nCount As Integer
	Dim aEvBaseStats As new aStats, aTempIV As New aStats
	
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
	aEvBaseStats = fnGetBaseStats (aBaseStats.sEvolveInto)
	subReadStarDust
	nCount = -1
	For fLevel = 1 To UBound (mStarDust) Step fStep
		If mStarDust (CInt (fLevel - 0.5)) = aQuery.nStarDust Then
	'For nI = 0 To UBound (maStarDust) Step nStep
	'	fLevel = maStarDust (nI).fLevel
	'	If maStarDust (nI).nStarDust = aQuery.nStarDust Then
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
									.sEvolveInto = aBaseStats.sEvolveInto
									.nEvolvedCP = fnCalcCP (aEvBaseStats, fLevel, nAttack, nDefense, nStamina)
								End With
								If aQuery.nPlayerLevel <> 0 Then
									maIV (nCount).nMaxCP = fnCalcCP (aEvBaseStats, aQuery.nPlayerLevel + 1.5, nAttack, nDefense, nStamina)
								Else
									maIV (nCount).nMaxCP = -1
								End If
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
	fnCompareIV = aIVb.nMaxCP - aIVa.nMaxCP
	If fnCompareIV <> 0 Then
		Exit Function
	End If
	fnCompareIV = aIVb.nEvolvedCP - aIVa.nEvolvedCP
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
	With aTo
		.sNo = aFrom.sNo
		.sPokemon = aFrom.sPokemon
		.fLevel = aFrom.fLevel
		.nAttack = aFrom.nAttack
		.nDefense = aFrom.nDefense
		.nStamina = aFrom.nStamina
		.nTotal = aFrom.nTotal
		.sEvolveInto = aFrom.sEvolveInto
		.nEvolvedCP = aFrom.nEvolvedCP
		.nMaxCP = aFrom.nMaxCP
	End With
End Function

' subSaveIV: Saves the found IV
Sub subSaveIV (aQuery As aFindIVParam, maIVs () As aStats)
	Dim oDoc As Object, oSheet As Object, oRange As Object
	Dim nI As Integer, oColumns As Object
	Dim mData (Ubound (maIVs) + 1) As Variant
	Dim mProps () As New com.sun.star.beans.PropertyValue
	
	oDoc = StarDesktop.loadComponentFromURL ( _
		"private:factory/scalc", "_default", 0, mProps)
	oSheet = oDoc.getSheets.getByIndex (0)
	mData (0) = Array ( _
		"No", "Pokemon", "CP", "HP", _
		"Lv", "Atk", "Def", "Sta", "IV", _
		"Evolve Into", "Evolved CP", "Max CP")
	mData (1) = Array ( _
		maIVs (0).sNo, aQuery.sPokemon, aQuery.nCP, aQuery.nHP, _
		maIVs (0).fLevel, maIVs (0).nAttack, maIVs (0).nDefense, _
		maIVs (0).nStamina, maIVs (0).nTotal / 45, _
		maIVs (0).sEvolveInto, maIVs (0).nEvolvedCP, _
		maIVs (0).nMaxCP)
	For nI = 1 To UBound (maIVs)
		mData (nI + 1) = Array ( _
			"", "", "", "", _
			maIVs (nI).fLevel, maIVs (nI).nAttack, maIVs (nI).nDefense, _
			maIVs (nI).nStamina, maIVs (nI).nTotal / 45, _
			maIVs (nI).sEvolveInto, maIVs (nI).nEvolvedCP, _
			maIVs (nI).nMaxCP)
	Next nI
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
		8, 1, 8, UBound (mData))
	oRange.setPropertyValue ("NumberFormat", 10)
	
	oColumns = oSheet.getColumns
	For nI = 0 To UBound (mData (0))
		oColumns.getByIndex (nI).setPropertyValue ( _
			"OptimalWidth", True)
	Next nI
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
	Dim mData As Variant, nI As Integer
	
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
				.sEvolveInto = mData (nI) (5)
			End With
		Next nI
	End If
End Sub

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
