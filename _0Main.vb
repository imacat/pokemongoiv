' _0Main: The main module for the Pokémon Go IV calculator
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

' The combat power multiplier.
Type aCPM
	fLevel As Double
	fCPM As Double
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
End Type

Private maBaseStats () As New aStats, maCPM () As New aCPM
Private maStarDust () As New aStarDust

' subMain: The main program
Sub subMain
	BasicLibraries.loadLibrary "XrayTool"
	Dim maIVs As Variant, nI As Integer, sOutput As String
	Dim aQuery As New aFindIVParam, aBaseStats As New aStats
	
	aQuery = fnAskParam
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

' fnFindIV: Finds the possible individual values of the Pokémon
Function fnFindIV (aQuery As aFindIVParam) As Variant
	Dim aBaseStats As New aStats, maIV () As New aStats
	Dim fLevel As Double, nStamina As Integer
	Dim nAttack As Integer, nDefense As integer
	Dim nI As Integer, nJ As Integer
	Dim nStep As Integer, nCount As Integer
	Dim aEvBaseStats As new aStats, aTempIV As New aStats
	
	If aQuery.sPokemon = "" Then
		fnFindIV = maIV
		Exit Function
	End If
	If aQuery.bIsNew Then
		nStep = 2
	Else
		nStep = 1
	End If
	aBaseStats = fnGetBaseStats (aQuery.sPokemon)
	aEvBaseStats = fnGetBaseStats (aBaseStats.sEvolveInto)
	subReadStarDust
	nCount = -1
	For nI = 0 To UBound (maStarDust) Step nStep
		fLevel = maStarDust (nI).fLevel
		If maStarDust (nI).nStarDust = aQuery.nStarDust Then
			For nStamina = 0 To 15
				If fnCalcHP (aBaseStats, fLevel, nStamina) = aQuery.nHP Then
					For nAttack = 0 To 15
						For nDefense = 0 To 15
							If fnCalcCP (aBaseStats, fLevel, nAttack, nDefense, nStamina) = aQuery.nCP _
									And Not (fnFilterAppraisals (aQuery, nAttack, nDefense, nStamina)) Then
								nCount = nCount +  1
								ReDim Preserve maIV (nCount) As New aStats
								With  maIV (nCount)
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
	Next nI
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
End  Function

' fnGetCPM: Returns the combat power multiplier.
Function fnGetCPM (fLevel As Double) As Double
	Dim nI As Integer
	
	subReadCPM
	For nI = 0 To UBound (maCPM)
		If maCPM (nI).fLevel = fLevel Then
			fnGetCPM = maCPM (nI).fCPM
			Exit Function
		End If
	Next nI
End  Function

' fnFloor: Returns the floor of the number
Function fnFloor (fNumber As Double) As Integer
	fnFloor = CInt (fNumber - 0.5)
End Function

' subReadBaseStats: Reads the base stats table.
Sub subReadBaseStats
	Dim oSheet As Object, oRange As Object, mData As Variant
	Dim nCount As Integer, nRow As Integer, nColumn As Integer
	
	If UBound (maBaseStats) = -1 Then
		oSheet = ThisComponent.getSheets.getByName ("basestat")
		oRange = oSheet.getCellRangeByName ("BaseStats")
		mData = oRange.getDataArray
		nCount = -1
		For nRow = 1 To UBound (mData) - 1
			nCount = nCount + 1
			ReDim Preserve maBaseStats (nCount) As New aStats
			With maBaseStats (nCount)
				.sNo = mData (nRow) (1)
				.sPokemon = mData (nRow) (0)
				.nStamina = mData (nRow) (3)
				.nAttack = mData (nRow) (4)
				.nDefense = mData (nRow) (5)
			End With
			For nColumn = 9 To 7 Step -1
				If mData (nRow) (nColumn) <> "" Then
					maBaseStats (nCount).sEvolveInto = mData (nRow) (nColumn)
					nColumn = 6
				End If
			Next nColumn
		Next nRow
	End If
End  Sub

' subReadCPM: Reads the CPM table.
Sub subReadCPM
	Dim oSheet As Object, oRange As Object, mData As Variant
	Dim nCount As Integer, nRow As Integer
	
	If UBound (maCPM) = -1 Then
		oSheet = ThisComponent.getSheets.getByName ("cpm")
		oRange = oSheet.getCellRangeByName ("CPM")
		mData = oRange.getDataArray
		nCount = -1
		For nRow = 1 To UBound (mData) - 1
			nCount = nCount + 1
			ReDim Preserve maCPM (nCount) As New aCPM
			With maCPM (nCount)
				.fLevel = mData (nRow) (0)
				.fCPM = mData (nRow) (1)
			End With
		Next nRow
	End If
End  Sub

' subReadStarDust: Reads the star dust table.
Sub subReadStarDust
	Dim oSheet As Object, oRange As Object, mData As Variant
	Dim nCount As Integer, nRow As Integer
	
	If UBound (maStarDust) = -1 Then
		oSheet = ThisComponent.getSheets.getByName ("lvup")
		oRange = oSheet.getCellRangeByName ("A2:D81")
		mData = oRange.getDataArray
		nCount = -1
		For nRow = 1 To UBound (mData) - 1
			nCount = nCount + 1
			ReDim Preserve maStarDust (nCount) As New aStarDust
			With maStarDust (nCount)
				.fLevel = mData (nRow) (0)
				.nStarDust = mData (nRow) (2)
			End With
		Next nRow
	End If
End  Sub
