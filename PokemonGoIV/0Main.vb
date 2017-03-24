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
	sPokemonId As String
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
	sPokemonId As String
	sPokemonName As String
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
	aBaseStats = fnGetBaseStats (aQuery.sPokemonId)
	maIVs = fnFindIV (aBaseStats, aQuery)
	If UBound (maIVs) = -1 Then
		MsgBox fnGetResString ("ErrorNotFound")
	Else
		subSaveIV (aBaseStats, aQuery, maIVs)
	End If
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
	Dim fStep As Double, nN As Integer, nMaxLevel As Integer
	
	If aQuery.sPokemonId = "" Then
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
	nMaxLevel = aQuery.nPlayerLevel + 1.5
	If nMaxLevel > 40 Then
	    nMaxLevel = 40
	End If
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
										aBaseStats, nMaxLevel, _
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
										maIV (nN).maEvolved (nI).nMaxCP _
										    = fnCalcCP ( _
										    maEvBaseStats (nI), _
										    nMaxLevel, nAttack, _
										    nDefense, nStamina)
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
	If aQuery.sPokemonId = "Eevee" Then
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
		If aQuery.sPokemonId = "Eevee" Then
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
	mData (1) (1) = aQuery.sPokemonName
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
	
	If aQuery.sPokemonId = "Eevee" Then
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
	If aQuery.sPokemonId = "Eevee" Then
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
Function fnGetBaseStats (sPokemonId As String) As aStats
	Dim nI As Integer
	
	subReadBaseStats
	For nI = 0 To UBound (maBaseStats)
		If maBaseStats (nI).sPokemonId = sPokemonId Then
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
				.sPokemonId = mData (nI) (0)
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
