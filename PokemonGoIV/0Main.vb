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
	bIsLastForm As Boolean
End Type

' The individual values of a Pokémon.
Type aIV
	fLevel As Double
	nStamina As Integer
	nAttack As Integer
	nDefense As Integer
	' For sorting
	nTotal As Integer
	nMaxCP As Integer
	nMaxMaxCP As Integer
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
		subCreateReport (aBaseStats, aQuery, maIVs)
	End If
End Sub

' fnFindIV: Finds the possible individual values of the Pokémon
Function fnFindIV ( _
		aBaseStats As aStats, aQuery As aFindIVParam) As Variant
	Dim maIV () As New aIV, nN As Integer
	Dim fLevel As Double, nStamina As Integer
	Dim nAttack As Integer, nDefense As integer
	Dim nI As Integer, nJ As Integer, fLevelStep As Double
	
	If aQuery.sPokemonId = "" Then
		fnFindIV = maIV
		Exit Function
	End If
	If aQuery.bIsNew Then
		fLevelStep = 1
	Else
		fLevelStep = 0.5
	End If
	subReadStardust
	nN = -1
	For fLevel = 1 To UBound (mStardust) Step fLevelStep
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
								End With
							End If
						Next nDefense
					Next nAttack
				End If
			Next nStamina
		End If
	Next fLevel
	fnFindIV = maIV
End Function

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
		fnGetCPM = ((mCPM (fLevel - 0.5) ^ 2 _
			+ mCPM (fLevel + 0.5) ^ 2) / 2) ^ 0.5
	End If
End Function

' fnFloor: Returns the floor of the number
Function fnFloor (fNumber As Double) As Integer
	fnFloor = CInt (fNumber - 0.5)
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

' subReadBaseStats: Reads the base stats table.
Sub subReadBaseStats
	Dim mData As Variant, nI As Integer, nJ As Integer, nK As Integer
	Dim nEvolved As Integer, mEvolved () As Variant
	
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
			nEvolved = UBound (mData (nI) (5)) + 1
			mEvolved = Array ()
			maBaseStats (nI).bIsLastForm = True
			If nEvolved > 0 Then
				ReDim mEvolved (nEvolved - 1) As Variant
				For nJ = 0 To nEvolved - 1
					mEvolved (nJ) = mData (nI) (5) (nJ)
				Next nJ
				maBaseStats (nI).mEvolved = mEvolved
				maBaseStats (nI).bIsLastForm = False
			End If
		Next nI
	End If
End Sub

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
