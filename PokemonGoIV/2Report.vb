' Copyright (c) 2017 imacat.
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

' 2Report: The Pokémon GO IV report generator.
'   by imacat <imacat@mail.imacat.idv.tw>, 2017-06-07

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

' subCreateReport: Creates the Pokémon GO IV report.
Sub subCreateReport ( _
		aBaseStats As aStats, aQuery As aFindIVParam, maIVs () As aIV)
	Dim oDoc As Object, oSheet As Object
	Dim oRange As Object, oColumns As Object, oRows As Object
	Dim oCell As Object, sFormula As String
	Dim nI As Integer, nJ As Integer, nCol As Integer
	Dim nLeadCols As Integer, nTotalCols As Integer
	Dim nEvolved As Integer, fMaxLevel As Double
	Dim sCPM As String, sMaxCPM As String
	Dim sColIVAttack As String, sColIVDefense As String
	Dim sColIVStamina As String
	Dim sPokemonName As String
	Dim mData (Ubound (maIVs) + 1) As Variant, mRow () As Variant
	Dim maEvBaseStats () As Variant
	Dim mProps () As New com.sun.star.beans.PropertyValue
	
	oDoc = StarDesktop.loadComponentFromURL ( _
		"private:factory/scalc", "_default", 0, mProps)
	oSheet = oDoc.getSheets.getByIndex (0)
	
	nEvolved = UBound (aBaseStats.mEvolved) + 1
	If nEvolved > 0 Then
		ReDim maEvBaseStats (nEvolved - 1) As Variant
		For nJ = 0 To nEvolved - 1
			maEvBaseStats (nJ) = fnGetBaseStats (aBaseStats.mEvolved (nJ))
		Next nJ
	End If
	
	If aQuery.nPlayerLevel <> 0 Then
		fMaxLevel = aQuery.nPlayerLevel + 1.5
		If fMaxLevel > 40 Then
			fMaxLevel = 40
		End If
	Else
		fMaxLevel = 40
	End If
	sMaxCPM = fnGetCPMFormula (fMaxLevel)
	
	' Sorts the IVs
	subSortIVs (aBaseStats, maEvBaseStats, maIVs, fMaxLevel)
	
	' Fills in the report data.
	mRow = Array ( _
		fnGetResString ("ReportNo"), _
		fnGetResString ("ReportPokemon"), _
		fnGetResString ("ReportCP"), _
		fnGetResString ("ReportHP"), _
		fnGetResString ("ReportStardust"), _
		fnGetResString ("ReportLevel"), _
		fnGetResString ("ReportAttack"), _
		fnGetResString ("ReportDefense"), _
		fnGetResString ("ReportStamina"), _
		fnGetResString ("ReportIVPercent"))
	nLeadCols = UBound (mRow) + 1
	
	nTotalCols = nLeadCols
	If aBaseStats.bIsLastForm Then
		nTotalCols = nTotalCols + 1
	End If
	For nJ = 0 To nEvolved - 1
		nTotalCols = nTotalCols + 1
		If maEvBaseStats (nJ).bIsLastForm Then
			nTotalCols = nTotalCols + 1
		End If
	Next nJ
	
	ReDim Preserve mRow (nTotalCols - 1) As Variant
	nCol = nLeadCols
	If aBaseStats.bIsLastForm Then
		mRow (nCol) = fnReplace (fnGetResString ("ReportCPPowerUp"), _
			"[Level]", fMaxLevel)
		nCol = nCol + 1
	End If
	For nJ = 0 To nEvolved - 1
		sPokemonName = fnGetResString ( _
			"Pokemon" & aBaseStats.mEvolved (nJ))
		mRow (nCol) = fnReplace (fnGetResString ("ReportCPEvolve"), _
			"[Pokémon]", sPokemonName)
		nCol = nCol + 1
		If maEvBaseStats (nJ).bIsLastForm Then
			mRow (nCol) = fnReplace (fnReplace ( _
				fnGetResString ("ReportCPEvolvePowerUp"), _
				"[Pokémon]", sPokemonName), _
				"[Level]", fMaxLevel)
		End If
	Next nJ
	mData (0) = mRow
	
	For nI = 0 To UBound (maIVs)
		mRow = Array ( _
			"", "", "", "", "", _
			maIVs (nI).fLevel, maIVs (nI).nAttack, _
			maIVs (nI).nDefense, maIVs (nI).nStamina, "")
		ReDim Preserve mRow (nTotalCols - 1) As Variant
		For nJ = nLeadCols To nEvolved - 1
			mRow (nJ) = ""
		Next nJ
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
	
	' Fills in the CP calculation.
	For nI = 0 To UBound (maIVs)
		sCPM = fnGetCPMFormula (maIVs (nI).fLevel)
		sColIVAttack = "G" & (nI + 2)
		sColIVDefense = "H" & (nI + 2)
		sColIVStamina = "I" & (nI + 2)
		
		oCell = oSheet.getCellByPosition (nLeadCols - 1, nI + 1)
		sFormula = "=(" & sColIVAttack & "+" & sColIVDefense _
			& "+" & sColIVStamina & ")/45"
		oCell.setFormula (sFormula)
		
		nCol = nLeadCols
		If aBaseStats.bIsLastForm Then
			oCell = oSheet.getCellByPosition (nCol, nI + 1)
			sFormula = fnGetCPFormula (aBaseStats, _
				sColIVAttack, sColIVDefense, sColIVStamina, sMaxCPM)
			oCell.setFormula (sFormula)
			nCol = nCol + 1
		End If
		For nJ = 0 To nEvolved - 1
			oCell = oSheet.getCellByPosition (nCol, nI + 1)
			sFormula = fnGetCPFormula (maEvBaseStats (nJ), _
				sColIVAttack, sColIVDefense, sColIVStamina, sCPM)
			oCell.setFormula (sFormula)
			nCol = nCol + 1
			If maEvBaseStats (nJ).bIsLastForm Then
				oCell = oSheet.getCellByPosition (nCol, nI + 1)
				sFormula = fnGetCPFormula (maEvBaseStats (nJ), _
					sColIVAttack, sColIVDefense, sColIVStamina, sMaxCPM)
				oCell.setFormula (sFormula)
				nCol = nCol + 1
			End If
		Next nJ
	Next nI
	
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
	
	oRange = oSheet.getCellRangeByPosition ( _
		nLeadCols, 0, nTotalCols - 1, 0)
	oRange.setPropertyValue ("IsTextWrapped", True)
	
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
	For nJ = nLeadCols To nTotalCols - 1
		oColumns.getByIndex (nJ).setPropertyValue ( _
			"Width", 2810)
	Next nJ
	oRows = oSheet.getRows
	oRows.getByIndex (0).setPropertyValue ("OptimalHeight", True)
End Sub

' subSortIVs: Sorts the IVs
Sub subSortIVs ( _
		aBaseStats As aStats, maEvBaseStats () As aIV, _
		maIVs () As aIV, fMaxLevel As Double)
	Dim nI As Integer, nJ As Integer
	Dim nCP As Integer
	
	' Calculate the sorting keys.
	For nI = 0 To UBound (maIVs) - 1
		maIVs (nI).nTotal = maIVs (nI).nAttack + maIVs (nI).nDefense _
			+ maIVs (nI).nStamina
		maIVs (nI).nMaxCP = fnCalcCP (aBaseStats, _
			maIVs (nI).fLevel, maIVs (nI).nAttack, _
			maIVs (nI).nDefense, maIVs (nI).nStamina)
		maIVs (nI).nMaxMaxCP = fnCalcCP (aBaseStats, _
			fMaxLevel, maIVs (nI).nAttack, _
			maIVs (nI).nDefense, maIVs (nI).nStamina)
		For nJ = 0 To UBound (aBaseStats.mEvolved)
			nCP = fnCalcCP (maEvBaseStats (nJ), _
				maIVs (nI).fLevel, maIVs (nI).nAttack, _
				maIVs (nI).nDefense, maIVs (nI).nStamina)
			If maIVs (nI).nMaxCP < nCP Then
				maIVs (nI).nMaxCP = nCP
			End If
			nCP = fnCalcCP (maEvBaseStats (nJ), _
				fMaxLevel, maIVs (nI).nAttack, _
				maIVs (nI).nDefense, maIVs (nI).nStamina)
			If maIVs (nI).nMaxMaxCP < nCP Then
				maIVs (nI).nMaxMaxCP = nCP
			End If
		Next nJ
	Next nI
	' Sort the IVs.
	For nI = 0 To UBound (maIVs) - 1
		For nJ = nI + 1 To UBound (maIVs)
			If fnCompareIV (maIVs (nI), maIVs (nJ)) > 0 Then
				' This is an array of data.  The data are actually
				' allocated in sequences.  maIVs (nI) is not a
				' reference.  They cannot simply be assigned.
				subSwapIV (maIVs (nI), maIVs (nJ))
			End If
		Next nJ
	Next nI
End Sub

' fnCompareIV: Compare two IVs for sorting
Function fnCompareIV (aIVa As aIV, aIVb As aIV) As Double
	Dim nCPa As Integer, nCPb As Integer, nI As Integer
	
	fnCompareIV = aIVb.nMaxMaxCP - aIVa.nMaxMaxCP
	If fnCompareIV <> 0 Then
		Exit Function
	End If
	fnCompareIV = aIVb.nMaxCP - aIVa.nMaxCP
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

' subSwapIV: Swaps two IVs
Function subSwapIV (aIVa As aIV, aIVb As aIV) As Double
	Dim aTempIV As New aIV
	
	With aTempIV
		.nAttack = aIVa.nAttack
		.nDefense = aIVa.nDefense
		.nStamina = aIVa.nStamina
		.nTotal = aIVa.nTotal
		.nMaxCP = aIVa.nMaxCP
		.nMaxMaxCP = aIVa.nMaxMaxCP
	End With
	With aIVa
		.nAttack = aIVb.nAttack
		.nDefense = aIVb.nDefense
		.nStamina = aIVb.nStamina
		.nTotal = aIVb.nTotal
		.nMaxCP = aIVb.nMaxCP
		.nMaxMaxCP = aIVb.nMaxMaxCP
	End With
	With aIVb
		.nAttack = aTempIV.nAttack
		.nDefense = aTempIV.nDefense
		.nStamina = aTempIV.nStamina
		.nTotal = aTempIV.nTotal
		.nMaxCP = aTempIV.nMaxCP
		.nMaxMaxCP = aTempIV.nMaxMaxCP
	End With
End Function

' fnGetCPFormula: Obtains the CP formula
Function fnGetCPFormula ( _
		aBaseStats As aStats, sColIVAttack As String, _
		sColIVDefense As String, sColIVStamina As String, _
		sCPM As String) As String
	fnGetCPFormula = "=FLOOR(" _
		& "(" & aBaseStats.nAttack & "+" & sColIVAttack & ")" _
		& "*SQRT(" & aBaseStats.nDefense & "+" & sColIVDefense & ")" _
		& "*SQRT(" & aBaseStats.nStamina & "+" & sColIVStamina & ")" _
		& "*POWER(" & sCPM & ";2)/10;1)"
End Function

' fnGetCPMFormula: Obtains the CPM
Function fnGetCPMFormula (fLevel As Double) As String
	If fLevel = CInt (fLevel) Then
		fnGetCPMFormula = "" & mCPM (fLevel)
	Else
		fnGetCPMFormula = "SQRT((" _
			& "POWER(" & mCPM (fLevel - 0.5) & ";2)" _
			& "+POWER(" & mCPM (fLevel + 0.5) & ";2))/2)"
	End If
End Function
