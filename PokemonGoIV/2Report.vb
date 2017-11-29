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

Sub subGetPokemonSheet
	Dim oDoc As Object
	Dim oEnum As Object, oComponent As Object, sTitles As String
	
	oDoc = fnCreateNewSpreadsheetDocument
	oDoc.setTitle ("Pokemon GO IV")
	oEnum = StarDesktop.getComponents.createEnumeration
	Do While oEnum.hasMoreElements
		oComponent = oEnum.nextElement
		If oComponent.supportsService ("com.sun.star.sheet.SpreadsheetDocument") Then
			If oComponent.getTitle = "Pokemon GO IV" Then
				Xray oComponent
			End If
		End If
	Loop
End Sub

' subCreateReport: Creates the Pokémon GO IV report.
Sub subCreateReport ( _
		aBaseStats As aStats, aQuery As aFindIVParam, maIVs () As aIV)
	Dim oDoc As Object, oSheet As Object
	Dim oRange As Object, oColumns As Object, oRows As Object
	Dim oCell As Object, sFormula As String, sFormulaLocal As String
	Dim nI As Integer, nJ As Integer, nCol As Integer
	Dim nLeadCols As Integer, nTotalCols As Integer
	Dim nEvolved As Integer, fMaxLevel As Double
	Dim sCPM As String, sMaxCPM As String
	Dim sColIVAttack As String, sColIVDefense As String
	Dim sColIVStamina As String
	Dim sPokemonName As String
	Dim mLeadHead () As Variant, nStartRow As Integer
	Dim mData (0) As Variant, mRow () As Variant
	Dim maEvBaseStats () As Variant
	Dim mProps () As New com.sun.star.beans.PropertyValue
	
	oSheet = fnFindPokemonGOIVSheet (aQuery.sPokemonName)
	
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
	
	' Gathers the header row.
	mLeadHead = Array ( _
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
	nLeadCols = UBound (mLeadHead) + 1
	
	' Calculating how many columns do we need to fill in the
	' CP of the evolved forms.
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
	
	' Adds the header row if this is a new spreadsheet
	oCell = oSheet.getCellByPosition (0, 0)
	If oCell.getString = "" Then
		' The leading columns of the header row
		mRow = mLeadHead
		' Fill in the header row with the CP of the evolved forms.
		ReDim Preserve mRow (nTotalCols - 1) As Variant
		nCol = nLeadCols
		If aBaseStats.bIsLastForm Then
			mRow (nCol) = fnReplace ( _
				fnGetResString ("ReportCPPowerUp"), _
				"[Level]", fMaxLevel)
			nCol = nCol + 1
		End If
		For nJ = 0 To nEvolved - 1
			sPokemonName = fnGetResString ( _
				"Pokemon" & aBaseStats.mEvolved (nJ))
			mRow (nCol) = fnReplace ( _
				fnGetResString ("ReportCPEvolve"), _
				"[Pokémon]", sPokemonName)
			nCol = nCol + 1
			If maEvBaseStats (nJ).bIsLastForm Then
				mRow (nCol) = fnReplace (fnReplace ( _
					fnGetResString ("ReportCPEvolvePowerUp"), _
					"[Pokémon]", sPokemonName), _
					"[Level]", fMaxLevel)
				nCol = nCol + 1
			End If
		Next nJ
	
		' Fills in the header row
		ReDim mData (0) As Variant
		mData (0) = mRow
		oRange = oSheet.getCellRangeByPosition ( _
			0, 0, UBound (mData (0)), UBound (mData))
		oRange.setDataArray (mData)
		oRange.setPropertyValue ("VertJustify", _
			com.sun.star.table.CellVertJustify.TOP)
		oRange = oSheet.getCellRangeByPosition ( _
			nLeadCols, 0, nTotalCols - 1, 0)
		oRange.setPropertyValue ("IsTextWrapped", True)
		
		' Sets the height of the header row
		oRows = oSheet.getRows
		oRows.getByIndex (0).setPropertyValue ("OptimalHeight", True)
		
		' Sets the widths of the columns
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
				"Width", 2500)
		Next nJ
		
		nStartRow = 1
	
	' Append to the end on an existing spreadsheet
	Else
		nStartRow = 0
		Do
			nStartRow = nStartRow + 1
			oCell = oSheet.getCellByPosition (5, nStartRow)
		Loop While oCell.getString <> ""
	End If
	
	' Gathers the data rows.
	ReDim mData (Ubound (maIVs)) As Variant
	For nI = 0 To UBound (maIVs)
		mRow = Array ( _
			"", "", "", "", "", _
			maIVs (nI).fLevel, maIVs (nI).nAttack, _
			maIVs (nI).nDefense, maIVs (nI).nStamina, "")
		ReDim Preserve mRow (nTotalCols - 1) As Variant
		For nJ = nLeadCols To nEvolved - 1
			mRow (nJ) = ""
		Next nJ
		mData (nI) = mRow
	Next nI
	
	' Fills the query information at the first row
	mData (0) (0) = aBaseStats.sNo
	mData (0) (1) = aQuery.sPokemonName
	mData (0) (2) = aQuery.nCP
	mData (0) (3) = aQuery.nHP
	mData (0) (4) = aQuery.nStardust
	
	oRange = oSheet.getCellRangeByPosition ( _
		0, nStartRow, _
		UBound (mData (0)), nStartRow + UBound (mData))
	oRange.setDataArray (mData)
	oRange.setPropertyValue ("VertJustify", _
		com.sun.star.table.CellVertJustify.TOP)
	
	' Fills in the CP calculation.
	For nI = 0 To UBound (maIVs)
		sCPM = fnGetCPMFormula (maIVs (nI).fLevel)
		sColIVAttack = "G" & (nStartRow + nI + 1)
		sColIVDefense = "H" & (nStartRow + nI + 1)
		sColIVStamina = "I" & (nStartRow + nI + 1)
		
		oCell = oSheet.getCellByPosition (nLeadCols - 1, nStartRow + nI)
		sFormula = "=(" & sColIVAttack & "+" & sColIVDefense _
			& "+" & sColIVStamina & ")/45"
		oCell.setFormula (sFormula)
		sFormulaLocal = oCell.getPropertyValue ("FormulaLocal")
		If sFormulaLocal <> sFormula Then
			oCell.setPropertyValue ("FormulaLocal", sFormulaLocal)
		End If
		
		nCol = nLeadCols
		If aBaseStats.bIsLastForm Then
			oCell = oSheet.getCellByPosition (nCol, nStartRow + nI)
			sFormula = fnGetCPFormula (aBaseStats, _
				sColIVAttack, sColIVDefense, sColIVStamina, sMaxCPM)
			oCell.setFormula (sFormula)
			sFormulaLocal = oCell.getPropertyValue ("FormulaLocal")
			If sFormulaLocal <> sFormula Then
				oCell.setPropertyValue ("FormulaLocal", sFormulaLocal)
			End If
			nCol = nCol + 1
		End If
		For nJ = 0 To nEvolved - 1
			oCell = oSheet.getCellByPosition (nCol, nStartRow + nI)
			sFormula = fnGetCPFormula (maEvBaseStats (nJ), _
				sColIVAttack, sColIVDefense, sColIVStamina, sCPM)
			oCell.setFormula (sFormula)
			sFormulaLocal = oCell.getPropertyValue ("FormulaLocal")
			If sFormulaLocal <> sFormula Then
				oCell.setPropertyValue ("FormulaLocal", sFormulaLocal)
			End If
			nCol = nCol + 1
			If maEvBaseStats (nJ).bIsLastForm Then
				oCell = oSheet.getCellByPosition (nCol, nStartRow + nI)
				sFormula = fnGetCPFormula (maEvBaseStats (nJ), _
					sColIVAttack, sColIVDefense, _
					sColIVStamina, sMaxCPM)
				oCell.setFormula (sFormula)
				sFormulaLocal = oCell.getPropertyValue ( _
					"FormulaLocal")
				If sFormulaLocal <> sFormula Then
					oCell.setPropertyValue ( _
						"FormulaLocal", sFormulaLocal)
				End If
				nCol = nCol + 1
			End If
		Next nJ
	Next nI
	
	' Merge the lead cells.
	oRange = oSheet.getCellRangeByPosition ( _
		0, nStartRow, 0, nStartRow + UBound (mData))
	oRange.merge (True)
	oRange = oSheet.getCellRangeByPosition ( _
		1, nStartRow, 1, nStartRow + UBound (mData))
	oRange.merge (True)
	oRange = oSheet.getCellRangeByPosition ( _
		2, nStartRow, 2, nStartRow + UBound (mData))
	oRange.merge (True)
	oRange = oSheet.getCellRangeByPosition ( _
		3, nStartRow, 3, nStartRow + UBound (mData))
	oRange.merge (True)
	oRange = oSheet.getCellRangeByPosition ( _
		4, nStartRow, 4, nStartRow + UBound (mData))
	oRange.merge (True)
	oRange = oSheet.getCellRangeByPosition ( _
		9, nStartRow, 9, nStartRow + UBound (mData))
	oRange.setPropertyValue ("NumberFormat", 10)
End Sub

' subSortIVs: Sorts the IVs
Sub subSortIVs ( _
		aBaseStats As aStats, maEvBaseStats () As aIV, _
		maIVs () As aIV, fMaxLevel As Double)
	Dim nI As Integer, nJ As Integer
	Dim nCP As Integer
	
	' Calculate the sorting keys.
	For nI = 0 To UBound (maIVs)
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
		.fLevel = aIVa.fLevel
		.nAttack = aIVa.nAttack
		.nDefense = aIVa.nDefense
		.nStamina = aIVa.nStamina
		.nTotal = aIVa.nTotal
		.nMaxCP = aIVa.nMaxCP
		.nMaxMaxCP = aIVa.nMaxMaxCP
	End With
	With aIVa
		.fLevel = aIVb.fLevel
		.nAttack = aIVb.nAttack
		.nDefense = aIVb.nDefense
		.nStamina = aIVb.nStamina
		.nTotal = aIVb.nTotal
		.nMaxCP = aIVb.nMaxCP
		.nMaxMaxCP = aIVb.nMaxMaxCP
	End With
	With aIVb
		.fLevel = aTempIV.fLevel
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

' fnFindPokemonGOIVSheet: Finds the existing sheet for the result.
Function fnFindPokemonGOIVSheet (sPokemon As String) As Object
	Dim oDoc As Object, sDocTitle As String
	Dim oSheets As Object, nCount As Integer, oSheet As Object
	Dim mNames () As String, nI As Integer
	Dim mProps () As New com.sun.star.beans.PropertyValue
	
	sDocTitle = "Pokémon GO IV"
	oDoc = fnFindDocByTitle (sDocTitle)
	If IsNull (oDoc) Then
		oDoc = StarDesktop.loadComponentFromURL ( _
			"private:factory/scalc", "_default", 0, mProps)
		oDoc.getDocumentProperties.Title = sDocTitle
		oSheets = oDoc.getSheets
		mNames = oSheets.getElementNames
		oSheets.insertNewByName (sPokemon, 0)
		oSheet = oSheets.getByName (sPokemon)
		For nI = 0 To UBound (mNames)
			oSheets.removeByName (mNames (nI))
		Next nI
	Else
		oSheet = fnFindSheetByName (oDoc, sPokemon)
		If IsNull (oSheet) Then
			oSheets = oDoc.getSheets
			nCount = oSheets.getCount
			oSheets.insertNewByName (sPokemon, nCount)
			oSheet = oSheets.getByName (sPokemon)
		End If
		oDoc.getCurrentController.setActiveSheet (oSheet)
	End If
	fnFindPokemonGOIVSheet = oSheet
End Function

' fnFindDocByTitle: Finds the document by its title.
Function fnFindDocByTitle (sTitle) As Object
	Dim oEnum As Object, oDoc As Object
	
	oEnum = StarDesktop.getComponents.createEnumeration
	Do While oEnum.hasMoreElements
		oDoc = oEnum.nextElement
		If oDoc.supportsService ( _
				"com.sun.star.sheet.SpreadsheetDocument") Then
			If oDoc.getDocumentProperties.Title = sTitle Then
				fnFindDocByTitle = oDoc
				Exit Function
			End If
		End If
	Loop
End Function

' fnFindSheetByName: Finds the spreadsheet by its name
Function fnFindSheetByName (oDoc As Object, sName As String) As Object
	Dim oSheets As Object, mNames () As String, nI As Integer
	
	oSheets = oDoc.getSheets
	mNames = oSheets.getElementNames
	For nI = 0 To UBound (mNames)
		If mNames (nI) = sName Then
			fnFindSheetByName = oSheets.getByIndex (nI)
			Exit Function
		End If
	Next nI
End Function
