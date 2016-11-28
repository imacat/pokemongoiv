' _1Dialog: The UI of the Pokémon IV calculator
'   by imacat <imacat@mail.imacat.idv.tw>, 2016-11-27

Option Explicit

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
	oDialogModel.setPropertyValue ("Title", "Pokémon Parameters")
	
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
	oTextModel.setPropertyValue ("Label", "~Star dust:")
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
		"Amazed me/wonder/best", _
		"Strong/caught my attention", _
		"Decent/above average", _
		"Not great/not make headway/has room")
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
		"WOW/incredible/stats are best", _
		"Excellent/impressed/impressive", _
		"Get the job done/noticeable/some good stats", _
		"No greatness/not out of the norm/kinda basic")
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
	oDialog.execute
	
	With aQuery
		.sPokemon = oDialog.getControl ("lstPokemon").getSelectedItem
		.nCP = oDialog.getControl ("numCP").getValue
		.nHP = oDialog.getControl ("numHP").getValue
		.nStarDust = CInt (oDialog.getControl ("lstStarDust").getSelectedItem)
		.nPlayerLevel = CInt (oDialog.getControl ("lstPlayerLevel").getSelectedItem)
		.nAppraisal1 = oDialog.getControl ("lstApprasal1").getSelectedItemPos + 1
		.nAppraisal2 = oDialog.getControl ("lstApprasal2").getSelectedItemPos + 1
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

sub subBtnOK_actionPerformed
	MsgBox "OK"
End Sub

sub subBtnOK_disposing
	MsgBox "OK"
End Sub
