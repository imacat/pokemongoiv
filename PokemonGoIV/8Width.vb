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

' 8Width: The dialog text width tester.
'   by imacat <imacat@mail.imacat.idv.tw>, 2017-02-22

Option Explicit

' subTestWidth: Tests the width of the dialog text.
Sub subTestWidth
	Dim oDialog As Object, oDialogModel As Object
	Dim oTextModel As Object, oListModel As Object
	Dim oEditModel As Object, oNumModel As Object
	Dim oButtonModel As Object
	Dim mItems () As String, oListener As Object
	
	' Creates a dialog
	oDialogModel = CreateUnoService ( _
		"com.sun.star.awt.UnoControlDialogModel")
	oDialogModel.setPropertyValue ("PositionX", 100)
	oDialogModel.setPropertyValue ("PositionY", 100)
	oDialogModel.setPropertyValue ("Height", 65)
	oDialogModel.setPropertyValue ("Width", 200)
	oDialogModel.setPropertyValue ("Title", _
	    "Localization Text Width Test")
	
	' Adds a text label.
	oTextModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlFixedTextModel")
	oTextModel.setPropertyValue ("PositionX", 5)
	oTextModel.setPropertyValue ("PositionY", 6)
	oTextModel.setPropertyValue ("Height", 8)
	oTextModel.setPropertyValue ("Width", 190)
	oTextModel.setPropertyValue ("BackgroundColor", RGB (0, 255, 0))
	oDialogModel.insertByName ("txtText", oTextModel)
	
	' Adds a drop down list.
	oListModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlListBoxModel")
	oListModel.setPropertyValue ("PositionX", 5)
	oListModel.setPropertyValue ("PositionY", 19)
	oListModel.setPropertyValue ("Height", 12)
	oListModel.setPropertyValue ("Width", 50)
	oListModel.setPropertyValue ("TabIndex", 0)
	oListModel.setPropertyValue ("Dropdown", True)
	oDialogModel.insertByName ("lstStat", oListModel)
		
	' Adds a text input field.
	oEditModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlEditModel")
	oEditModel.setPropertyValue ("PositionX", 5)
	oEditModel.setPropertyValue ("PositionY", 34)
	oEditModel.setPropertyValue ("Height", 12)
	oEditModel.setPropertyValue ("Width", 150)
	oEditModel.setPropertyValue ("TabIndex", 1)
	oDialogModel.insertByName ("edtText", oEditModel)
	
	' Adds a numeric input field.
	oNumModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlNumericFieldModel")
	oNumModel.setPropertyValue ("PositionX", 160)
	oNumModel.setPropertyValue ("PositionY", 34)
	oNumModel.setPropertyValue ("Height", 12)
	oNumModel.setPropertyValue ("Width", 35)
	oNumModel.setPropertyValue ("TabIndex", 2)
	oNumModel.setPropertyValue ("Value", _
		oTextModel.getPropertyValue ("Width"))
	oNumModel.setPropertyValue ("ValueMax", 190)
	oNumModel.setPropertyValue ("ValueMin", 1)
	oNumModel.setPropertyValue ("DecimalAccuracy", 0)
	oNumModel.setPropertyValue ("Spin", True)
	oDialogModel.insertByName ("numWidth", oNumModel)
	
	' Adds a button.
	oButtonModel = oDialogModel.createInstance ( _
		"com.sun.star.awt.UnoControlButtonModel")
	oButtonModel.setPropertyValue ("PositionX", 5)
	oButtonModel.setPropertyValue ("PositionY", 49)
	oButtonModel.setPropertyValue ("Height", 12)
	oButtonModel.setPropertyValue ("Width", 50)
	oButtonModel.setPropertyValue ("PushButtonType", _
		com.sun.star.awt.PushButtonType.OK)
	oButtonModel.setPropertyValue ("DefaultButton", True)
	oDialogModel.insertByName ("btnClose", oButtonModel)
	
	' Adds the dialog model to the control and runs it.
	oDialog = CreateUnoService ("com.sun.star.awt.UnoControlDialog")
	oDialog.setModel (oDialogModel)
	oDialog.setVisible (True)
	
	oListener = CreateUnoListener ("subTextChanged_", "com.sun.star.awt.XTextListener")
	oDialog.getControl ("edtText").addTextListener (oListener)
	oDialog.getControl ("edtText").setFocus
	
	oListener = CreateUnoListener ("subWidthChanged_", "com.sun.star.awt.XTextListener")
	oDialog.getControl ("numWidth").addTextListener (oListener)
	
	oDialog.execute
End Sub

' subTextChanged_disposing: When the text input box is disposed.
Sub subTextChanged_disposing (oEvent As object)
End Sub

' subTextChanged_textChanged: When the text is changed.
Sub subTextChanged_textChanged (oEvent As object)
	Dim oEdit As Object, oText As Object, oDropdown As Object
	
	oEdit = oEvent.Source
	oText = oEdit.getContext.getControl ("txtText")
	oText.setText (oEdit.getText)
	oDropdown = oEdit.getContext.getControl ("lstStat")
	oDropdown.removeItems (0, oDropdown.getItemCount)
	oDropdown.addItem (oEdit.getText, 0)
	oDropdown.selectItemPos (0, True)
End Sub

' subWidthChanged_disposing: When the width input box is disposed.
Sub subWidthChanged_disposing (oEvent As object)
End Sub

' subWidthChanged_textChanged: When the width is changed.
Sub subWidthChanged_textChanged (oEvent As object)
	Dim oEdit As Object, oText As Object, oDropdown As Object
	
	oEdit = oEvent.Source
	oText = oEdit.getContext.getControl ("txtText")
	oText.getModel.setPropertyValue ("Width", oEdit.getValue)
	oDropdown = oEdit.getContext.getControl ("lstStat")
	oDropdown.getModel.setPropertyValue ("Width", oEdit.getValue)
End Sub
