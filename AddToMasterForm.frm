VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddToMasterForm 
   Caption         =   "Item Name Error"
   ClientHeight    =   3720
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7950
   OleObjectBlob   =   "AddToMasterForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddToMasterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







'Handle form button click
Private Sub AddBtn_Click()

'Use abstracted method
AddDataFromForm

'Hide the form on click
AddToMasterForm.Hide

End Sub

'Data validate on change
Private Sub CaseWeightBox_Change()
Validate
End Sub

'Data validate on change
Private Sub NewNameBox_Change()
Validate
End Sub

'Validation logic
Private Sub Validate()

'Enable button if new name is not empty and case weight is not empty as well as a valid number
AddBtn.Enabled = NewNameBox.value <> "" And CaseWeightBox.value <> "" And IsNumeric(CaseWeightBox.value)

'If case weight scenario is false show user error hint
If CaseWeightBox.value <> "" And Not IsNumeric(CaseWeightBox.value) Then
    NumberErrorLabel.Visible = True
Else
    NumberErrorLabel.Visible = False
End If

'If the criteria for button are met change the button caption
If AddBtn.Enabled = True Then
    AddBtn.Caption = "Add"
Else
    AddBtn.Caption = "Fill out all fields"
End If
End Sub


'Takes care of category drop down upon creation
Private Sub UserForm_Initialize()
AddBtn.Enabled = False
AddBtn.Caption = "Fill out all fields"
NumberErrorLabel.Visible = False
CategoryBox.AddItem "Fruits"
CategoryBox.AddItem "Vegetables"
End Sub


