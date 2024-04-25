VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MeasurementForm 
   Caption         =   "Add Measurement"
   ClientHeight    =   2700
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4470
   OleObjectBlob   =   "MeasurementForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MeasurementForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







<<<<<<< HEAD
=======

<<<<<<< HEAD
>>>>>>> parent of 694c62f (Starting Print stuff)
=======
>>>>>>> parent of 694c62f (Starting Print stuff)
'Handle form submit
Private Sub AddMeasButton_Click()

'Add user input to Master List
AddMeasurementFromForm

'Hide form
MeasurementForm.Hide

End Sub

'Enable button if text box has value
Private Sub ValidateMeasurement()
AddMeasButton.Enabled = NewMeasurementBox.value <> ""
End Sub

'When there's a change in the text box, run the validate sub
Private Sub NewMeasurementBox_Change()
ValidateMeasurement
End Sub

'Upon launch, add button starts disabled
Private Sub UserForm_Initialize()
AddMeasButton.Enabled = False
End Sub
