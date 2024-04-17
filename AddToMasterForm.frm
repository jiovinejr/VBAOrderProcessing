VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddToMasterForm 
   Caption         =   "Item Name Error"
   ClientHeight    =   3480
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

'Takes care of category drop down upon creation
Private Sub UserForm_Initialize()
CategoryBox.AddItem "Fruits"
CategoryBox.AddItem "Vegetables"
End Sub


