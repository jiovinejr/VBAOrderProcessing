VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddToMaster 
   Caption         =   "Item Name Error"
   ClientHeight    =   3480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7950
   OleObjectBlob   =   "AddToMaster.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddToMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AddBtn_Click()
AddDataFromForm
AddToMaster.Hide
End Sub
Private Sub UserForm_Initialize()
CategoryBox.AddItem "Fruits"
CategoryBox.AddItem "Vegetables"
End Sub


