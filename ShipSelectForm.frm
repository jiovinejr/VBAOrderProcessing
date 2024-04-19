VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ShipSelectForm 
   Caption         =   "UserForm1"
   ClientHeight    =   9765.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9180.001
   OleObjectBlob   =   "ShipSelectForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ShipSelectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Ok_Click()

Dim answer As Integer

answer = MsgBox("Is today a new day?", vbYesNo, "Daily Needs")

If answer = vbYes Then
    ClearDailyDB
    AddShipsToDB
Else
    AddShipsToDB
End If

ShipSelectForm.Hide

End Sub

Private Sub ShipsOnDeck_Change()
Ok.Enabled = True
End Sub

Private Sub UserForm_Initialize()
Ok.Enabled = False
End Sub



