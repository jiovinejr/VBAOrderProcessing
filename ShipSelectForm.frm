VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ShipSelectForm 
   Caption         =   "Select Ships"
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


'Ship Selection for daily list
'Handle when someone clicks the ok button after selection
Private Sub Ok_Click()

'Initialize
Dim answer As Integer

'Save the user selection to a variable
answer = MsgBox("Is today a new day?", vbYesNo, "Daily Needs")

'If yes, it is in fact a new day
If answer = vbYes Then
    'Clear out any data still in the daily sheet
    ClearDailyDB
    'Add ships from on deck to the daily
    AddShipsToDB
Else
    'If its not a new day, just add ship from deck to daily
    AddShipsToDB
End If

'Hide the form since the user is finished using
ShipSelectForm.Hide

'Write data to respective sheets from user selection
WriteLists

'Tally up items needed
CreateNeedsSheet
End Sub

'Upon opening, the ok button is disabled
Private Sub UserForm_Initialize()
Ok.Enabled = False
End Sub

'When user selects at least 1 item, the ok button activates
Private Sub ShipsOnDeckBox_Change()
Ok.Enabled = True
End Sub


