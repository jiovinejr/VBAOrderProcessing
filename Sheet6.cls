VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'Navigation
Private Sub OrderToHomeBtn_Click()
Worksheets("Home").Activate
End Sub

'Pastes in ordered copied from Drive
Private Sub PasteBtn_Click()
PasteSpeacial
End Sub

'Processes the order pasted in
Private Sub ProcessBtn_Click()
shipName = Range("C1")
ProcessPastedOrder
Worksheets("Home").Activate
Worksheets("Home").ShipsDrop.value = shipName
End Sub
