VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Navigation
Private Sub NeedsToHomeBtn_Click()
Worksheets("Home").Activate
End Sub

'Prints out the list of needs
Private Sub PrintNeedsBtn_Click()
PrintNeedsSheet
End Sub
