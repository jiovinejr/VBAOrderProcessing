VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub DailyListBtn_Click()

DisplayShipSelectForm

End Sub

Private Sub DailyRadio_Click()
Dim lastRow As Integer
lastRow = Worksheets("DailyDatabase").Range("A" & Rows.Count).End(xlUp).Row
ShipsDrop.ListFillRange = "DailyDatabase!A2:A" & lastRow
End Sub

Private Sub DeckRadio_Click()
Dim lastRow As Integer
lastRow = Worksheets("ShipsOnDeck").Range("A" & Rows.Count).End(xlUp).Row
ShipsDrop.ListFillRange = "ShipsOnDeck!A1:A" & lastRow
End Sub


Private Sub ShipsDrop_Change()
Dim searchRange As Range, lastInSearchRange As Integer
Dim shipName As String, arrToSort() As OrderRecord
Dim orderArr() As OrderRecord, sortedArr() As OrderRecord
shipName = ShipsDrop.Text
lastInSearchRange = Worksheets("ShipDatabase").Range("A" & Rows.Count).End(xlUp).Row
Set searchRange = Worksheets("ShipDatabase").Range("A1:A" & lastInSearchRange)
If Application.CountIf(searchRange, shipName) > 0 Then
    arrToSort = CreateRecordFromDB(shipName)
    orderArr = CreateRecordFromDB(shipName)
    sortedArr = SortOrderRecord(arrToSort)
    CreateBothReports orderArr
    MakeStickers sortedArr, shipName
End If
End Sub
