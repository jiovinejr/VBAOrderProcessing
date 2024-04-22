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
Dim arr As Variant, check As String
arr = GetShipsFromDB("DailyDatabase")
If IsEmpty(arr) Then
    MsgBox "Add ships to daily"
    DeckRadio.value = True
ElseIf Application.WorksheetFunction.CountA(arr) = 1 Then
    check = CStr(arr)
    ShipsDrop.Clear
    ShipsDrop.value = check
    ShipsDrop.AddItem (check)
Else
    ShipsDrop.List = arr
    ShipsDrop.Text = ShipsDrop.List(1)
End If

End Sub

Private Sub DeckRadio_Click()
Dim arr As Variant
arr = GetShipsFromDB("ShipsOnDeck")
ShipsDrop.value = arr(1, 1)
ShipsDrop.List = arr
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
