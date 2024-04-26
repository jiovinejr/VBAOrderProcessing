VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Code involving form controls that are imbedded in the Home Sheet
Option Compare Text

'Clicking the "Daily List" button uses method in "UserFormLogic" and other code in the form itself to trigger
Private Sub DailyListBtn_Click()

DisplayShipSelectForm

End Sub

'Checking "Daily" radio button
Private Sub DailyRadio_Click()

'Initialize
Dim arr As Variant, check As String

'Store ships from daily database sheet to a variable
arr = GetShipsFromDB("DailyDatabase")

'If the variable is empty
If IsEmpty(arr) Then
    'Alert the user
    MsgBox "Add ships to daily"
    'Uncheck the radio button
    DailyRadio.value = False
    
'If theres only one item in the sheet the variable will be a string
ElseIf Application.WorksheetFunction.CountA(arr) = 1 Then
    'Cast to a string JIC
    check = CStr(arr)
    'Clear out data in the dropdown menu
    ShipsDrop.Clear
    'Display the string in the window
    ShipsDrop.value = check
    'Add it to the dropdown list
    ShipsDrop.AddItem (check)

'If the variable has more than one item treat it as an array
Else
    'Set entire array in the dropdown box
    ShipsDrop.List = arr
    'Display the first item in the window
    ShipsDrop.Text = ShipsDrop.List(1)
End If

End Sub

Private Sub DeckRadio_Click()
Dim arr As Variant
arr = GetShipsFromDB("ShipsOnDeck")
If IsEmpty(arr) Then
    MsgBox "Deck Clear"
    DeckRadio.value = False
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


Private Sub NewOrderButton_Click()
Worksheets("Order").Activate
End Sub

Private Sub ShipsDrop_Change()
Dim searchRange As Range, lastInSearchRange As Integer
Dim shipName As String, arrToSort() As OrderRecord
Dim orderArr() As OrderRecord, sortedArr() As OrderRecord
shipName = ShipsDrop.Text
lastInSearchRange = Worksheets("ShipDatabase").Range("A" & Rows.Count).End(xlUp).Row
Set searchRange = Worksheets("ShipDatabase").Range("A1:A" & lastInSearchRange)
If Application.CountIf(searchRange, shipName) > 0 And shipName <> "" Then
    arrToSort = CreateRecordFromDB(shipName)
    orderArr = CreateRecordFromDB(shipName)
    sortedArr = SortOrderRecord(arrToSort)
    CreateBothReports orderArr
    MakeStickers sortedArr, shipName
End If
End Sub

Private Sub ItemSearchBar_Change()
If IsEmpty(ItemSearchBar) Then
    ItemSearchResults.Clear
End If

ShowFoundItems ItemSearchBar.Text
End Sub

Sub ShowFoundItems(searchTerm As String)

Dim dailyRange As Range, i As Integer, last As Integer, lookUp As String, j As Integer

ItemSearchResults.Clear

lookUp = "*" & searchTerm & "*"

If lookUp = "**" Then
    ItemSearchResults.Clear
End If

last = Worksheets("Daily").Range("D" & Rows.Count).End(xlUp).Row
Set dailyRange = Worksheets("Daily").Range("A2:D" & last)
j = 0
For i = 1 To last
    If dailyRange.Cells(i, 3).value Like lookUp Then
        ItemSearchResults.AddItem
        ItemSearchResults.List(j, 0) = dailyRange.Cells(i, 1).value
        ItemSearchResults.List(j, 1) = dailyRange.Cells(i, 2).value
        ItemSearchResults.List(j, 2) = dailyRange.Cells(i, 3).value
        ItemSearchResults.List(j, 3) = dailyRange.Cells(i, 4).value
        j = j + 1
    End If
Next i

End Sub
