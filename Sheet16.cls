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
'Allows for case insensitive searches
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

'Same fuctionality as DailyRadio
'TODO: Farm the over arching logic out to a different method called when clicked
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


Private Sub LabelsBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
''fjfj
End Sub

'Navigation
'Brings you to Order Sheet
Private Sub NewOrderButton_Click()
Worksheets("Order").Activate
End Sub

'Sub to handle bringing up reports on each ship
Private Sub ShipsDrop_Change()

'Lots to initialize
Dim searchRange As Range, lastInSearchRange As Integer
Dim shipName As String, arrToSort() As OrderRecord
Dim orderArr() As OrderRecord, sortedArr() As OrderRecord

'Save the name of the ship that is selected from the drop down menu
shipName = ShipsDrop.Text

'Find that ship in the ship DB
lastInSearchRange = Worksheets("ShipDatabase").Range("A" & Rows.Count).End(xlUp).Row
Set searchRange = Worksheets("ShipDatabase").Range("A1:A" & lastInSearchRange)

'If ship is in the ship DB and the Ship name isn't blank
'Ship DB will have blank rows so technically shipName = "" exisits in the DB
If Application.CountIf(searchRange, shipName) > 0 And shipName <> "" Then
    'Retrieve that ship from the OrderDB
    orderArr = CreateRecordFromDB(shipName)
    'Create order sheet and check sheet reports
    CreateBothReports orderArr
    'Sort the order
    sortedArr = SortOrderRecord(orderArr)
    'Write out stickers to the Home Sheet
    MakeStickers sortedArr, shipName
End If
End Sub

'Used in conjuction with ShowFoundItems to create a dynamic filtering list of items in daily sheet
'Re-filters with every key stroke
Private Sub ItemSearchBar_Change()
ShowFoundItems ItemSearchBar.Text
End Sub

'Displays results of what is in the search bar
Sub ShowFoundItems(searchTerm As String)

'Initialize
'TODO: Work on variable names
Dim dailyRange As Range, i As Integer, last As Integer, lookUp As String, j As Integer

'Clear the results box on key stroke
ItemSearchResults.Clear

'Add wild cards before and after whatever gets typed into the search bar
lookUp = "*" & searchTerm & "*"

'Establish search area
last = Worksheets("Daily").Range("D" & Rows.Count).End(xlUp).Row
Set dailyRange = Worksheets("Daily").Range("A2:D" & last)

'List box row incrementor
j = 0

'Loop through the rows in daily
For i = 1 To last
    'If anything resembles the search term write it into the box
    If dailyRange.Cells(i, 3).value Like lookUp Then
        ItemSearchResults.AddItem
        ItemSearchResults.List(j, 0) = dailyRange.Cells(i, 1).value
        ItemSearchResults.List(j, 1) = dailyRange.Cells(i, 2).value
        ItemSearchResults.List(j, 2) = dailyRange.Cells(i, 3).value
        ItemSearchResults.List(j, 3) = dailyRange.Cells(i, 4).value
        j = j + 1
    End If
Next i

ItemSearchResults.Height = 144.75

End Sub


