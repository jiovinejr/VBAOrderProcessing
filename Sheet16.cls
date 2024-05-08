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

'Navigation
Private Sub HomeToIODBtn_Click()
Worksheets("Items on Deck").Activate
End Sub

'Navigation
Private Sub HomeToNeedsBtn_Click()
Worksheets("Needs").Activate
End Sub

'Double clicking an item in the labels box where all the stickers are displayed
'Prints an individual label by using the index in the list where the item is located
Private Sub LabelsBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    PrintBoxLabels LabelsBox.ListIndex, LabelsBox.ListIndex
End Sub

'Triggers when a radio button is clicked
'Pass in DB where you want your list from and the button itself
Sub HandleRadioClick(dbName As String, radio As Object)
'Initialize
Dim arr As Variant, check As String

'Store ships from daily database sheet to a variable
arr = GetShipsFromDB(dbName)

'If the variable is empty
If IsEmpty(arr) Then
    'Alert the user
    MsgBox "No ships in this list"
    'Uncheck the radio button
    radio.value = False

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
    ShipsDrop.list = arr
    'Display the first item in the window
    ShipsDrop.Text = ShipsDrop.list(1)
End If

End Sub

'Checking "Daily" radio button
Private Sub DailyRadio_Click()
MoveBackToDeckBtn.Visible = True
HandleRadioClick "DailyDatabase", DailyRadio
End Sub

'Checking "Deck" radio button
Private Sub DeckRadio_Click()
MoveBackToDeckBtn.Visible = False
HandleRadioClick "ShipsOnDeck", DeckRadio
End Sub

'Button to move something in daily back to on deck
'Either user error or cancelation
Private Sub MoveBackToDeckBtn_Click()

'Initialize and save the name in the dropdown
Dim shipName As String
shipName = ShipsDrop.Text

'Remove from Daily
DeleteFromDailyDB shipName

'Add to On Deck
PostToDeckDB shipName

'Write data to respective sheets from user selection
WriteLists

'Tally up items needed
CreateNeedsSheet

'Sort The Deck
SortRange

'Alert the user
MsgBox shipName & " has been put back on deck"

'Show deck to reload daily
DeckRadio.value = True
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

'Prints every label present in the labels box
Private Sub PrintAllLabels_Click()

'Use 0 index to get all items in the box
PrintBoxLabels 0, Worksheets("Home").LabelsBox.ListCount - 2

'Print a little label to mark the rolled up labels
'Using the name present in the dropdown box
PrintRollLabel ShipsDrop.Text
End Sub

'Print big labels to mark the wrapped pallets
Private Sub PrintSkidLabelBtn_Click()
PrintSkidLabel ShipsDrop.Text
End Sub

'Print out order sheets for ship displayed in dropdown
Private Sub OrderAndCheckBtn_Click()
PrintOrderAndCheck
End Sub

'Prints helper labels for orders with more than one pallet
Private Sub MultiSkidBtn_Click()
MultiSkid
End Sub

'Clicking the "Daily List" button uses method in "UserFormLogic" and other code in the form itself to trigger
Private Sub DailyListBtn_Click()
DisplayShipSelectForm
Worksheets("Needs").Activate
End Sub

'Navigation
'Brings you to Order Sheet
Private Sub NewOrderButton_Click()
Worksheets("Order").Activate
End Sub

'Used in conjuction with ShowFoundItems to create a dynamic filtering list of items in daily sheet
'Re-filters with every key stroke
Private Sub ItemSearchBar_Change()
ShowFoundItems ItemSearchBar.Text, ItemSearchBar2.Text
End Sub
'Key Code 9 is "tab"
'Tab key bouncing between search bars
Private Sub ItemSearchBar_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 9 Then ItemSearchBar2.Activate
End Sub

'Used in conjuction with ShowFoundItems to create a dynamic filtering list of items in daily sheet
'Re-filters with every key stroke
Private Sub ItemSearchBar2_Change()
ShowFoundItems ItemSearchBar.Text, ItemSearchBar2.Text
End Sub
'Key Code 9 is "tab"
'Tab key bouncing between search bars
Private Sub ItemSearchBar2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 9 Then ItemSearchBar.Activate
End Sub

'Displays results of what is in the search bars
Sub ShowFoundItems(LeftSearchTerm As String, RightSearchTerm As String)

'Initialize
'TODO: Work on variable names
Dim dailyRange As Range, i As Integer, last As Integer, j As Integer
Dim leftLookUp As String, rightLookup As String
'Clear the results box on key stroke
ItemSearchResults.Clear

'Add wild cards before and after whatever gets typed into the search bar
leftLookUp = "*" & LeftSearchTerm & "*"

If RightSearchTerm = "" Then
    rightLookup = leftLookUp
Else
    rightLookup = "*" & RightSearchTerm & "*"
End If

'Establish search area
last = Worksheets("Daily").Range("D" & Rows.Count).End(xlUp).Row
Set dailyRange = Worksheets("Daily").Range("A2:D" & last)

'List box row incrementor
j = 0

'Loop through the rows in daily
For i = 1 To last
    'If anything resembles the search term write it into the box
    If dailyRange.Cells(i, 3).value Like leftLookUp Or dailyRange.Cells(i, 3).value Like rightLookup Then
        ItemSearchResults.AddItem
        ItemSearchResults.list(j, 0) = dailyRange.Cells(i, 1).value
        ItemSearchResults.list(j, 1) = dailyRange.Cells(i, 2).value
        ItemSearchResults.list(j, 2) = dailyRange.Cells(i, 3).value
        ItemSearchResults.list(j, 3) = dailyRange.Cells(i, 4).value
        j = j + 1
    End If
Next i

ItemSearchResults.Height = 144.75

End Sub
