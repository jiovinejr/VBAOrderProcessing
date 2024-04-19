Attribute VB_Name = "Database"
'Database Functionality

'Insert an order to be saved into the database
Sub PostToOrderDB(arr() As OrderRecord)

'Initialize
Dim db As Worksheet, targetRange As Range, shipName As String
Dim startRow As Integer, numberOfItems As Integer, i As Integer

'Set the ship name from the parameter array and save the DB sheet in a var
shipName = arr(1).ship
Set db = Worksheets("OrderDatabase")

'Checks DB for the ship name which includes a unique identifier inheirently
'If it find an occurance of the ship name it deletes all line items associated
'Just in case the re-entering has additions or subtractions
If Application.CountIf(db.Range("G:G"), shipName) > 0 Then
    DeleteFromOrderDB shipName
End If

'Find the first empty row
startRow = db.Range("A" & Rows.Count).End(xlUp).row + 1

'Length of the order to help with locating full orders later
numberOfItems = UBound(arr) + 1

'Use our variable to carve out the chunk of the DB we need
Set targetRange = db.Range("A" & startRow & ":G" & startRow + (numberOfItems - 1))

'Incrementor
i = 1

'Loop through array and write data to database
For Each ordRec In arr
    targetRange.Cells(i, 1) = ordRec.Quantity
    targetRange.Cells(i, 2) = ordRec.OrderMeasurement
    targetRange.Cells(i, 3) = ordRec.OrderItem
    targetRange.Cells(i, 4) = ordRec.CleanMeasurement
    targetRange.Cells(i, 5) = ordRec.CleanItem
    targetRange.Cells(i, 6) = ordRec.ItemCaseWeight
    targetRange.Cells(i, 7) = ordRec.ship
    'Increment
    i = i + 1
Next ordRec

'Also add the shipname and number of items to help with location of full order later
PostToShipDB shipName, numberOfItems

End Sub

'Adds ship name and number of items to help with location of full order later
Sub PostToShipDB(shipName As String, lineItemCount As Integer)

'Initialize destination variables
Dim db As Worksheet, targetRow As Integer

'Set those vatiables
Set db = Worksheets("ShipDatabase")

'Uses the algorithm to find the lasr row containing data and adds 1 to give the first empty row
targetRow = db.Range("A" & Rows.Count).End(xlUp).row + 1

'Write data to destination
db.Range("A" & targetRow) = shipName
db.Range("B" & targetRow) = lineItemCount

'Adds new ship to On Deck sheet
PostToDeckDB shipName

End Sub

'Used by userform to add item to Master List
Sub PostNewItemToMasterList(orderName As String, newName As String, category As String, caseWeight As Double)

'Initialize
Dim master As Worksheet, targetRow As Integer, targetRange As Range

'Set where to write data
Set master = Worksheets("Master List")
targetRow = master.Cells(Rows.Count, "B").End(xlUp).row + 1
Set targetRange = master.Range("B" & targetRow & ":E" & targetRow)

'Write data
With targetRange
    .Cells(, 1) = orderName
    .Cells(, 2) = newName
    .Cells(, 3) = category
    .Cells(, 4) = caseWeight
End With
End Sub

'Used by new measurement form
Sub PostNewMeasurmentToMasterList(orderMeasurementName As String, newMeasurementName As String)

'Initialize
Dim master As Worksheet, targetRow As Integer, targetRange As Range

'Set writing destination
Set master = Worksheets("Master List")
targetRow = master.Cells(Rows.Count, "F").End(xlUp).row + 1
Set targetRange = master.Range("F" & targetRow & ":G" & targetRow)

'Write
With targetRange
    .Cells(, 1) = orderMeasurementName
    .Cells(, 2) = newMeasurementName
End With

End Sub

'Master function for adding a single ship to a sheet
Sub PostShipName(shipName As String, sheetName As String)

'Initialize
Dim db As Worksheet, targetRow As Integer, targetCell As Range

'Establish a destination
Set db = Worksheets(sheetName)
targetRow = db.Range("A" & Rows.Count).End(xlUp).row + 1
Set targetCell = db.Range("A" & targetRow)

'Write the shipName to the destination
targetCell.value = shipName
End Sub

'Sends a ship name to Deck
Sub PostToDeckDB(shipName As String)
PostShipName shipName, "ShipsOnDeck"
End Sub

'Sends a ship name to Daily
Sub PostToDailyDB(shipName As String)
PostShipName shipName, "DailyDatabase"
End Sub


'Deletes data from the DB by ship name which has an inheirent unique identifier
Sub DeleteFromOrderDB(shipName As String)

'Initialize
Dim db As Worksheet, allShipsRange As Range
Dim startRowOfOrder As Integer, numOfItems As Integer

'Set sheet and search range for ship names
Set db = Worksheets("OrderDatabase")
Set allShipsRange = db.Range("G:G")

'Finds the row number of the first instance of the ship name
'i.e. Where to start deleting
startRowOfOrder = allShipsRange.Find(shipName).row

'Establish how many items to delete
numOfItems = Application.WorksheetFunction.XLookup(shipName, Worksheets("ShipDatabase").Range("A:A"), Worksheets("ShipDatabase").Range("B:B"))

'Since deleting a row shifts everything up, the rownumber to delete doesn't change
'Just need to delete for the number of items in an order
For i = 1 To numOfItems
    allShipsRange.Rows(startRowOfOrder).EntireRow.Delete
Next

'Also when that's done, delete from the ship DB
DeleteFromShipDB shipName

End Sub

'Used in conjunction with DeleteFromOrderDB
'Deletes helper record in ShipDB
Sub DeleteFromShipDB(shipName As String)
DeleteSingleShipFromDB shipName, "ShipDatabase"
DeleteFromDeckDB shipName
End Sub

'Master function for deleting a single ship from a sheet
Sub DeleteSingleShipFromDB(shipName As String, sheetName As String)

'Initialize
Dim db As Worksheet, allShipsRange As Range
Dim shipRow As Integer

'Establish search area
Set db = Worksheets(sheetName)
Set allShipsRange = db.Range("A:A")

'Find shipname
shipRow = allShipsRange.Find(shipName).row

'Delete row from DB
allShipsRange.Rows(shipRow).EntireRow.Delete

End Sub

'Deletes a shipName from On Deck
Sub DeleteFromDeckDB(shipName As String)
DeleteSingleShipFromDB shipName, "ShipsOnDeck"
End Sub

'Deletes a shipName from Daily
Sub DeleteFromDailyDB(shipName As String)
DeleteSingleShipFromDB shipName, "DailyDatabase"
End Sub

'Its a new day clear out
Sub ClearDailyDB()
Worksheets("DailyDatabase").Range("A2:F300").ClearContents
End Sub

'TEST
Sub AddDBTest()
Dim arr() As OrderRecord
arr = CreateRecordFromPaste
PostToOrderDB arr
End Sub

'TEST
Sub DeleteTest()
arr = CreateRecordFromPaste
DeleteFromOrderDB arr(1).ship
End Sub
