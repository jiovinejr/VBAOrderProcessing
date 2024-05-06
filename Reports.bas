Attribute VB_Name = "Reports"
'Create Reports

'Create check sheet from an order array
Sub CreateCheckSheet(arr() As OrderRecord)

'Initialize
Dim sortedArr() As OrderRecord, checkSheet As Worksheet
Dim checkRange As Range, checkShipName As Range

'Sort the incoming array
sortedArr = SortOrderRecord(arr)

'Set excel poits to variables
Set checkSheet = Worksheets("CheckPrint")
Set checkShipName = checkSheet.Range("B1")
Set checkRange = checkSheet.Range("A4:C" & UBound(sortedArr))

'Clear anything and set headers
checkSheet.Cells.ClearContents
checkSheet.Range("A1").value = "Name:"
checkSheet.Range("A2").value = "Date:"
checkSheet.Range("D3").value = "Notes"

'Set ship name
checkShipName.value = sortedArr(1).ship

'Incrementor
Dim i As Integer
i = 1

'Loop through array and write data to cells
For Each ordRec In sortedArr
    checkRange.Cells(i, 1) = ordRec.Quantity
    checkRange.Cells(i, 2) = ordRec.CleanMeasurement
    checkRange.Cells(i, 3) = ordRec.CleanItem
    'Next row
    i = i + 1
Next ordRec

'Hide sheet
checkSheet.Visible = xlSheetHidden

End Sub

'Create an order sheet from an order array
Sub CreateOrderSheet(arr() As OrderRecord)

'Initialize
Dim orderSheet As Worksheet
Dim orderRange As Range, orderShipName As Range

'Set excel points to variables
Set orderSheet = Worksheets("OrderPrint")
Set orderShipName = orderSheet.Range("C1")
Set orderRange = orderSheet.Range("A4:C" & UBound(arr))

'Clear out sheet
orderSheet.Cells.ClearContents

'Set the ship name
orderShipName.value = arr(1).ship

'Incrementor
Dim i As Integer
i = 1

'Loop through un-sorted array
For Each ordRec In arr
    orderRange.Cells(i, 1) = ordRec.Quantity
    orderRange.Cells(i, 2) = ordRec.OrderMeasurement
    orderRange.Cells(i, 3) = ordRec.OrderItem
    'Next row
    i = i + 1
Next ordRec

'Hide sheet
orderSheet.Visible = xlSheetHidden

End Sub

'Used to bring up both order and check for any given order
'Just in case you need to reprint
Sub CreateBothReports(arr() As OrderRecord)

'Create the order sheet first so that the array going in isn't sorted yet
CreateOrderSheet arr

'Will sort arr
CreateCheckSheet arr

End Sub

'Sub to check subs in this Mod
Sub CheckReportTest()
Dim orderArr() As OrderRecord
orderArr = CreateRecordFromPaste
CreateCheckSheet orderArr
CreateOrderSheet orderArr
End Sub


'Used to loop through add items to either on deck sheet or daily sheet for reoorting and analysis
Sub WriteToItemList(arr() As OrderRecord, sheetName As String)

'Initialize
Dim db As Worksheet, targetRange As Range
Dim startRow As Integer, numberOfItems As Integer, i As Integer

'Set the DB sheet to a var
Set db = Worksheets(sheetName)

'Find the first empty row
startRow = db.Range("A" & Rows.Count).End(xlUp).Row + 1

'Length of the order to help with locating full orders later
numberOfItems = UBound(arr) + 1

'Use our variable to carve out the chunk of the DB we need
Set targetRange = db.Range("A" & startRow & ":D" & startRow + (numberOfItems - 1))

'Incrementor
i = 1

'Loop through array and write data to database
For Each ordRec In arr
    targetRange.Cells(i, 1) = ordRec.Quantity
    targetRange.Cells(i, 2) = ordRec.CleanMeasurement
    targetRange.Cells(i, 3) = ordRec.CleanItem
    targetRange.Cells(i, 4) = ordRec.ship
    'Increment
    i = i + 1
Next ordRec
End Sub

'Writes item lists for on deck and daily for reporting
Sub WriteLists()

'Initialize
Dim dailyArr As Variant, deckArr As Variant, ship As Variant
Dim ordRec() As OrderRecord, sorted() As OrderRecord
Dim dailyType As Integer, deckType As Integer

'Clear out old data from sheets while keeping the headers
Worksheets("Daily").Range("A2:D10000").ClearContents
Worksheets("On Deck").Range("A2:D10000").ClearContents

'Set the list of ships in the Daily DB
dailyArr = GetShipsFromDB("DailyDatabase")

'Find out the data type of Variant
dailyType = VarType(dailyArr)

'If the variable has any data
If Not IsEmpty(dailyArr) Then
    'If there's only 1 ship in the DB, dailyArr will be saved as a String rather than an array length 1
    'So if it isn't a string (data type number 8)
    If dailyType <> 8 Then
        'Loop through all the ship names in DB
        For Each ship In dailyArr
            'Create a record using the ship name
            ordRec = CreateRecordFromDB(CStr(ship))
            'Sort the current order
            sorted = SortOrderRecord(ordRec)
            'Use our helper method to write the order record to the Daily list
            WriteToItemList sorted, "Daily"
        Next ship
    'If the data type is a string
    Else
        'Run through the same motions but using dailyArr as the string rather than looping through
        ordRec = CreateRecordFromDB(CStr(dailyArr))
        sorted = SortOrderRecord(ordRec)
        WriteToItemList sorted, "Daily"
    End If
End If

'Do the same thing with the ships on deck
'Should farm this logic out to separate methods
deckArr = GetShipsFromDB("ShipsOnDeck")
deckType = VarType(deckArr)

If Not IsEmpty(deckArr) Then
    If deckType <> 8 Then
        For Each ship In deckArr
            ordRec = CreateRecordFromDB(CStr(ship))
            sorted = SortOrderRecord(ordRec)
            WriteToItemList sorted, "On Deck"
        Next ship
    Else
        ordRec = CreateRecordFromDB(CStr(deckArr))
        sorted = SortOrderRecord(ordRec)
        WriteToItemList sorted, "On Deck"
    End If
End If

End Sub


'Needs sheet takes the items in the daily
'Gives a list of items needed for the day by case
Sub CreateNeedsSheet()

'Initialize
Dim dict As Scripting.Dictionary, mapRange As Range, lastMapRange As Integer
Dim k As String, v As Double, cw As Double

'Clear any data inside the sheet
Worksheets("Needs").Cells.ClearContents

'Set a new Map/dictionary
Set dict = New Scripting.Dictionary

'mapRange is all the data in the daily sheet
lastMapRange = Worksheets("Daily").Range("C" & Rows.Count).End(xlUp).Row
Set mapRange = Worksheets("Daily").Range("A2:C" & lastMapRange)

'Loop through the rows in the Daily sheet
'TODO: Work on these god awful variable names
For Each r In mapRange.Rows
    'Set the Name of the item to the key in our Map
    k = r.Cells(, 3).value
    'Find the case weight of the item in question
    cw = Worksheets("Master List").Range("C3:C" & Worksheets("Master List").Range("C" & Rows.Count).End(xlUp).Row).Find(k).Offset(0, 2).value
    
    'Go through different case scenarios of measurement to set a value for the map
    'Using case weight to find out how many cases would be needed to fulfill this record
    If r.Cells(, 2) = "Pound" Then
        v = Format((r.Cells(, 1).value / cw), "0.00")
    ElseIf r.Cells(, 2) = "Pint*" Then
        v = Format((r.Cells(, 1).value / 12), "0.00")
    ElseIf r.Cells(, 2) = "Pieces" Or r.Cells(, 2) = "Bunch" Or r.Cells(, 2) = "Each" Then
        v = Format((r.Cells(, 1).value / 40), "0.00")
    Else
        v = Format(r.Cells(, 1).value, "0.00")
    End If
    'This will either add key value pair to map
    'Or update key's value
    dict(k) = dict(k) + v
Next r

'Initialize some more variables
Dim key As Variant, writeRange As Range, i As Integer, sortHelp As Range

'Incrementor
i = 1

'Establis a range to write our Map data
Set writeRange = Worksheets("Needs").Range("A1:B" & dict.Count)
Set sortHelp = Worksheets("Needs").Range("A1:A" & dict.Count)

'Loop through the Map and write the data to the rows
For Each key In dict.Keys
    writeRange.Cells(i, 1) = key
    writeRange.Cells(i, 2) = dict(key)
    i = i + 1
Next key

'Sort by column A
writeRange.Sort sortHelp

End Sub
