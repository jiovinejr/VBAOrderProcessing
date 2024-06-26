Attribute VB_Name = "Main"

'Routine for when a new order is being processed
Sub ProcessPastedOrder()

'Initialize
Dim orderArr() As OrderRecord, home As Worksheet
Dim shipName As String

'Use the new order that is pasted to the order sheet
'To crate an array of order record objects
orderArr = CreateRecordFromPaste

'Add the order (unsorted) to the DB
'That way the OrderPrint sheet always matches the original order
PostToOrderDB orderArr

'This routine will sort the array so do this last
CreateBothReports orderArr

'Retrieve the ship name from an object in the array
'Preferably the first, JIC the array is length 1
shipName = orderArr(0).ship

'Make the dropdown menu display this ship so the list of labels will appear
Worksheets("Home").ShipsDrop.value = CStr(shipName)

'Switch the radio buttons
Worksheets("Home").DeckRadio.value = True

'Re-write the Daily and Deck item lists to include new order
WriteLists


End Sub

'Button to paste order that is copied from google sheets
Sub PasteSpeacial()

'Try
On Error Resume Next
    'Clear the sheet
    Worksheets("Order").Range("A1:C200").ClearContents
    'Select to left cell
    Worksheets("Order").Range("A1").Select
    'Paste
    Worksheets("Order").PasteSpecial Format:="HTML", Link:=False, DisplayAsIcon:=False, NoHTMLFormatting:=True
    
'Catch
'If nothing is copied alert the user
If Err.Number <> 0 Then
    MsgBox ("Nothing copied. Copy something, will ya!")
End If

End Sub

'TEST to check how to manipulate whats in the dropdown menu window
Sub ddTest()
Worksheets("Home").ShipsDrop.value = "hello"
End Sub
