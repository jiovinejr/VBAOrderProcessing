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

Sub CreateBothReports(arr() As OrderRecord)
CreateCheckSheet arr
CreateOrderSheet arr
End Sub
Sub CreateNeedsSheet()
Dim dict As Scripting.Dictionary, mapRange As Range, lastMapRange As Integer
Dim k As String, v As Double, cw As Double
Worksheets("Needs").Cells.ClearContents
Set dict = New Scripting.Dictionary

lastMapRange = Worksheets("Daily").Range("C" & Rows.Count).End(xlUp).Row
Set mapRange = Worksheets("Daily").Range("A2:C" & lastMapRange)


For Each r In mapRange.Rows
    k = r.Cells(, 3).value
    cw = Worksheets("Master List").Range("C3:C" & Worksheets("Master List").Range("C" & Rows.Count).End(xlUp).Row).Find(k).Offset(0, 2).value
    If r.Cells(, 2) = "Pound" Then
        v = Format((r.Cells(, 1).value / cw), "0.00")
    ElseIf r.Cells(, 2) = "Pint*" Then
        v = Format((r.Cells(, 1).value / 12), "0.00")
    ElseIf r.Cells(, 2) = "Pieces" Or r.Cells(, 2) = "Bunch" Or r.Cells(, 2) = "Each" Then
        v = Format((r.Cells(, 1).value / 40), "0.00")
    Else
        v = Format(r.Cells(, 1).value, "0.00")
    End If
    dict(k) = dict(k) + v
Next r

Dim key As Variant, writeRange As Range, i As Integer, sortHelp As Range

i = 1
Set writeRange = Worksheets("Needs").Range("A1:B" & dict.Count)
Set sortHelp = Worksheets("Needs").Range("A1:A" & dict.Count)


For Each key In dict.Keys
    Debug.Print key, dict(key)
    writeRange.Cells(i, 1) = key
    writeRange.Cells(i, 2) = dict(key)
    i = i + 1
Next key

writeRange.Sort sortHelp

End Sub
'Sub to check subs in this Mod
Sub CheckReportTest()
Dim orderArr() As OrderRecord
orderArr = CreateRecordFromPaste
CreateCheckSheet orderArr
CreateOrderSheet orderArr
End Sub
