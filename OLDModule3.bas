Attribute VB_Name = "OLDModule3"


Sub OldMakeStickers(arr() As OrderRecord, shipName As Variant)

Dim targetCell As Range, splitSize As Double
    
'TargetCell is where the result will start being written
Set targetCell = Worksheets("Label").Range("A1")
splitSize = 1

'Clear the previous labels, if there are any
With Worksheets("Label")
    If .Range("A1").value <> "" Then
        .Range("A1:C" & .Range("C" & .Rows.Count).End(xlUp).Row).Clear
    End If
End With

'Switch to label sheet
Worksheets("Label").Activate

Dim Quantity As Double, packaging As String, item As String, rowCounter As Long
Dim caseWeight As Double, i As Integer


'Loop through all rows of the order
For i = 1 To UBound(arr)
    Quantity = arr(i).Quantity
    packaging = arr(i).CleanMeasurement
    item = arr(i).CleanItem
    caseWeight = arr(i).ItemCaseWeight

    If packaging = "Bag" And item Like "*Radish*" Then
        ProcessBagRadish Quantity, packaging, item, targetCell, rowCounter
    ElseIf item Like "*Watermelon*" Then
        ProcessWatermelon Quantity, packaging, item, targetCell, rowCounter, caseWeight
    ElseIf packaging = "Pieces" Or packaging = "Bunch" Or packaging = "Pints" Or packaging = "Each" Or packaging = "Head" Then
        ProcessBunch Quantity, packaging, item, targetCell, rowCounter
    ElseIf packaging <> "Pound" Then
        ProcessNonPound Quantity, packaging, item, targetCell, rowCounter, splitSize
    Else
        ProcessPound Quantity, packaging, item, targetCell, rowCounter, caseWeight
    End If
Next i

End Sub

Sub GetSticksForOrder()

Dim lookUp As Range, ORDER As Range, arr() As Variant, shipName As String
Dim lastInDaily As Integer, shipRngStart As Integer, numOfRows As Integer, r As Integer

shipName = Worksheets("Label").Range("E1").Text
lastInDaily = Worksheets("Daily").Range("A" & Rows.Count).End(xlUp).Row

Set lookUp = Worksheets("Daily").Range("A1:D" & lastInDaily)

shipRngStart = 0
numOfRows = 0

With lookUp
    For r = 1 To .Rows.Count
        If .Cells(r, 4) = shipName Then
            If shipRngStart > 0 Then
                numOfRows = numOfRows + 1
            Else
                shipRngStart = r
                numOfRows = numOfRows + 1
            End If
        End If
    Next r
    
    Set ORDER = .Range("A" & shipRngStart & ":C" & (shipRngStart + numOfRows) - 1)
    
End With

arr = ORDER

MakeStickers arr, shipName


End Sub

Sub NewBreakdown()

Dim rng As Range, arr() As Variant
    
    'find the last row on the "order" sheet, assign the range with the order to the arr array
    With Worksheets("Check")
        Set rng = .Range("A4:C" & .Range("C" & .Rows.Count).End(xlUp).Row)
    End With
    arr = rng
    
    Dim shipName As String, Quantity As Double, packaging As String, item As String, rowCounter As Long, caseWeight As Double, i As Integer
    shipName = Worksheets("Check").Range("B1")
    Worksheets("Label").Range("E1").value = shipName
MakePDFs
MakeStickers arr, shipName
AddToOnDeck
FilterDeck
RefreshOnDeckPivot

End Sub

