Attribute VB_Name = "Breakdown"
'Creates the list of stickers to be printed
Sub MakeStickers(arr() As OrderRecord, shipName As Variant)

Dim targetCell As Object, splitSize As Double
    
'TargetCell is where the result will start being written
Set targetCell = Worksheets("Home").LabelsBox
splitSize = 1

'Clear the previous labels, if there are any
With Worksheets("Home")
    targetCell.Clear
End With

'Switch to Home
Worksheets("Home").Activate

Dim Quantity As Double, packaging As String, item As String, rowCounter As Long
Dim caseWeight As Double, i As Integer
rowCounter = 0

'Loop through all rows of the order
For i = 0 To UBound(arr)
    Quantity = arr(i).Quantity
    packaging = arr(i).CleanMeasurement
    item = arr(i).CleanItem
    caseWeight = arr(i).ItemCaseWeight
    
    'Process based on different scenarios
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

WriteLabel "", "", "", targetCell, rowCounter

End Sub

Sub ProcessBagRadish(Quantity As Double, packaging As String, item As String, targetCell As Object, ByRef rowCounter As Long)
    While Quantity > 30
        WriteLabel 30, packaging, item, targetCell, rowCounter
        Quantity = Quantity - 30
    Wend
    WriteLabel Quantity, packaging, item, targetCell, rowCounter
End Sub

Sub ProcessWatermelon(Quantity As Double, packaging As String, item As String, targetCell As Object, ByRef rowCounter As Long, caseWeight As Double)
    While Quantity > caseWeight
        WriteLabel "", packaging, item, targetCell, rowCounter
        Quantity = Quantity - caseWeight
    Wend
    WriteLabel "", packaging, item, targetCell, rowCounter
End Sub

Sub ProcessBunch(Quantity As Double, packaging As String, item As String, targetCell As Object, ByRef rowCounter As Long)
    While Quantity > 48
        WriteLabel 48, packaging, item, targetCell, rowCounter
        Quantity = Quantity - 48
    Wend
    WriteLabel Quantity, packaging, item, targetCell, rowCounter
End Sub

Sub ProcessNonPound(Quantity As Double, packaging As String, item As String, targetCell As Object, ByRef rowCounter As Long, splitSize As Double)
    While Quantity > splitSize
        WriteLabel splitSize, packaging, item, targetCell, rowCounter
        Quantity = Quantity - splitSize
    Wend
    WriteLabel Quantity, packaging, item, targetCell, rowCounter
End Sub

Sub ProcessPound(Quantity As Double, packaging As String, item As String, targetCell As Object, ByRef rowCounter As Long, caseWeight As Double)
    While Quantity > caseWeight
        WriteLabel caseWeight, packaging, item, targetCell, rowCounter
        Quantity = Quantity - caseWeight
    Wend
    WriteLabel Quantity, packaging, item, targetCell, rowCounter
End Sub

Sub WriteLabel(Quantity As Variant, packaging As String, item As String, targetCell As Object, ByRef rowCounter As Long)
    ' Write label information to the list box and increment the row counter
    targetCell.AddItem
    targetCell.list(rowCounter, 0) = Quantity
    targetCell.list(rowCounter, 1) = packaging
    targetCell.list(rowCounter, 2) = item
    targetCell.Height = 555
    rowCounter = rowCounter + 1
End Sub

