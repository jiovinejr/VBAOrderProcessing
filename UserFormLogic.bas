Attribute VB_Name = "UserFormLogic"

'Launch form when you find a missing item
Sub DisplayItemForm(item As String)

'Uses parameter to inform user which item
With AddToMasterForm
    .OldOrderNameDynamic.Caption = item
    .Prompt.Caption = "Item " & item & " not found in the Master List. Please fill out to add item."
    With .OrderNameBox
        .value = item
        .Enabled = False
    End With
    'Resets text boxes
    .NewNameBox.value = ""
    .CategoryBox.value = "Vegetable"
    .CaseWeightBox.value = ""
    .Show
End With

End Sub

'Insert data in Master List
Sub AddDataFromForm()
'Initialize variables
Dim orderName As String, newName As String
Dim category As String, caseWeight As Double

'Set vars with user input
With AddToMasterForm
    orderName = .OrderNameBox.Text
    newName = .NewNameBox.Text
    category = .CategoryBox.Text
    caseWeight = CDbl(.CaseWeightBox.Text)
End With

'Use DB method to add data
PostNewItemToMasterList orderName, newName, category, caseWeight

End Sub

'Launch when measurement abbreviation is not found
'e.g. "LB" for Pound
Sub DisplayMeasurementForm(oldMeasurement As String)

'Use class property to prompt
With MeasurementForm
    .OldItem.Caption = oldMeasurement
    .MeasurementPrompt.Caption = oldMeasurement & " doesn't exist in Master List. " & _
                                "Please enter full word for this abbreviation."
    'Reset
    .NewMeasurementBox.value = ""
    .Show
End With

End Sub

'Insert Measurement word into Master List
Sub AddMeasurementFromForm()

'Initialize
Dim measurementinput As String, OldItem As String

'Use input to set variables
With MeasurementForm
    OldItem = .OldItem.Caption
    
    'Make user input correct format
    'i.e. Proper Case
    'e.g. BoXeS = Boxes
    measurementinput = Application.WorksheetFunction.Proper(.NewMeasurementBox.Text)
    
End With

'Use DB method to add new measurement
PostNewMeasurmentToMasterList OldItem, measurementinput

End Sub

Sub DisplayShipSelectForm()
Dim listRange As Variant
Dim lastRow As Integer

lastRow = Worksheets("ShipsOnDeck").Range("A" & Rows.Count).End(xlUp).row
Set listRange = Worksheets("ShipsOnDeck").Range("A1:A" & lastRow)

With listRange
    .Sort key1:=.Cells(1, 1), _
              order1:=xlAscending, _
              Header:=xlNo
End With

With ShipSelectForm
    .ShipsOnDeck.RowSource = "ShipsOnDeck!A1:A" & lastRow
End With

ShipSelectForm.Show
    
End Sub

Sub AddShipsToDB()
Dim i As Integer, shipName As String

With ShipSelectForm
For i = 0 To .ShipsOnDeck.ListCount - 1
    shipName = CStr(.ShipsOnDeck.List(i))
    If .ShipsOnDeck.Selected(i) Then
        DeleteFromDeckDB shipName
        PostToDailyDB shipName
    End If
Next i
End With

End Sub


'TEST
Sub FormTest()
DisplayItemForm "TEST ITEM 100LBS"
End Sub
