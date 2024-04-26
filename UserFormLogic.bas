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
    newName = Application.WorksheetFunction.Proper(.NewNameBox.Text)
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

'Form for selecting what ships for today
Sub DisplayShipSelectForm()

'Initialize
Dim listRange As Range, shipList As Variant
Dim lastRow As Integer

'Establish the range where the ships on deck are listed
lastRow = Worksheets("ShipsOnDeck").Range("A" & Rows.Count).End(xlUp).Row
Set listRange = Worksheets("ShipsOnDeck").Range("A1:A" & lastRow)
shipList = GetShipsFromDB("ShipsOnDeck")
check = VarType(shipList)

If WorksheetFunction.CountA(listRange) <> 0 Then
        SortRange
    'Make the list box draw data from range using location string rather than range method
    'Just a finicky part of VBA
    With ShipSelectForm
        If check = 8 Then
            .ShipsOnDeckBox.AddItem (shipList)
        Else
            .ShipsOnDeckBox.List = shipList
        End If
    End With
    
    'Show the form
    ShipSelectForm.Show
Else
    MsgBox "Empty Deck"
    Exit Sub
End If
End Sub

Sub SortRange()

'Initialize
Dim listRange As Variant
Dim lastRow As Integer

'Establish the range where the ships on deck are listed
lastRow = Worksheets("ShipsOnDeck").Range("A" & Rows.Count).End(xlUp).Row
Set listRange = Worksheets("ShipsOnDeck").Range("A1:A" & lastRow)

'Sort the range
With listRange
    .Sort key1:=.Cells(1, 1), _
              order1:=xlAscending, _
              Header:=xlNo
End With
End Sub

'Method to handle user input upon submittion
Sub AddShipsToDB()

'Initialize
Dim i As Integer, shipName As String

'In the form
With ShipSelectForm
'Loop through the list box
For i = 0 To .ShipsOnDeckBox.ListCount - 1
    'Helper variable to cast input to String
    shipName = CStr(.ShipsOnDeckBox.List(i))
    'When you come across a selected item
    If .ShipsOnDeckBox.Selected(i) Then
        'Take it out of the on deck sheet
        DeleteFromDeckDB shipName
        'Put into the daily sheet
        PostToDailyDB shipName
    End If
Next i
End With

End Sub

'TEST
Sub FormTest()
DisplayItemForm "TEST ITEM 100LBS"
End Sub
