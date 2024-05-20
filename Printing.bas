Attribute VB_Name = "Printing"
'Constants Initialize
Const LABEL_PATH As String = "C:\Users\vince\OneDrive\Desktop\Delaware Ship\Protected Folder - DO NOT DELETE\"
Const SKID_LABEL_FILE As String = "ZeeSkidLabel.lbx"
Const ROLL_LABEL_FILE As String = "ZeeRollLabel.lbx"
Const BOX_LABEL_FILE As String = "ZeeCaseLabels2.lbx"
Const MULTI_SKID_LABEL_FILE As String = "ZeeMulti.lbx"

'Initialize document with bPac SDK
Dim ObjDoc As bpac.Document


'Print out a series of labels for more than one pallet
'e.g. "1 of 2, 2 of 2"
Sub MultiSkid()

'Initialize
Dim labelPath As String, skids As Variant
Dim i As Integer, textSend As String, numOfSkids As Integer
Set ObjDoc = CreateObject("bpac.Document")

'Set path with constants
labelPath = LABEL_PATH & MULTI_SKID_LABEL_FILE

'Store user input
skids = InputBox("How many skids?", "MultiSkid", "2")

'Handle NaN error
If Not IsNumeric(skids) Then
    MsgBox "Enter valid number"
Else
    'If skids isn't blank
    If skids <> "" Then
    'Cast input to an integer
    numOfSkids = CInt(skids)
    
    'Open the Ptouch document
    ObjDoc.Open (labelPath)
            
        'Start printing and cut when printing is finished
        ObjDoc.StartPrint "", bpoCutAtEnd
        
            'Loop to the amount of
            For i = 1 To numOfSkids
            
                'i is equal to the current pallet and print the total pallets
                textSend = i & " of " & numOfSkids
                
                'Find the object in the Ptouch file and set it to the textSend
                ObjDoc.GetObject("Multi").Text = textSend
                
                'Print out 2 copies
                ObjDoc.PrintOut 2, bpoDefault
            
            'Iterate
            Next i
        
        'Stop the Print and close the file
        ObjDoc.EndPrint
        ObjDoc.Close
    End If
End If

End Sub

'Prints a series of labels by list box index
Sub PrintBoxLabels(begin As Variant, last As Variant)

'Initialize variables and set object document
Dim sheetName As String, labelPath As String, shipName As String, i As Integer
Dim kg As Variant, labels As Variant
Set ObjDoc = CreateObject("bpac.Document")

'Set path with constants, ship name from the dropdown menu
'And create an array of list items
labelPath = LABEL_PATH & BOX_LABEL_FILE
shipName = Worksheets("Home").ShipsDrop.Text
labels = Worksheets("Home").LabelsBox.list

'Open the document and start a print
ObjDoc.Open (labelPath)
ObjDoc.StartPrint "", bpoCutAtEnd

'Initialize variable from the data
Dim qty As String, meas As String, item As String

'Loop through the array
For i = begin To last
    
    'Set variables
    qty = labels(i, 0)
    meas = labels(i, 1)
    item = labels(i, 2)
    
    'Sometimes qty will be an empty string that won't cast to 0
    'So we have to handle that with a conditional
    If qty = "" Then
        kg = 0
    Else
        kg = Format(CDbl(labels(i, 0)) / 2.2, "0.00")
    End If
    
    'Set the objects in the Ptouch file to variables
    ObjDoc.GetObject("DelShip").Text = "Delaware Ship Supply Co."
    ObjDoc.GetObject("Ship").Text = shipName
    ObjDoc.GetObject("Qty").Text = qty
    ObjDoc.GetObject("Measure").Text = meas
    ObjDoc.GetObject("Item").Text = item
    
    'If there was math done on KG
    'Include it on the label
    If kg <> 0 Then
        ObjDoc.GetObject("Kilo").Text = "(" & kg & " Kilo)"
    Else
        'If not pass in a blank string
        ObjDoc.GetObject("Kilo").Text = ""
    End If
    
    'Send label data to the printer
    ObjDoc.PrintOut 1, bpoDefault

'Iterate
Next i

'Stop print and close file
ObjDoc.EndPrint
ObjDoc.Close
    
End Sub

'Print a large sticker to mark the whole order
'Pass in the ship name
Sub PrintSkidLabel(shipName As String)

'Initialize, set the object document, and set the location of the document
Dim labelPath As String
Set ObjDoc = CreateObject("bpac.Document")
labelPath = LABEL_PATH & SKID_LABEL_FILE

'Open doc, start a print
ObjDoc.Open (labelPath)
    ObjDoc.StartPrint "", bpoDefault
        'Fill object in file with parameter
        ObjDoc.GetObject("ShipName").Text = shipName
        'Print out 2 but cut between
        ObjDoc.PrintOut 1, bpoDefault
        ObjDoc.PrintOut 1, bpoDefault
    'Stop Print
    ObjDoc.EndPrint
'Close file
ObjDoc.Close

End Sub

'Print little roll label
Sub PrintRollLabel(shipName As String)

'Initialize, set the object document, and set the location of the document
Dim labelPath As String
Set ObjDoc = CreateObject("bpac.Document")
labelPath = LABEL_PATH & ROLL_LABEL_FILE

'Same motions as others
ObjDoc.Open (labelPath)
    ObjDoc.StartPrint "", bpoDefault
        ObjDoc.GetObject("RollLabel").Text = shipName
        'Only print out 1
        ObjDoc.PrintOut 1, bpoDefault
    ObjDoc.EndPrint
ObjDoc.Close

End Sub

'Print out hidden sheets for order and check
Sub PrintOrderAndCheck()

'Initialize and set
Dim orderRng As Variant, checkRng As Variant, lastInOrder As Integer
'last in order sheet will match last on check sheet
lastInOrder = Worksheets("OrderPrint").Range("A" & Rows.Count).End(xlUp).Row
Set orderRng = Worksheets("OrderPrint").Range("A1", "E" & lastInOrder)
Set checkRng = Worksheets("CheckPrint").Range("A1", "D" & lastInOrder)

'TODO: select printer
Application.ActivePrinter = "ET-5880 Series(Network) on Ne05:"

'Unhide both sheets
Worksheets("OrderPrint").Visible = True
Worksheets("CheckPrint").Visible = True

'Print the ranges
checkRng.PrintOut
orderRng.PrintOut

'Hide the sheets when finished
Worksheets("OrderPrint").Visible = False
Worksheets("CheckPrint").Visible = False

End Sub

'Prints the needs map
Sub PrintNeedsSheet()

'Initialize and set
Dim last As Integer, needsRng As Range
last = Worksheets("Needs").Range("A" & Rows.Count).End(xlUp).Row
Set needsRng = Worksheets("Needs").Range("A1:B" & last)

'Print range
needsRng.PrintOut

End Sub

