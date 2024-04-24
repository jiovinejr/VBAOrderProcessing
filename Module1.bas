Attribute VB_Name = "Module1"
'Print Module
Sub MultiSkid()

Dim labelPath As String, skids As Variant, i As Integer, t As String, numOfSkids As Integer
Dim ObjDoc As bpac.Document
Set ObjDoc = CreateObject("bpac.Document")

labelPath = "C:\Users\vince\OneDrive\Desktop\Delaware Ship\Protected Folder - DO NOT DELETE\ZeeMulti.lbx"
skids = InputBox("How many skids?", "MultiSkid", "2")

If skids <> "" Then
numOfSkids = CInt(skids)

ObjDoc.Open (labelPath)
        ObjDoc.StartPrint "", bpoCutAtEnd
        
        For i = 1 To numOfSkids
        
            t = i & " of " & numOfSkids
            
            ObjDoc.GetObject("Multi").Text = t
            ObjDoc.PrintOut 2, bpoDefault
        
        Next i
    
    ObjDoc.EndPrint
    ObjDoc.Close
End If


End Sub
Sub PrintBoxLabels(begin As Variant, last As Variant)
Dim sheetname As String, labelPath As String, shipName As String, i As Integer

Dim ObjDoc As bpac.Document, kg As Double
    Set ObjDoc = CreateObject("bpac.Document")

    labelPath = "C:\Users\vince\OneDrive\Desktop\Delaware Ship\Protected Folder - DO NOT DELETE\ZeeCaseLabels2.lbx"
    shipName = Worksheets("Label").Range("E1").Text
    
  
    sheetname = ActiveSheet.Name
    
    ObjDoc.Open (labelPath)
        ObjDoc.StartPrint "", bpoCutAtEnd

Dim qty As String, meas As String, item As String
            For i = begin To last
                kg = Format(Range("A" & i) / 2.2, "0.00")
                qty = Range("A" & i).Text
                meas = Range("B" & i).Text
                item = Range("C" & i).Text
                
                If sheetname <> "Label" Then
                    shipName = Range("D" & i).Text
                End If
                
                ObjDoc.GetObject("DelShip").Text = "Delaware Ship Supply Co."
                
                ObjDoc.GetObject("Ship").Text = shipName
                
                ObjDoc.GetObject("Qty").Text = qty
                
                ObjDoc.GetObject("Measure").Text = meas
                
                ObjDoc.GetObject("Item").Text = item
                
                If kg <> 0 Then
                    ObjDoc.GetObject("Kilo").Text = "(" & kg & " Kilo)"
                Else
                    ObjDoc.GetObject("Kilo").Text = ""
                End If
                
                ObjDoc.PrintOut 1, bpoDefault
            Next i
            
            
        ObjDoc.EndPrint
    ObjDoc.Close
    
End Sub


Sub LabelsForFullOrder()
Dim l As Integer, shipName As String
    l = Worksheets("Label") _
        .Range("C" & Rows.Count).End(xlUp).Row
        
    shipName = Worksheets("Label").Range("E1").Text
    
    PrintBoxLabels 1, l
    
    PrintRollLabel
End Sub
Sub SelectedLabels()

    Dim Selected As Integer, r As Integer, last As Integer

    Selected = Selection.Areas(1).Rows.Count
    r = CInt(Selection.Areas(1).Cells.Row)
    last = (r + Selected) - 1
    
    PrintBoxLabels r, last
    
End Sub
Sub PrintOrderAndCheck()
    Dim orderRng As Variant, checkRng As Variant, mainFolder As String
    Dim shipName As String, lastInOrder As Integer, ship As String, filePath As String
    
    ship = Worksheets("Label").Range("E1")
    
    mainFolder = "C:\Users\vince\OneDrive\Desktop\Delaware Ship\OrderPDFs\"
    filePath = mainFolder & ship & "\" & ship
    
    shipName = Worksheets("Check").Range("B1")
    
    lastInOrder = Worksheets("Order").Range("A" & Rows.Count).End(xlUp).Row
    
    Application.ActivePrinter = "ET-5880 Series(Network) on Ne05:"
    
    Set orderRng = Worksheets("Order").Range("A1", "E" & lastInOrder)
    Set checkRng = Worksheets("Check").Range("A1", "D" & lastInOrder)
    
    If ship = shipName Then
        checkRng.PrintOut
        orderRng.PrintOut
    Else
        
        PrintFile filePath & "-check.pdf"
        Application.Wait (Now + TimeValue("0:00:04"))
        PrintFile filePath & "-order.pdf"
        
        
    End If
    
    '"C:\Users\vince\OneDrive\Desktop\Delaware Ship\OrderPDFs\MV GRAND PIONEER-328995\MV GRAND PIONEER-328995-order.pdf"
    
End Sub
Sub PrintSkidLabel()
    
    Dim labelPath As String, shipName As String

    Dim ObjDoc As bpac.Document
    Set ObjDoc = CreateObject("bpac.Document")

    labelPath = "C:\Users\vince\OneDrive\Desktop\Delaware Ship\Protected Folder - DO NOT DELETE\ZeeSkidLabel.lbx"
    shipName = Worksheets("Label").Range("E1").Text
    
    ObjDoc.Open (labelPath)
        ObjDoc.StartPrint "", bpoDefault
            ObjDoc.GetObject("ShipName").Text = shipName
            ObjDoc.PrintOut 1, bpoDefault
            ObjDoc.PrintOut 1, bpoDefault
        ObjDoc.EndPrint
    ObjDoc.Close

    
End Sub



Sub PrintRollLabel()

    Dim ObjDoc As bpac.Document, labelPath As String, shipSend As String
    Set ObjDoc = CreateObject("bpac.Document")

    labelPath = "C:\Users\vince\OneDrive\Desktop\Delaware Ship\Protected Folder - DO NOT DELETE\ZeeRollLabel.lbx"
    shipSend = Worksheets("Label").Range("E1").Text

    
    ObjDoc.Open (labelPath)
        ObjDoc.StartPrint "", bpoDefault
            ObjDoc.GetObject("RollLabel").Text = shipSend
            ObjDoc.PrintOut 1, bpoDefault
        ObjDoc.EndPrint
    ObjDoc.Close
 
 

End Sub



