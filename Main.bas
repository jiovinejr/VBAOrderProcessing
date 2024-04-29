Attribute VB_Name = "Main"
Sub ProcessPastedOrder()

Dim orderArr() As OrderRecord, home As Worksheet
Dim shipName As String

orderArr = CreateRecordFromPaste
PostToOrderDB orderArr
CreateBothReports orderArr
shipName = orderArr(1).ship
Worksheets("Home").ShipsDrop.value = CStr(shipName)
Worksheets("Home").DailyRadio.value = True
WriteLists
End Sub

Sub ddTest()
Worksheets("Home").ShipsDrop.value = "hello"
End Sub
