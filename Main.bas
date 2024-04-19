Attribute VB_Name = "Main"
Sub ProcessPastedOrder()

Dim orderArr() As OrderRecord

orderArr = CreateRecordFromPaste
CreateBothReports orderArr
PostToOrderDB orderArr

End Sub
