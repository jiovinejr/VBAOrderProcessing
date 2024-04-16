Attribute VB_Name = "UserFormLogic"
Sub DisplayItemForm(item As String)

With AddToMaster
    .OldOrderNameDynamic.Caption = item
    .Prompt.Caption = "Item " & item & " not found in the Master List. Please fill out to add item."
    With .OrderNameBox
        .value = item
        .Enabled = False
    End With
    .Show
End With

End Sub

Sub FormTest()
DisplayItemForm "TEST ITEM 100LBS"
End Sub
