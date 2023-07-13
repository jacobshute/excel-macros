Attribute VB_Name = "Test"

Sub test()
'
' Macro1 Macro
'
    SelectedRange = Selection.Rows.Count
    ActiveCell.Offset(0, 0).Select
    For i = 1 To SelectedRange
        If ActiveCell.Value = "" Then
            Selection.Value = "did it"
        End If
        
        ActiveCell.Offset(1, 0).Select
    Next i
    
End Sub
