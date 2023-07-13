Attribute VB_Name = "iWire"
Attribute VB_Name = "Module3"
Sub iWireOut()
Attribute iWireOut.VB_ProcData.VB_Invoke_Func = " \n14"
'
' This macro builds a text file for the IRS iWire system from an Excel file
' https://learn.microsoft.com/en-us/office/vba/excel/concepts/cells-and-ranges/create-an-html-file-with-a-table-of-contents-based-on-cell-data
'
' VB main lang tutorial
' https://learn.microsoft.com/en-us/dotnet/visual-basic/

    Dim f As String
    f = ActiveWorkbook.Path & "test.txt"
    Close
    Open f For Output As #1
    
    Dim row As Long
    row = 2
    Dim column As Long
    column = 1
    Dim a As String
    a = "test"
    
    Dim test As String
    
    
    
    Cells(2, 1).Select
    
    
    Do While WorksheetFunction.CountA(Rows(row)) > 0
        
        Select Case ActiveCell.Value
        Case "T"
            a = ActiveCell.Value
            Do While ActiveCell <> ""
                ActiveCell.Offset(0, 1).Select
                a = a & ActiveCell.Value
                column = column + 1
            Loop
        Case "A"
            a = ActiveCell.Value
            Do While ActiveCell <> ""
                ActiveCell.Offset(0, 1).Select
                a = a & ActiveCell.Value
                column = column + 1
            Loop
        Case "B"
            a = ActiveCell.Value
            Do While ActiveCell <> ""
                ActiveCell.Offset(0, 1).Select
                a = a & ActiveCell.Value
                column = column + 1
            Loop
        Case "C"
            a = ActiveCell.Value
            Do While ActiveCell <> ""
                ActiveCell.Offset(0, 1).Select
                a = a & ActiveCell.Value
                column = column + 1
            Loop
        Case "K"
            a = ActiveCell.Value
            Do While ActiveCell <> ""
                ActiveCell.Offset(0, 1).Select
                a = a & ActiveCell.Value
                column = column + 1
            Loop
        Case "F"
            a = ActiveCell.Value
            Do While ActiveCell <> ""
                ActiveCell.Offset(0, 1).Select
                a = a & ActiveCell.Value
                column = column + 1
            Loop
        Case Else
            a = "Invalid record type"
        End Select
        
        
        Print #1, a
        
        'ActiveCell.Offset(1, column * -1).Select
        row = row + 1
        Cells(row, 1).Select
        column = 1
        
    Loop
    
    ActiveCell.Offset(-1, 1).Select
    a = a & ActiveCell.Value
    Print #1, a
    
    Print #1, row & " " & column
    Print #1, RowToFile(row, column)
    
    Close
    
    

End Sub


Function RowToFile(row As Long, column As Long)
    If row = 0 Or column = 0 Then
        RowToFile = 0
        Exit Function
    End If
    
    RowToFile = Cells(row, column).Value
End Function

