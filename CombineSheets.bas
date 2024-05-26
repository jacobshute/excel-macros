Attribute VB_Name = "CombineSheets"
' Combines sheets selected in a form into a new sheet
' PowerQuery might be a better tool for this: autorefresh, no macros, error handling, etc.
Sub CombineSheets()
Application.ScreenUpdating = False
    
    Dim Sheets() As Variant
    Dim i As Long
    
    ' Get list of sheet names
    ReDim Sheets(ActiveWorkbook.Sheets.Count - 1)
    i = 0
    For Each sheet In ActiveWorkbook.Sheets
        Sheets(i) = sheet.Name
        i = i + 1
    Next sheet

    ' Load the form for the user to select the sheets to combine
    Load CombineSheetsForm
    With CombineSheetsForm
        .SheetList.List = Sheets
    End With
    CombineSheetsForm.Show

Application.ScreenUpdating = True
End Sub

