VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CombineSheetsForm 
   Caption         =   "Combine Sheets"
   ClientHeight    =   5580
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8295
   OleObjectBlob   =   "CombineSheetsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CombineSheetsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CancelButton_Click()
    Unload Me
End Sub

' Create new merge sheet and copy contents of selected sheets into it
Private Sub OkButton_Click()
    
    Dim sheet As Worksheet
    Dim hasSelection As Boolean
    
    Dim rowOffset As Long, lastRow As Long, numberOfMergedSheets As Long
    rowOffset = 0
    lastRow = 1
    numberOfMergedSheets = 0
    
    ' Exit sub if nothing was selected
    For x = 0 To SheetList.ListCount - 1
        If SheetList.Selected(x) Then
            hasSelection = True
        End If
    Next x
    If Not hasSelection Then
        MsgBox ("No sheets selected")
        Exit Sub
    End If
    
    ' If the merge sheet already exists, alert the user and exit sub
    For Each sheet In ActiveWorkbook.Sheets
        If sheet.Name = mergeSheetNameText.value Then
            MsgBox ("A sheet with that name already exists. Please choose a new name for the merge sheet")
            Exit Sub
        End If
    Next
    
    ' Create the sheet to merge into
    Set mergeSheet = ActiveWorkbook.Sheets.Add(after:=Worksheets(Worksheets.Count))
    ActiveWorkbook.Sheets(mergeSheet.Index).Name = mergeSheetNameText.value
  
    ' header row
    If HeaderCheckBox.value Then
        rowOffset = 1
        mergeSheet.Cells(1, 1).value = "Sheets Merged:"
    End If

    ' Loop through the sheets and merge them
    For x = 0 To SheetList.ListCount - 1
        If SheetList.Selected(x) Then
            Debug.Print (SheetList.List(x))
            
            Set sheet = ActiveWorkbook.Sheets(SheetList.List(x))
            
            If HeaderCheckBox.value Then
                mergeSheet.Cells(1, numberOfMergedSheets + 2).value = sheet.Name
            End If
            
            numberOfMergedSheets = numberOfMergedSheets + 1
            
            Debug.Print ("Merging: " & sheet.Name & " | at index: " & sheet.Index)
            For Each cell In sheet.UsedRange.Cells
                mergeSheet.Cells(cell.row + rowOffset, cell.column) = cell
                lastRow = cell.row
            Next cell
            
            ' get the last row number of the sheet and add it to the offset
            ' this is where the next sheet will start copying into the merge sheet
            rowOffset = rowOffset + lastRow
            
        End If
    Next x
    
    MsgBox ("Sheets Merged into " & mergeSheetNameText.value & " sheet")
    Unload Me
End Sub

Private Sub SelectAllSheets_Click()
    For x = 0 To SheetList.ListCount - 1
        SheetList.Selected(x) = True
    Next x
End Sub

Private Sub DeselectAllSheets_Click()
    For x = 0 To SheetList.ListCount - 1
        SheetList.Selected(x) = False
    Next x
End Sub

Private Sub HeaderCheckBoxLabel_Click()
    If HeaderCheckBox.value Then
        HeaderCheckBox.value = False
    Else
        HeaderCheckBox.value = True
    End If
End Sub
