Attribute VB_Name = "Shortcuts"

' takes in a path from a cell and outputs the files in the folder into the cells below
Sub FilesInDirectory()
    
    Dim fs As Object
    Dim folder As Object
    Dim files As Object
    
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    If fs.FolderExists(ActiveCell.value) Then
        Set folder = fs.GetFolder(ActiveCell.value)
        Set files = folder.files
        
        If files.Count = 0 Then
            MsgBox ("No files in folder")
            Exit Sub
        End If
        
        For Each file In files
            ActiveCell.Offset(1, 0).Select
            ActiveCell.value = file.Name
        Next
        
    Else
        MsgBox ("Folder does not exist")
    End If
    
End Sub

' insert a row, shifts down
Sub add_row()
Attribute add_row.VB_ProcData.VB_Invoke_Func = "e\n14"
    Selection.EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
End Sub

' inserts a single cell, shifts down
' TODO make this work in tables? right now it throws an error
Sub add_cell()
Attribute add_cell.VB_ProcData.VB_Invoke_Func = "w\n14"
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
End Sub

' delete cell, shift up
Sub delete_cell()
Attribute delete_cell.VB_ProcData.VB_Invoke_Func = "W\n14"
    Selection.Delete Shift:=xlUp
End Sub

' delete row, shift up
Sub delete_row()
Attribute delete_row.VB_ProcData.VB_Invoke_Func = "E\n14"
    Selection.EntireRow.Delete Shift:=xlUp
End Sub

' Clear all contents and formatting
Sub Clear()
Attribute Clear.VB_ProcData.VB_Invoke_Func = "D\n14"
    Selection.Clear
End Sub

' toggle blue highlight on a selection
Sub highlight()
Attribute highlight.VB_ProcData.VB_Invoke_Func = "h\n14"

    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .TintAndShade = 0
        .PatternTintAndShade = 0
        
        If .ThemeColor = xlThemeColorAccent5 Then
            Debug.Print ("is highlighted" & .ThemeColor)
            .ThemeColor = xlNone
        Else
            Debug.Print ("is not highlighted" & .ThemeColor)
                    .ThemeColor = xlThemeColorAccent5
        End If

    End With
End Sub

' set the row height of the selection to 15
Sub rowHeight15()
Attribute rowHeight15.VB_ProcData.VB_Invoke_Func = "Q\n14"
    Selection.RowHeight = 15
End Sub

' Fills selection with random characters
' for testing large random sets of data
' Separates characters into words within the cell randomly to emulate sentences
Sub randomWords()
Attribute randomWords.VB_Description = "Generate Random tast data for KanBan project"
Attribute randomWords.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim i As Long, length As Long, spaceChance As Long, spaceCount As Long, value As String
    
    For Each cell In Selection.Cells
        value = ""
        spaceCount = 0
        
        ' length in letters
        length = Int((20 * Rnd) + 5)

        ' loop for the number of letters, filling a random ASCII character vlaue
        ' Adds a 1 in 4 chance for a space instead of a letter
        For i = 1 To length
            spaceChance = Int((4 * Rnd) + 1)
            If spaceChance > 1 Then
                value = value & Chr(Rnd * (Asc("z") - Asc("a") + 1) + Asc("a") - 1)
            Else
                value = value & Chr(32)
                spaceCount = spaceCount + 1
            End If
        Next
        cell.value = value
        
        ' Print number of letters and spaces and the value
        Debug.Print ("length: " & length & "  number of spaces: " & spaceCount & " value: " & value)
        
    Next
End Sub

' Takes the value of a cell and turns it into a formula for that value
Sub valueToFormula()
    Dim value As String
    
    value = ActiveCell.value
    ' add equals (ASCII 61) and wrap in quotes (ASCII 34)
    ActiveCell.value = Chr(61) & Chr(34) & value & Chr(34)
    
End Sub

' converts a cell formula output to a value
Sub formulaToValue()
    ActiveCell.value = ActiveCell.value
End Sub


' Loops through selection and copies the contents to the cell below as plain text
Sub copyContentsDown()
    
    Dim r As Range
    Set r = Selection.Cells

    For Each cell In r
        cell.Offset(1, 0).Select
        ActiveCell.value = cell.value
    Next cell

End Sub


