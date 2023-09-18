Attribute VB_Name = "Shortcuts"


Sub FilesInDirectory()
    '
    ' takes in a path from a cell and outputs the files in the folder into the cells below
    
    Dim fs As Object
    Dim folder As Object
    Dim files As Object
    
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    If fs.FolderExists(ActiveCell.Value) Then
        Set folder = fs.GetFolder(ActiveCell.Value)
        Set files = folder.files
        
        If files.Count = 0 Then
            MsgBox ("No files in folder")
            Exit Sub
        End If
        
        For Each file In files
            ActiveCell.Offset(1, 0).Select
            ActiveCell.Value = file.Name
        Next
        
    Else
        MsgBox ("Folder does not exist")
    End If
    
End Sub

Sub add_row()
Attribute add_row.VB_ProcData.VB_Invoke_Func = "e\n14"
'
' insert a row
'
' Keyboard Shortcut: Ctrl+e
'
    Selection.EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
End Sub

Sub add_cell()
Attribute add_cell.VB_ProcData.VB_Invoke_Func = "w\n14"
'
' inserts a single cell
'
' Keyboard Shortcut: Ctrl+w
'
' TODO make this work in tables? right now it throws an error
'
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
End Sub

Sub delete_cell()
Attribute delete_cell.VB_ProcData.VB_Invoke_Func = "W\n14"
'
' delete cell, shift up
'
' Keyboard Shortcut: Ctrl+Shift+W
'
    Selection.Delete Shift:=xlUp
End Sub

Sub delete_row()
Attribute delete_row.VB_ProcData.VB_Invoke_Func = "E\n14"
'
' delete row, shift up
'
' Keyboard Shortcut: Ctrl+Shift+R
'
    Selection.EntireRow.Delete Shift:=xlUp
End Sub

Sub Clear()
Attribute Clear.VB_ProcData.VB_Invoke_Func = "D\n14"
'
' Keyboard Shortcut: Ctrl+Shift+D
'
    Selection.Clear
End Sub

Sub highlight()
Attribute highlight.VB_ProcData.VB_Invoke_Func = "h\n14"
'
' toggle blue highlight on a selection
'
' Keyboard Shortcut: Ctrl+h
'
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

Sub row_size_20()
Attribute row_size_20.VB_ProcData.VB_Invoke_Func = "Q\n14"
'
' make the row height 20 pixels
'
' Keyboard Shortcut: Ctrl+Shift+Q
'
    Selection.RowHeight = 15
End Sub

Sub RandomWord()
Attribute RandomWord.VB_ProcData.VB_Invoke_Func = "q\n14"
    Dim i As Long, l As Long, B As Long, c As Long, ret As String
    
    For Each cell In Selection.Cells
    ret = ""
    
    l = Int((20 * Rnd) + 5)
    B = Int((4 * Rnd) + 1)
    Debug.Print (l & " : " & B & " : " & c)
    If B = 1 Then
        ret = ret & "L:"
    ElseIf B = 2 Then
        ret = ret & "M:"
    ElseIf B = 3 Then
        ret = ret & "H:"
    End If
        
    
    
    
    For i = 1 To l
        c = Int((4 * Rnd) + 1)
        If c > 1 Then
            ret = ret & Chr(Rnd * (Asc("z") - Asc("a") + 1) + Asc("a") - 1)
        Else
            ret = ret & Chr(32)
        End If
    Next
    cell.Value = ret
    
    Next
End Sub
