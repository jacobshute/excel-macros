Attribute VB_Name = "Macros"
Sub FilesInDirectory()

    '

    ' This macro takes in a path and outputs the files in the folder into the cells below

   

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

Sub highlight()

'

' highlight Macro

' highlight or unhighlight a selection

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




Sub FilesInDirectory()
    '
    ' This macro takes in a path and outputs the files in the folder into the cells below
    
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
'
' new_row Macro
'
' Keyboard Shortcut: Ctrl+e
'
    Selection.EntireRow.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
End Sub

Sub add_cell()
'
' add_partial_row Macro
'
' Keyboard Shortcut: Ctrl+w
'
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
End Sub

Sub Clear()

'
' Clear Macro
' Clear all
'
' Keyboard Shortcut: Ctrl+Shift+D
'
    Selection.Clear
End Sub

Sub delete_cells()

'
' delete_cells Macro
' delete cells, shift up
'
' Keyboard Shortcut: Ctrl+Shift+W
'
    Selection.Delete Shift:=xlUp
End Sub

Sub delete_row()

'
' delete_row Macro
' delete row, shift up
'
' Keyboard Shortcut: Ctrl+Shift+R
'
    Selection.EntireRow.Delete Shift:=xlUp
End Sub

