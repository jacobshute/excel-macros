Attribute VB_Name = "KanBan"
' Everything works in the table A:E
' columns: non-work upcoming | work upcoming | in progress | done | archive

' TODO make this work with a range and not just a single cell
' TODO if it is going into the done or archive cell, put at top instead of bottom
' Moves a cell right
Sub KanBan_right()
    
    If ActiveSheet.Name <> "KanBan" And ActiveSheet.Name <> "KanBan TEST" Then
        Debug.Print ("working in the wrong sheet!!!!")
        Exit Sub
    End If
    
    Dim task_table As Range, column As Range
    Set task_table = ActiveSheet.Range("A2:D2", ActiveSheet.Range("D" & ActiveSheet.Rows.Count).End(xlUp))
    
    ' quit if not on a valid cell
    If Selection.row = 1 Then
        Debug.Print ("tried to move header row")
        Exit Sub
    End If
    If Selection.Value = "" Then
        Debug.Print ("tried to move empty cell")
        Exit Sub
    End If
        
    ' look at column to the right, iterate until empty cell is found, then cut and paste
    For i = 2 To task_table.Rows().Count
        If Cells(i, Selection.column + 1) = "" Then
            Cells(i, Selection.column + 1).Value = ActiveCell.Value
            ActiveCell.Value = ""
            Set column = task_table.Columns(Selection.column)
            KanBan_shift_up column
            KanBan_in_progress
            Exit Sub
        End If
    Next i

End Sub


' Moves a cell left
Sub KanBan_left()
    Dim task_table As Range, column As Range
    Set task_table = ActiveSheet.Range("A2:D2", ActiveSheet.Range("D" & ActiveSheet.Rows.Count).End(xlUp))
    
    ' quit if not on a valid cell
    If Selection.row = 1 Then
        Debug.Print ("tried to move header row")
        Exit Sub
    End If
    If Selection.Value = "" Then
        Debug.Print ("tried to move empty cell")
        Exit Sub
    End If
    
    ' look at column to the left, iterate until empty cell is found, then cut and paste
    For i = 2 To task_table.Rows().Count
        If Cells(i, Selection.column - 1) = "" Then
            Cells(i, Selection.column - 1).Value = ActiveCell.Value
            ActiveCell.Value = ""
            Set column = task_table.Columns(Selection.column)
            KanBan_shift_up column
            KanBan_in_progress
            Exit Sub
        End If
    Next i
    
End Sub


' TODO create a check and warning for too many things in the "in progress" tab
Sub KanBan_in_progress()
    Dim task_table As Range
    Set task_table = ActiveSheet.Range("C2", ActiveSheet.Range("C" & ActiveSheet.Rows.Count).End(xlUp))
    
    For Each Item In task_table
        If Item <> "" Then
            i = i + 1
        End If
    Next Item
    
    If i > 4 Then
        MsgBox ("there are too many items in progress!")
    End If
End Sub



' look for empty cells and rotate up
Sub KanBan_shift_up(column As Object)
    ' iterators i and j       row reference number iterator
    Dim i As Integer, j As Integer, row As Integer
    i = 1
    j = 1
    
    ' cell data
    Dim cell As Range, scratch As Range
    Set cell = column.Cells(1)
    Set scratch = column.Cells(1)
    
    ' for every cell, go through and rotate up
    Do While i <= column.Rows.Count
        
        Set cell = column.Cells(i)
        If cell.Value <> "" Then
            ' go through all cells in the column looking for a blank spot
            ' TODO make this actually efficient... maybe store a var that is "highest open spot" or something. Then solve how to adjust that once it has been assigned?
            j = 1
            Do While j < i
                Set scratch = column.Cells(j)
                If scratch.Value = "" Then
                    scratch.Value = cell.Value
                    cell.Value = ""
                    Exit Do
                End If
                j = j + 1
            Loop
        End If
        
        i = i + 1
    Loop
    
End Sub

Sub KanBan_repeat()
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
''''''''''
'
' Messing with sorting algorithms here.
' I am just going to impliment a lot of them to try them out
'
Sub KanBan_Sort()
    ' for benchmarking
    Dim time As Double
    time = GetTickCount
    Dim sw As Stopwatch
    Set sw = New Stopwatch
    sw.StartTimer
    
    Dim sheet As Worksheet
    
    For i = 1 To Worksheets.Count
    If Worksheets(i).Name = "KanBan_scratch" Then
        exists = True
        Set scratch_sheet = Worksheets(i)
    End If
    Next i

    If Not exists Then
        Set scratch_sheet = Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
        scratch_sheet.Name = "KanBan_scratch"
    End If
            
    
    Dim task_table As Range, column As Range, work_table As Range
    Set task_table = ActiveSheet.Range("A2:E2", ActiveSheet.Range("E" & ActiveSheet.Rows.Count).End(xlUp))
    Set work_table = scratch_sheet.Range("A2:A2", scratch_sheet.Range("A" & task_table.Rows.Count))
    For Each cell In work_table
        cell.Value = ""
    Next cell
    
    
    For Each column In task_table.Columns
        Debug.Print ("---------- Column " & column.column & " ----------")
        KanBan_shift_up column
        KanBan_merge_sort column
        'KanBan_bubble_sort column
        'KanBan_rough_sort column
    Next
    
    Debug.Print ("Total time to execute: " & sw.EndTimer) '& Application.Text(sw.GetTickCount - time, "mm:ss.000"))

End Sub


' This one is just doing replacement and only looks at HIGH MEDIUM and LOW, it does not do alphabetical
' TODO I don't know if this should be working. Maybe check through it to make sure it isn't broken...
Sub KanBan_rough_sort(column As Range)

    ' loop iterators
    Dim i As Integer, j As Integer
    ' location high, medium, and low cells -- not the last one, the next open one
    Dim high_ptr As Integer, med_ptr As Integer, low_ptr As Integer

    ' for swapping cells
    Dim scratch As String
        
    high_ptr = 1
    low_ptr = 1
    med_ptr = 1

    ' LOOP through CELLS in the column and sort
    j = 1
    Do While j <= column.Rows.Count
        Set cell = column.Cells(j)
            
        If InStr(cell.Value, "H:") <> 0 Then
                
            If cell.row = high_ptr + 1 Then
                ' don't move cell, move on if the cell is already where the pointer would be
                ' don't need to incriment because .cells is 0 based and .row is 1 based. Thanks microsoft.
                j = j + 1
            Else
                ' swap the cell with the pointer cell
                scratch = cell.Value
                cell.Value = column.Cells(high_ptr).Value
                column.Cells(high_ptr).Value = scratch
                    
                ' make sure lower pointers aren't behind higher ones
                high_ptr = high_ptr + 1
                If med_ptr < high_ptr Then
                    med_ptr = high_ptr
                End If
                If low_ptr < high_ptr Then
                    low_ptr = high_ptr
                End If
                    
            End If
            Debug.Print (cell.Value)
        ElseIf InStr(cell.Value, "M:") <> 0 Then
                
            If cell.row = med_ptr + 1 Then
                ' cell is where it belongs, move on
                ' don't increment because 0 based and 1 based arrays are hard.
                j = j + 1
                    
            Else
                ' swap the cell with the pointer cell
                scratch = cell.Value
                cell.Value = column.Cells(med_ptr).Value
                column.Cells(med_ptr).Value = scratch
                    
                ' make sure lower pointers aren't behind higher ones
                med_ptr = med_ptr + 1
                If low_ptr < med_ptr Then
                    low_ptr = med_ptr
                End If
            End If
                
            Debug.Print (cell.Value)
        ElseIf InStr(cell.Value, "L:") <> 0 Then
                
            If cell.row = low_ptr + 1 Then
                ' cell is where it belongs, move on
                ' don't increment because 0 based and 1 based arrays are hard.
                j = j + 1
                    
            Else
                ' swap the cell with the pointer cell
                scratch = cell.Value
                cell.Value = column.Cells(low_ptr).Value
                column.Cells(low_ptr).Value = scratch
                    
                low_ptr = low_ptr + 1
                    
            End If
                
            Debug.Print (cell.Value)
        Else
            ' if there's nothing here, or there's no marker, just leave it and move on to the next cell.
            j = j + 1
        End If
            
    Loop
    
End Sub

Sub KanBan_bubble_sort(column As Range)

    ' loop iterators
    Dim i As Integer, j As Integer, cell_count As Integer
    Dim swapped As Boolean




    j = 0
    swapped = True

    Do While swapped <> False
        i = 1
        
        swapped = False
        Set cell = column.Cells(i)
        Set next_cell = column.Cells(i + 1)
        cell_count = column.Cells.Count + 1
            
        Do While i < cell_count - j And next_cell <> ""
            
            If check_swap(cell, next_cell) Then
                swapped = True
            End If
            
            Set cell = column.Cells(i)
            Set next_cell = column.Cells(i + 1)
            i = i + 1
            
        Loop
        
        j = j + 1
    Loop


End Sub

'''''''''''''''''''''''''''''''''''''''
' Merge sort
' https://en.wikipedia.org/wiki/Merge_sort
' doing the top down one
' takes 76844 milliseconds to run on full test set

Sub KanBan_merge_sort(List As Range)

    Dim n As Integer
    n = List.Rows.Count
    
    If n <= 1 Then
        Exit Sub
    End If
    
    Dim A As Range, B As Range
    Set A = List.Parent.Range(List.Cells(1), List.Cells(n / 2))
    Set B = List.Parent.Range(List.Cells((n / 2) + 1), List.Cells(n))

    KanBan_merge_sort A
    KanBan_merge_sort B
    
    Dim scratch As String
    Dim i As Integer, j As Integer
    i = 1
    j = 1
    Do While i <= A.Rows.Count
' this works actually...
'        If B(i) = "" Then
'            'do nothing
'        ElseIf B(i) < cell Then
'            scratch = cell.Value
'            cell.Value = B(i).Value
'            B(i).Value = scratch
'            i = i + 1
'        End If
        
        If check_swap(A(i), B(j)) Then
            j = j + 1
        Else
            i = i + 1
        End If
        
    Loop
    
End Sub




' TODO - not actually swapping alphabetically for some reason...
' works with bubble sort, not with merge sort.

' was copy-pasting this in different sort algos so made a function
' looks at the two ranges and swaps in correct priority then alphabetical order
Function check_swap(A, B)
    check_swap = False
    If B = "" Then
        Exit Function
    End If
    If mid(A, 1, 2) = mid(B, 1, 2) Then
        If StrComp(Replace(A, " ", ""), Replace(B, " ", ""), vbTextCompare) = 1 Then
            check_swap = True
            scratch = A.Value
            A.Value = B.Value
            B.Value = scratch
        End If
    ElseIf mid(A, 1, 2) = "H:" Or mid(A, 1, 2) = "M:" Or mid(A, 1, 2) = "L:" Then
        If mid(A, 1, 2) = "H:" Or mid(B, 1, 2) = "L:" Then
            ' As in right order
        ElseIf mid(B, 1, 2) <> "H:" And mid(B, 1, 2) <> "M:" And mid(B, 1, 2) <> "L:" Then
            ' in right place
        ElseIf mid(A, 1, 2) = "L:" Then
            check_swap = True
            scratch = A.Value
            A.Value = B.Value
            B.Value = scratch
        ElseIf mid(B, 1, 2) = "M:" Then
            check_swap = True
            scratch = A.Value
            A.Value = B.Value
            B.Value = scratch
        ElseIf mid(B, 1, 2) = "H:" Then
            check_swap = True
            scratch = A.Value
            A.Value = B.Value
            B.Value = scratch
        End If
    ElseIf mid(B, 1, 2) = "H:" Or mid(B, 1, 2) = "M:" Or mid(B, 1, 2) = "L:" Then
        check_swap = True
        scratch = A.Value
        A.Value = B.Value
        B.Value = scratch
    ElseIf StrComp(Replace(A, " ", ""), Replace(B, " ", ""), vbTextCompare) = 1 Then
        check_swap = True
        scratch = A.Value
        A.Value = B.Value
        B.Value = scratch
    End If
    
End Function
