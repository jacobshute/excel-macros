Attribute VB_Name = "Lispy"

'
' TODO: figure out how to get all the operators and such into another function
' can't use globals...
' Could do GoTo statement and make everyone die inside?
'
' String manipulator functions (hard to find...)
' https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/string-manipulation-keyword-summary
'


Sub lispy()

    ' following buildyourownlisp.com tutorial and translating to VBA (:
    
    Debug.Print ("Start")
    
    Dim expression As String
    expression = ActiveCell.Value
    
    ' check if the expression is valid before calling eval()
    If Mid(expression, 1, 1) <> "(" Then
        MsgBox ("Expressions must start with a " & Chr(34) & "(" & Chr(34))
        Exit Sub
    End If
    
    ActiveCell.Offset(0, 1).Value = eval(expression)
    

End Sub


Function eval(expression As String)

    
    Dim term As String, operator As String, t As String, i As Long, j As Long, done As Boolean, para_count As Integer
    term = ""
    operator = ""
    para_count = 0
    
    
    ' trying to use a worksheet for the operations list instead of a variable
    ' the worksheet that the Lisp operations are stored in
    Dim lispy_sheet As String
    lispy_sheet = "LISPY_DATA"
    
    Dim main_sheet As Variant
    main_sheet = ActiveSheet.Name
    
    Worksheets(lispy_sheet).Activate
    
    
    ' move this out of this function so that the function can be recursive.
    If Mid(expression, 1, 1) <> "(" Then
        Debug.Print ("beginning of expression")
    End If
    
    ' get the operator
    done = False
    i = 1
    Do While i < Len(expression) And done = False
        t = Mid(expression, i, 1)
        
        ' TODO: This is just looking for the operator, not the term,
        ' so an open paentheses is not valid. Need to check for that when looking for terms
        
        ' checking for the end of the operator
        If t = " " Then
            If operator = "" Then
                Debug.Print (t)
            Else
                done = True
            End If
        ElseIf t = "(" Then
            para_count = para_count + 1
            Debug.Print (para_count)
            Debug.Print (t)
        ElseIf t = ")" Then
            para_count = para_count - 1
            Debug.Print (para_count)
            Debug.Print (t)
        Else
            operator = operator + t
            Debug.Print (operator)
        End If
        
        
        i = i + 1
    Loop
    
    ' check operator against valid operators
    Cells(2, 1).Select
    done = False
    Do While ActiveCell.Value <> "" And done = False
            
        If operator = ActiveCell.Value Then
            Debug.Print (operator & " is the operator")
            done = True
        End If
            
            
        Debug.Print ("works")
        ActiveCell.Offset(1, 0).Select
        
    Loop
        
        ' checking for end of expression
        If Mid(expresssion, i, 1) = ")" Then
            Debug.Print ("space was here")
        End If
        
        
    
    ' output the final term that the expression evaluated to
    eval = term
    
    Worksheets(main_sheet).Activate
    
End Function
