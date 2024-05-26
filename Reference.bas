Attribute VB_Name = "Reference"
' Reference for VBA in Excel
Sub Reference()              ' Will throw an error on other code if this is not in a procedure.

                                       ' Variables and Types
' Microsoft help page:
' https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/data-type-summary

' Dim declares a variable name
' As declares the type
' Const declares a constant





Dim var As Variant           ' 16 bytes for numbers, 22 (24 on 64 bit system) for characters.
                             ' stores many kinds of values - can be slow because it does type checks
                             ' is the default for VBA. Very versetile

' Number types
Dim var As Byte              ' 8 bit int
Dim var As Integer           ' 16 bit int
Dim var As Long              ' 32 bit int
Dim var As LongLong          ' 64 bit int - not supported in Excel VBA 32 bit?
Dim var As Single            ' 32 bit float
Dim var As Double            ' 64 bit float


' Other types
Dim var As Date              ' 32 bit date stored as a signed int.
Dim var As Object            ' stores an object
Dim var As String            ' variable length string. Use (#) for fixed length
Dim var As Boolean           ' 16 bit - stores True or False values
Dim var As Currency          ' 64 bit currency value

' pointers
' there isn't much reason to declare pointers
' I think that they are used under the hood by "ByRef" in function parameters
' Also used when making DLL calls
Dim var As LongPtr           ' 32 or 64 bit depending on system


' User defined types
' These act similar to C structures
' Lets you return multiple values from functions
' Faster than using a class
Type MyType
    Name As String
    Age As Integer
End Type

Dim var As MyType
var.Name = "Bob"
var.Age = 23


' Setting values
' normal vlaues are set with a single =
var = 1
' objects are set using the "Set" keyword
Set var = ActiveSheet



' Arrays
' 0 indexed
' fixed size after initialization, can reinitialize to new size
Dim arr(10) As String        ' declare an array

ReDim arr(20)                ' resize the array


' Casting/Converting and Checking Types
' returns number or string expression as a specific type
var = CBool(expression)
var = CByte(expression)
var = CCur(expression)
var = CDate(expression)
var = CDbl(expression)
var = CDec(expression)
var = CInt(expression)
var = CLng(expression)
var = var = CLngLng(expression) '(Valid on 64-bit platforms only.)
var = CLngPtr(expression)
var = CSng(expression)
var = CStr(expression)
var = CVar(expression)

' Returns True or False, depending on if the type is a match
IsArray (var)
IsDate (var)
IsEmpty (var)
IsError (var)
IsMissing (var)
IsNull (var)
IsNumeric (var)
IsObject (var)


                                       ' Comparisons and operators
' making comparisons
' when the expression i evaluated, it will return True or False
' Variant types have strange interaction with this if the two expressions are not the same underlying type
var = 1                      ' "=" is used for both setting a variable and for comparing values
                             ' VB will handle it depending on context
var <> 1                     ' not equal to
var > 1                      ' greater than
var < 1                      ' less than
var >= 1                     ' greater than or equal
var <= 1                     ' less than or equal

' logical operations
expression1 And expression2  ' Returns True only if both expressions are True
expression1 Or expression2   ' returns True if either, or both expressions are True
expression1 Xor expressin2   ' exclusive or - true only if both expressions are different
Not expression1              ' Returns False if the expression is True and vice versa
expression1 Imp expression2  ' bitwise comparison of numeric expressions
expression1 Eqv expression2  ' bitwise comparison that sets the result to 1 if the same and 0 if different






                                       ' Conditionals
' If Else
If Condition Then
    ' do thing
ElseIf Condition Then
    ' do other thing
Else
    ' do another thing
End If


' Switch/Select Statement
' runs only the ossociated case where the value of the input matches
Select Case var
Case 1 To 5
    ' do thing
Case 6, 7, 8
    ' do other thing
Case Else
    ' do another thing
End Select



                                       ' Loops and Control flow

' Do While and Until
Do While x > 10              ' loops while x not equal to 10
    ' do thing
Loop

Do
    ' do thing
Loop While x > 10            ' does the conditional check after the first run


Do Until x = 10              ' loops until the x is 10
    ' do thing               ' Until can also be put after loop similar to While
Loop


' For loops
For x = 1 To 10 Step 2       ' iterate over x from value 1 to 10 in steps of 2
                             ' default, without step keyword, is steps of 1. Steps can be negative
    ' do thing               ' x does not need to be declared before hand
Next


For Each l In List           ' loop over each item "l" in an array or object "List"
    ' do thing
Next l                       ' the next keyword can be used on its own, or with the iterator after


' Loops can be exited using the following statements:
Exit Do
Exit For


                                       ' Functions and Subroutines
' commented out because it was causing compile issues in another file?

'' Subroutines
'' can take parameters
'' do not return values
'Sub func()
'    ' do thing
'End Sub
'
'' Functions
'' takes parameters
'' returns a value
'Function func()
'    ' do thing
'    func = 1                 ' The name of the function is the variable that will be returned
'End Function
'
'
'' Passing variables
'' Variables are passed by reference by default
'Sub func(ByRef var As Integer)    ' Passes by reference
'End Sub
'
'
'Sub func(ByVal var As Integer)    ' Passes by Value
'End Sub




                                       ' With Statements
' with statements allow for shorthand when interacting with an object
' mildly redundant in most situations, but useful when accessing many object members
With ActiveSheet
    .Name = "new name"
End With





                                       ' Directives
                                       
' Directives are like C pre-processor directives

' Constants
#Const con = 1

' If Else
' conditional depending on compiler constants (see above)
#If thing Then
    ' only if "thing" declared
#ElseIf otherThing Then
    ' only if "otherThing" declared
#Else
    ' otherwise, do this
#End If


                                       ' Excel Samples

         ' Ranges
' Ranges are the main data structure/type in excel for working with multiple cells
' There are many properties and methods
' Ranges can be created without a reference to the sheet, but they are typically passed
' by reference and associated with a set of cells in the sheet.




' Cells are stored in the range.Cells(x,y) array



         ' Cells
' Cells are where all of a cells properties are stored
' Many properties and functions




End Sub

