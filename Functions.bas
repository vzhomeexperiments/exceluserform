Attribute VB_Name = "Functions"
        ' (C) 2017 VZ Home Experiments Vladimir Zhbanko //vz.home.experiments@gmail.com
        ' VBA code to make work with Excel User Forms easier
        ' More time to spend on more interesting stuff.

'========================================
' FUNCTION that keep First available Capital letter in the string
'========================================
Function getFirstCapitalLetter(myInput As String) As String
    ' Declaring variables
    Dim myResult As String    ' This is the return string
    Dim i As Long             ' Counter for character position

    ' Initialise return string to empty
    myResult = ""

    ' For every character in input string, copy digits to
    '   return string if they are passing criteria
    For i = 1 To Len(myInput)
        If Mid(myInput, i, 1) >= "A" And Mid(myInput, i, 1) <= "Z" Then
            myResult = myResult + Mid(myInput, i, 1)
            Exit For
        End If
    Next

    ' Then return the return string.                          '
    getFirstCapitalLetter = myResult
End Function

'========================================
' FUNCTION that keep All available Capital letters in the string
'========================================
Function getAllCapitalLetters(myInput As String) As String
    ' Declaring variables
    Dim myResult As String    ' This is the return string
    Dim i As Long             ' Counter for character position

    ' Initialise return string to empty
    myResult = ""

    ' For every character in input string, copy digits to
    '   return string if they are passing criteria
    For i = 1 To Len(myInput)
        If Mid(myInput, i, 1) >= "A" And Mid(myInput, i, 1) <= "Z" Then
            myResult = myResult + Mid(myInput, i, 1)
        End If
    Next

    ' Then return the return string.                          '
    getAllCapitalLetters = myResult
End Function
'========================================
' FUNCTION that removes all text from string, and leave only numbers
'========================================
Function getOnlyDigit(myInput As String) As String
    ' Declaring variables
    Dim myResult As String  ' This is the return string
    Dim i As Long           ' Counter for character position

    ' Initialise return string to empty
    myResult = ""

    ' For every character in input string, copy digit to
    '   return string if they are passing criteria
    For i = 1 To Len(myInput)
        If Mid(myInput, i, 1) >= "0" And Mid(myInput, i, 1) <= "9" Then
            myResult = myResult + Mid(myInput, i, 1)
            Exit For
        End If
    Next

    ' Then return the return string.                          '
    getOnlyDigit = myResult
End Function
'========================================
' FUNCTION that tells if string contains digits
'========================================
' function is adapted using function getOnlyDigits
Function isDigit(myInput As String) As Boolean
    ' Variables needed (remember to use "option explicit")
    Dim myResult As Boolean ' This is the return boolean
    Dim i As Integer        ' Counter for character position

    ' Initialise return result to be false
    myResult = False

    ' For every character in input string, check if there are
    ' numbers. Stop if found at least one number
    For i = 1 To Len(myInput)
        If Mid(myInput, i, 1) >= "0" And Mid(myInput, i, 1) <= "9" Then
            myResult = True
            Exit For
        Else
            myResult = False
        End If
    Next

    ' Then return the results
    isDigit = myResult
End Function
'========================================
' FUNCTION that count cell color in a range
'========================================
' This is a user defined function! UDF!
Function CountCellColor(range_data As Range, criteria As Range) As Long
    Dim datax As Range
    Dim xcolor As Long
xcolor = criteria.Interior.ColorIndex
For Each datax In range_data
    If datax.Interior.ColorIndex = xcolor Then
        CountCcolor = CountCcolor + 1
    End If
Next datax
End Function
