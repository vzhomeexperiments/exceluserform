Attribute VB_Name = "Functions"
        ' (C) 2017 VZ Home Experiments Vladimir Zhbanko //vz.home.experiments@gmail.com
        ' VBA code to make work with Excel User Forms easier
        ' More time to spend on more interesting stuff.

'========================================
' FUNCTION that keep First available Capital letter in the string
'========================================
' function taken somewhere from Stackoverflow
Function getOnlyCapitalLetter(s As String) As String
    ' Variables needed (remember to use "option explicit").   '
    Dim retval As String    ' This is the return string.      '
    Dim i As Integer        ' Counter for character position. '

    ' Initialise return string to empty                       '
    retval = ""

    ' For every character in input string, copy digits to     '
    '   return string.                                        '
    For i = 1 To Len(s)
        If Mid(s, i, 1) >= "A" And Mid(s, i, 1) <= "Z" Then
            retval = retval + Mid(s, i, 1)
            Exit For
        End If
    Next

    ' Then return the return string.                          '
    getOnlyCapitalLetter = retval
End Function

'========================================
' FUNCTION that keep All available Capital letters in the string
'========================================
' function taken somewhere from Stackoverflow
Function getOnlyCapitalLetters(s As String) As String
    ' Variables needed (remember to use "option explicit").   '
    Dim retval As String    ' This is the return string.      '
    Dim i As Integer        ' Counter for character position. '

    ' Initialise return string to empty                       '
    retval = ""

    ' For every character in input string, copy digits to     '
    '   return string.                                        '
    For i = 1 To Len(s)
        If Mid(s, i, 1) >= "A" And Mid(s, i, 1) <= "Z" Then
            retval = retval + Mid(s, i, 1)
        End If
    Next

    ' Then return the return string.                          '
    getOnlyCapitalLetters = retval
End Function
'========================================
' FUNCTION that removes all text from string, and leave only numbers
'========================================
' function taken somewhere from stackoverflow
Function getOnlyDigits(s As String) As String
    ' Variables needed (remember to use "option explicit").   '
    Dim retval As String    ' This is the return string.      '
    Dim i As Integer        ' Counter for character position. '

    ' Initialise return string to empty                       '
    retval = ""

    ' For every character in input string, copy digits to     '
    '   return string.                                        '
    For i = 1 To Len(s)
        If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Then
            retval = retval + Mid(s, i, 1)
        End If
    Next

    ' Then return the return string.                          '
    getOnlyDigits = retval
End Function
'========================================
' FUNCTION that tells if string contains digits
'========================================
' function is adapted using function onlyDigits
Function isDigit(s As String) As Boolean
    ' Variables needed (remember to use "option explicit").   '
    Dim retbool As Boolean    ' This is the return boolean.   '
    Dim i As Integer        ' Counter for character position. '

    ' Initialise return string to empty                       '
    retbool = False

    ' For every character in input string, check if there are '
    ' numbers. Stop if found at least one number              '
    For i = 1 To Len(s)
        If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Then
            retbool = True
            Exit For
        Else
            retbool = False
        End If
        
    Next

    ' Then return the results
    isDigit = retbool
End Function
'========================================
' FUNCTION that count cell color in a range
'========================================
' This is a user defined function! UDF!
' source: https://support.microsoft.com/en-us/help/2815384/count-the-number-of-cells-with-specific-cell-color-by-using-vba
Function CountCcolor(range_data As Range, criteria As Range) As Long
    Dim datax As Range
    Dim xcolor As Long
xcolor = criteria.Interior.ColorIndex
For Each datax In range_data
    If datax.Interior.ColorIndex = xcolor Then
        CountCcolor = CountCcolor + 1
    End If
Next datax
End Function
