Attribute VB_Name = "Functions"
''''''''' (C) 2020 VZ Home Experiments Vladimir Zhbanko https://vladdsm.github.io/myblog_attempt/
''''''''' VBA code to make work with Excel User Forms easier
''''''''' More time to spend on more interesting stuff.
''''''''' For donations and support: https://www.paypal.me/Zhbanko

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

'========================================
' FUNCTION that checks if specific string is present in specific worksheet
'========================================
' This is a user defined function! UDF!
' @Description Purpose of this function to check if there is a matching id on the specific worksheet
' @Param id_string - string, indicating what needs to be matched
' @Param sh_name - string, on which worksheet match would take place
'
' @Return Function return True if there is a match and record exists but when matching field is not empty string

Function isExistingID_for_loop(id_string As String, sh_name As String) As Boolean
    Dim wshDest As Worksheet: Set wshDest = Worksheets(sh_name)
    Dim j As Long: Dim itemIDval As String

    For j = 2 To 2000         ' There can not be more than a 2000 rows really!?
            itemIDval = wshDest.CELLS(j, 1).Value

            ' find where is the match and matching value is not empty
            If itemIDval = id_string And Not itemIDval = "" Then
                isExistingID = True ' there is a match and field is not empty
                Exit For
            Else
                isExistingID = False ' there is no match
            End If

        Next


End Function

'========================================
' FUNCTION that checks if specific string is present in specific worksheet on the column 'A'
'========================================
' This is a user defined function! UDF!
' @Description Purpose of this function to check if there is a matching id on the specific worksheet
' @Detail Use method .Find without for loop, only search inside column 'A'
' @Param id_string - string, indicating what needs to be matched
' @Param sh_name - string, on which worksheet match would take place
' @Usage
' isExistingID("Hello World", "Code Textbook")

' @Return Function return True if there is a match and record exists but when matching field is not empty string

Function isExistingID(id_string As String, sh_name As String) As Boolean
    Dim wshDest As Worksheet: Set wshDest = Worksheets(sh_name)
    
      ' find location of the item on the worksheet
        Dim FindRng As Range
        
        Set FindRng = wshDest.CELLS.Columns("A:A").Find(What:=id_string, LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=False)
        If Not FindRng Is Nothing Then
            'lRowRep = FindRng.Row
            isExistingID = True
        Else
            isExistingID = False
        End If
       
    
End Function

'========================================
' FUNCTION that checks where specific string is present in specific worksheet of column 'A'
'========================================
' This is a user defined function! UDF!
' @Description Purpose of this function to check where there is a matching id on the specific worksheet
' @Detail Function is not fail safe, presence of the string must be checked first
' @Param id_string - string, indicating what needs to be matched
' @Param sh_name - string, on which worksheet match would take place
' @Usage
' getRowID("Hello World", "Code Textbook")

' @Return Function return True if there is a match and record exists but when matching field is not empty string

Function getRowID(id_string As String, sh_name As String) As Long
    
    Dim wshDest As Worksheet: Set wshDest = Worksheets(sh_name)
    
      ' find location of the item on the worksheet
        Dim FindRng As Range
        
        Set FindRng = wshDest.CELLS.Columns("A:A").Find(What:=id_string, LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=False)
        If Not FindRng Is Nothing Then
            getRowID = FindRng.Row

        Else
            getRowID = 0
        End If
        
End Function
'========================================
' FUNCTION that converts string (if possible) to integer
'========================================
'@Description if not possible to convert a - is returned
Function ConvertString(myString)
    Dim finalNumber As Variant
    If IsNumeric(myString) Then
        If IsEmpty(myString) Then
            finalNumber = "-"
        Else
            finalNumber = CInt(myString)
        End If
    Else
        finalNumber = "-"
    End If
    
    ConvertString = finalNumber
End Function

'========================================
' FUNCTION that finds the last available row where to write
'========================================
' This is a user defined function! UDF!
' @Description Purpose of this function to check where is a matching id on the specific worksheet
' @Param id_string - string, indicating what needs to be matched
' @Param sh_name - string, on which worksheet match would take place

' @Return Function return True if there is a match and record exists but when matching field is not empty string

Function getLastEmptyRow(sh_name As String) As Long
    Dim wshDest As Worksheet: Set wshDest = Worksheets(sh_name)
    Dim j As Long: Dim currentRowValue As String
    
        For j = 2 To 2000         ' There can not be more than a 2000 rows really!?
            
            currentRowValue = wshDest.CELLS(j, 3).Value
    
           ' find where is the match and matching value is not empty
            If IsEmpty(currentRowValue) Or currentRowValue = "" Then
                getLastEmptyRow = j ' there is a match and field is not empty
                Exit For
            ElseIf j = 2000 Then
                MsgBox "This table contains more than 2000 Rows. Make sure to archive some records!"
            End If
    
        Next
    
        
End Function


'========================================
' FUNCTION to return status key numeric value based on input text
'========================================
' This is a user defined function! UDF!
' @Description Purpose of this function to return number based on user text
' @Param status_item - string, indicating what needs to be matched

' @Return Function return numeric code defined in the function

Function getStatusKey(status_item As String) As Long

' Define color coding based on the text entries:
If status_item = "" Then
    getStatusKey = 8
ElseIf status_item = "Planned" Then
    getStatusKey = 6
ElseIf status_item = "In Progress" Then
    getStatusKey = 5
ElseIf status_item = "Done" Then
    getStatusKey = 4
ElseIf status_item = "Canceled" Then
    getStatusKey = 2
ElseIf status_item = "Investigation" Then
    getStatusKey = 6
ElseIf status_item = "Roadblock" Then
    getStatusKey = 3
ElseIf status_item = "Delay" Then
    getStatusKey = 7
ElseIf status_item = "Routine" Then
    getStatusKey = 7
ElseIf status_item = "High" Then
    getStatusKey = 3
ElseIf status_item = "Medium" Then
    getStatusKey = 6
ElseIf status_item = "Low" Then
    getStatusKey = 4
Else
    getStatusKey = 8
End If

        
End Function


'========================================
' FUNCTION to return vb color code
'========================================
' This is a user defined function! UDF!
' @Description Purpose of this function to return vbColor Name
' @Param status_item - string, indicating what needs to be matched

' @Return Function returns vbColor Name

Function getvbColor(status_item As String) As Variant

' Define color coding based on the text entries:
If status_item = "" Then
    getvbColor = vbCyan
ElseIf status_item = "Planned" Then
    getvbColor = vbYellow
ElseIf status_item = "In Progress" Then
    getvbColor = vbBlue
ElseIf status_item = "Done" Then
    getvbColor = vbGreen
ElseIf status_item = "Canceled" Then
    getvbColor = vbWhite
ElseIf status_item = "Investigation" Then
    getvbColor = vbYellow
ElseIf status_item = "Roadblock" Then
    getvbColor = vbRed
ElseIf status_item = "Delay" Then
    getvbColor = vbMagenta
ElseIf status_item = "Routine" Then
    getvbColor = vbMagenta
ElseIf status_item = "High" Then
    getvbColor = vbRed
ElseIf status_item = "Medium" Then
    getvbColor = vbYellow
ElseIf status_item = "Low" Then
    getvbColor = vbGreen
Else
    getvbColor = vbCyan
End If

        
End Function
