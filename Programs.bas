Attribute VB_Name = "Programs"
Option Explicit

'========================================
'PASTE PICTURE TO CELL
'========================================

' This Sub gets picture path and the row number where to place picture
' as the column and sheet are fixed we will always use them
Sub PastePicture(picPath, iRow)
Dim sShape As Shape
  ' resize row height first
  Worksheets("Report").Rows(iRow).RowHeight = 79

      With Worksheets("Report").Pictures.Insert(picPath)
      
        With .ShapeRange
            .LockAspectRatio = msoTrue
            .Width = 90         'width of the picture
            .Height = 75        'height of the picture
        End With
        ' define where to place the picture in the cell
        .Left = Worksheets("Report").Cells(iRow, 2).Left + 2
        .Top = Worksheets("Report").Cells(iRow, 2).Top + 2
        .Placement = 1
        .PrintObject = True
        .Name = "Sample" & iRow    ' use .Name property to name the picture with known name
      
        ' optimize RAM usage by keeping the picture in the cell, not linked to folder source
        ' using the "known" name we perform operation on the picture
        With ActiveSheet.Shapes.Range(Array("Sample" & iRow)).Select
          Selection.Cut
          Cells(iRow, 2).Select
          ActiveSheet.Pictures.Paste.Select
          ' method to move the Shape
          Selection.ShapeRange.IncrementLeft 2
          Selection.ShapeRange.IncrementTop 2
          Cells(iRow, 2).Select
        End With
       
    End With
   
End Sub

'========================================
'UPDATE USER FORM INPUTS
'========================================

' This Sub update the input information to the User Form
' information is found using 'iRow' argument that represent worksheet row

Sub UpdateInputs(iRow)

' *** ---------------------------------------------------------------------------------------
' Define variables needed
Dim Item As String: Dim Category As String: Dim Key As Integer: Dim Checkpoint As String: Dim Tools As String: Dim Fail As String: Dim Comments As String
Dim SheetName As String

' Initialize variables
SheetName = ActiveSheet.Name
Item = getOnlyDigits(SheetName) & "-" & getOnlyCapitalLetters(SheetName) & "-" & Range("A" & iRow)
Category = Range("B" & iRow).Value
Key = Range("C" & iRow).Interior.ColorIndex ' save color property value to Key variable
Checkpoint = Range("D" & iRow).Value
Tools = Range("E" & iRow).Value
Fail = Range("F" & iRow).Value
Comments = Range("G" & iRow).Value

' Defining page reference as D1.Nr
MyCarCheckListForm.tboxItem.Text = Item
' Store value of iRow to the form
MyCarCheckListForm.tboxRow.Value = iRow
' Store name of Worksheet
MyCarCheckListForm.tboxSheet.Value = SheetName
' Store name of Category
MyCarCheckListForm.tboxCategory.Text = Category
' Store name of Tools
MyCarCheckListForm.tboxTools.Text = Tools
' Copy the Checkpoint for better overview
MyCarCheckListForm.tboxCheckpoint.Text = Checkpoint

' Returning a Fail Option
If Fail = "Yes" Then
MyCarCheckListForm.optionYes.Value = True
End If
If Fail = "Fix needed" Then
MyCarCheckListForm.optionNo.Value = True
End If
' Put color index number to the tboxKey
MyCarCheckListForm.tboxKey.Value = Key
' Put color to the text box
If Key = 3 Then
MyCarCheckListForm.tboxKey.BackColor = vbRed
ElseIf Key = 14 Then
MyCarCheckListForm.tboxKey.BackColor = vbGreen
ElseIf Key = 6 Then
MyCarCheckListForm.tboxKey.BackColor = vbYellow
ElseIf Key = 7 Then
MyCarCheckListForm.tboxKey.BackColor = vbMagenta
End If

' Defining Issue from the Comment
MyCarCheckListForm.tboxComments.Text = Comments

End Sub
