Attribute VB_Name = "Programs"
Option Explicit

        ' (C) 2017 VZ Home Experiments Vladimir Zhbanko //vz.home.experiments@gmail.com
        ' VBA code to make work with Excel User Forms easier
        ' More time to spend on more interesting stuff.
        
        ' ===================================
        ' Version history updates information:
        ' ===================================
        ' version 1: First Commit
        ' version 2: For Case Study
        ' version 3: Solution for Case Study 3
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
Else
MyCarCheckListForm.optionYes.Value = False   ' <- CASE STUDY 2 - Reset for using Up/Down Arrows
End If
                 
If Fail = "No" Then
MyCarCheckListForm.optionNo.Value = True
Else
MyCarCheckListForm.optionNo.Value = False      ' <- CASE STUDY 2 - Reset for using Up/Down Arrows
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

' Update budget Field
Call UpdateBudget

End Sub

'========================================
'UPDATE USER FORM Report
'========================================

' This Sub update the input information to the User Form Report
' information is found using 'iRow' argument that represent worksheet row
' Goal is to bring all available information from the report page including picture and create powerpoint slide

Sub UpdateInputsR(iRow)


' *** ---------------------------------------------------------------------------------------
' Define variables needed
Dim Item As String: Dim Category As String: Dim Key As Integer: Dim Comments As String
Dim SheetName As String: Dim Action As String: Dim Cost As String: Dim picPath As String

' Initialize variables
SheetName = ActiveSheet.Name
Item = Range("A" & iRow).Value
Category = Range("C" & iRow).Value
Key = Range("D" & iRow).Interior.ColorIndex ' save color property value to Key variable
Comments = Range("E" & iRow).Value
Action = Range("F" & iRow).Value
Cost = Range("G" & iRow).Value
picPath = Range("H" & iRow).Value ' get picture path
' Defining item reference
MyCarCheckListFormR.tboxItem.Text = Item
' Store value of iRow to the form
MyCarCheckListFormR.tboxRow.Value = iRow
' Store name of Category
MyCarCheckListFormR.tboxCategory.Text = Category
' Put color index number to the tboxKey
MyCarCheckListFormR.tboxKey.Value = Key
' Put color to the text box
If Key = 3 Then
MyCarCheckListFormR.tboxKey.BackColor = vbRed
ElseIf Key = 14 Then
MyCarCheckListFormR.tboxKey.BackColor = vbGreen
ElseIf Key = 6 Then
MyCarCheckListFormR.tboxKey.BackColor = vbYellow
ElseIf Key = 7 Then
MyCarCheckListFormR.tboxKey.BackColor = vbMagenta
End If

' Defining Issue from the Comment
MyCarCheckListFormR.tboxComments.Text = Comments
' Defining Issue from the Comment
MyCarCheckListFormR.tboxAction.Text = Action
' Defining Issue from the Comment
MyCarCheckListFormR.tboxCost.Text = Cost
' Bringing a picture into the image box
MyCarCheckListFormR.imageReport.Picture = LoadPicture(picPath)

' Update budget Field
Call UpdateBudget

End Sub


'========================================
'UPDATE BUDGET FIELD OF THE USERFORM
'========================================

' This Sub update the input information to the User Form
' CASE STUDY 1 - Adding Budget field

Sub UpdateBudget()

Dim Budget As Double: Dim SumCost As Double 'Declare variables

' Initialize them
Budget = Worksheets("Summary").Range("B9").Value
SumCost = Application.WorksheetFunction.Sum(ThisWorkbook.Sheets("Report").Range("G2:G500"))
' Bring info to the field
MyCarCheckListForm.tboxBudget.Value = Budget - SumCost
MyCarCheckListFormR.tboxBudget.Value = Budget - SumCost
' Color red if Budget is below zero
If Budget - SumCost < 0 Then
MyCarCheckListForm.tboxBudget.BackColor = vbRed
MyCarCheckListFormR.tboxBudget.BackColor = vbRed
Else
MyCarCheckListForm.tboxBudget.BackColor = vbGreen
MyCarCheckListFormR.tboxBudget.BackColor = vbGreen
End If

End Sub

'========================================
'CREATE POWER POINT SLIDES
'========================================

' This Sub create powerpoint slide from every row of the Report page
' code taken and modified from http://www.dummies.com/software/microsoft-office/excel/sending-excel-data-to-a-powerpoint-presentation/
' combined with code from: https://stackoverflow.com/questions/31653192/how-to-add-a-picture-to-a-powerpoint-slide-from-excel-solved

Sub WorkbooktoPowerPoint(iRow)
    
'Step 1:  Declare your variables
    Dim pp As Object: Dim PPPres As Object: Dim PPSlide As Object: Dim wrkbook As Workbook: Dim xlwksht As Worksheet
    Dim oPicture As PowerPoint.Shape: Dim tboxAction As PowerPoint.Shape
    ' CASE STUDY 3
    Dim tboxCost As PowerPoint.Shape: Dim figCircle As PowerPoint.Shape: Dim Key As Integer
       
    Dim MyItem As String: Dim MyCategory As String: Dim MyKey As String: Dim MyIssue As String: Dim MyAction As String: Dim MyCost As String
    Dim MyPicPath As String: Set xlwksht = Worksheets("Report")
    
'Step 2:  Open PowerPoint, add a new presentation and make visible
    Set pp = CreateObject("PowerPoint.Application")
    Set PPPres = pp.Presentations.Add
    pp.Visible = True
        
'Step 3:  Set the ranges for your data and title
    MyItem = xlwksht.Range("A" & iRow).Value  '<<<Change this range
    MyCategory = xlwksht.Range("C" & iRow).Value
    Key = Range("D" & iRow).Interior.ColorIndex ' save color property value to Key variable 'CASE STUDY 3
    MyIssue = xlwksht.Range("E" & iRow).Value
    MyAction = xlwksht.Range("F" & iRow).Value
    MyCost = xlwksht.Range("G" & iRow).Value
    MyPicPath = xlwksht.Range("H" & iRow).Value
    
'Step 4:  Add new blank slide and set the title
    Set PPSlide = PPPres.Slides.Add(Index:=1, Layout:=ppLayoutTitleOnly) 'SlideCount + Index:=1, Index:=1, Layout:=ppLayoutTitleOnly
    PPSlide.Select
    PPSlide.Shapes.Title.TextFrame.TextRange.Text = MyCategory & "-" & MyIssue
         
'Step 5:  Paste the picture and adjust its position
    Set oPicture = PPSlide.Shapes.AddPicture(MyPicPath, msoFalse, msoTrue, Left:=100, Top:=150, Width:=400, Height:=300)

'Step 5.1: Add text box for Action
Set tboxAction = PPSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, Left:=500, Top:=150, Width:=400, Height:=250)

    With tboxAction.TextFrame.TextRange
        .Text = "Action Suggested: " & MyAction
        With .Font
            .Size = 24
            .Name = "Arial"
        End With
    End With

'CASE STUDY 3 - Add another text box
'Step 5.2: Add text box for cost
Set tboxCost = PPSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, Left:=500, Top:=450, Width:=400, Height:=250)

    With tboxCost.TextFrame.TextRange
        .Text = "Approx.cost: " & MyCost & " CHF"
        With .Font
            .Size = 24
            .Name = "Arial"
        End With
    End With
    
'Step 5.3: Add circle with issue color code
Set figCircle = PPSlide.Shapes.AddShape(Type:=msoShapeOval, Left:=550, Top:=350, Width:=70, Height:=70)
          'Decide which color
            If Key = 3 Then
                figCircle.Fill.ForeColor.RGB = vbRed
            ElseIf Key = 14 Then
                figCircle.Fill.ForeColor.RGB = vbGreen
            ElseIf Key = 6 Then
                figCircle.Fill.ForeColor.RGB = vbYellow
            ElseIf Key = 7 Then
                figCircle.Fill.ForeColor.RGB = vbMagenta
            End If

'Step 5.4: Apply Template
On Error Resume Next
' set your path...
PPPres.Application.ActivePresentation.ApplyTemplate "C:\Users\fxtrams\Downloads\WidescreenPresentation.potx"

'Step 6:  Memory Cleanup
    pp.Activate
    Set PPSlide = Nothing
    Set PPPres = Nothing
    Set pp = Nothing
               
End Sub
