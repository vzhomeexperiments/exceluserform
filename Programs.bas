Attribute VB_Name = "Programs"
Option Explicit
'========================================
'PASTE PICTURE TO CELL
'========================================
' This Sub gets picture path and the row number where to place picture
' as the column and sheet are fixed we will always use them
Sub PastePicture(picPath, iRow)

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

' Define variables needed
Dim Item As String: Dim Category As String: Dim Key As Integer: Dim Checkpoint As String: Dim Tools As String: Dim Fail As String: Dim Comments As String
Dim SheetName As String

' Initialize variables
SheetName = ActiveSheet.Name
Item = getOnlyDigit(SheetName) & "-" & getAllCapitalLetters(SheetName) & "-" & Range("A" & iRow)
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

' Update Budget Field - Case Study 1
Call UpdateBudget

End Sub

'========================================
'UPDATE USER FORM Report ' CASE STUDY 2
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
Budget = Worksheets("Summary").Range("B8").Value
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

' This Sub creates PowerPoint slide from the given row (iRow) of the Report page
' Important: Go to Tools -> References -> Enable Microsoft PPT Object Library

Sub WorkbooktoPowerPoint(iRow)
    
' Declare variables
    'for PowerPoint slides
    Dim PPT As Object: Dim PPTPres As Object: Dim PPTSlide As Object
    Dim oPicture As PowerPoint.Shape: Dim tboxAction As PowerPoint.Shape
    ' CASE STUDY 3
    Dim tboxCost As PowerPoint.Shape: Dim figCircle As PowerPoint.Shape: Dim Key As Integer
    ' for Excel worksheet
    Dim wshSrc As Worksheet: Set wshSrc = Worksheets("Report")
    Dim Item As String: Dim Category As String: Dim Issue As String: Dim Action As String: Dim Cost As String
    Dim picPath As String
    
' Open PowerPoint, Add Presentation, Make it visible
    Set PPT = CreateObject("PowerPoint.Application")
    Set PPTPres = PPT.Presentations.Add
    PPT.Visible = True
        
' Set the data to variables
    Item = wshSrc.Range("A" & iRow).Value
    Category = wshSrc.Range("C" & iRow).Value
    Key = Range("D" & iRow).Interior.ColorIndex ' save color property value to Key variable 'CASE STUDY 3
    Issue = wshSrc.Range("E" & iRow).Value
    Action = wshSrc.Range("F" & iRow).Value
    Cost = wshSrc.Range("G" & iRow).Value
    picPath = wshSrc.Range("H" & iRow).Value ' extend this list
    
' Add new blank slide and set the title
    Set PPTSlide = PPTPres.Slides.Add(Index:=1, Layout:=ppLayoutTitleOnly)
    PPTSlide.Select: PPTSlide.Shapes.Title.TextFrame.TextRange.Text = Category & "-" & Issue
         
' Paste the picture and adjust its position
    Set oPicture = PPTSlide.Shapes.AddPicture(picPath, msoFalse, msoTrue, Left:=100, Top:=150, Width:=400, Height:=300)

' Add text box for Action
    Set tboxAction = PPTSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, Left:=500, Top:=150, Width:=400, Height:=250)
    
    ' Format the text range
    With tboxAction.TextFrame.TextRange
        .Text = "Action Suggested: " & Action
        With .Font
            .Size = 24
            .Name = "Arial"
        End With
    End With

'CASE STUDY 3 - Add another text box
' Add text box for cost
Set tboxCost = PPTSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, Left:=500, Top:=450, Width:=400, Height:=250)

    With tboxCost.TextFrame.TextRange
        .Text = "Approx.cost: " & Cost & " CHF"
        With .Font
            .Size = 24
            .Name = "Arial"
        End With
    End With
    
' Add circle with issue color code
Set figCircle = PPTSlide.Shapes.AddShape(Type:=msoShapeOval, Left:=550, Top:=350, Width:=70, Height:=70)
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
PPTPres.Application.ActivePresentation.ApplyTemplate "C:\Users\fxtrams\Downloads\WidescreenPresentation.potx"

' Memory Cleanup (Step useful when adding for loop)
    PPT.Activate
    Set PPTSlide = Nothing
    Set PPTPres = Nothing
    Set PPT = Nothing
               
End Sub

'========================================
'SAVE THIS FILE TO A SHAREPOINT
'========================================

Sub Push2SharePoint()

    ' define variables
    Dim SharePointPath As Variant
    Dim FileAsNamed As Variant
    ' retrieve SharePoint path indicated by the user inside Excel Sheet named "Select" on cell B9
    SharePointPath = ThisWorkbook.Sheets("Summary").Range("B9").Text
    ' provide some error message if it's not populated
    On Error GoTo NoStorageSelected
    If Not SharePointPath <> False Then
        'Displaying a message if file not choosedn in the above step
        MsgBox "No storage space was selected.", vbExclamation, "Sorry!"
        'And existing from the procedure
        Exit Sub
    Else
        'Create the new file name, note we place data format in ISO 8601 format in front of the file name
        FileAsNamed = SharePointPath & Year(Date) & "-" & Month(Date) & "-" & Day(Date) & "_" & ThisWorkbook.Name

        ' save the document in the current location
        ThisWorkbook.Save
        ' push the document to the Archive location
        ThisWorkbook.SaveAs Filename:=FileAsNamed, FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False


    End If

Exit Sub
' Error Management
NoStorageSelected:
           MsgBox "Error: Excel can not reach SharePoint Folder Storage location" & vbCrLf & _
           "Possible reasons are: Storage location was not defined in the Worksheet 'Select' cell B9 or " & vbCrLf & _
           "Not having sufficient previledges to access SharePoint location " & vbCrLf & _
           "Make sure to add forward slash after SharePoint Document Library"
           Exit Sub

End Sub



'========================================
'SEND EMAILS TO THE USERS
'========================================

Sub Send_Email()

' define variables
Dim OutApp As Object
Dim OutMail As Object
Dim wshS As Worksheet: Set wshS = Worksheets("Summary")
Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)
Dim question As VbMsgBoxResult
Dim EmailTo As Range: Dim EmailCc As Range
Dim cline As Range: Dim tline As Range
Dim sTo As String: Dim cTo As String

' setup question for the message box
question = MsgBox("Sending email to all contacts, Are you sure? [Preview will follow]", vbYesNo + vbQuestion)
' retrieve emails from the worksheet
Set EmailTo = wshS.Range("B10:B11")
Set EmailCc = wshS.Range("B12:B14")
' joining string for email 'To'
For Each cline In EmailTo
        sTo = sTo & ";" & cline.Value
Next
' joining string for email 'Cc'
For Each tline In EmailCc
        cTo = cTo & ";" & tline.Value
Next
' cleaning of the strings
sTo = Mid(sTo, 2)
cTo = Mid(cTo, 2)

If question = vbYes Then
    
    With OutMail
        .To = sTo
        .CC = cTo
        .BCC = ""
        .Subject = "CarStatusReport" & wshS.Range("B2").Text & "-" & wshS.Range("B3").Text
        .Body = "Dear participants, thank you for productive work." & _
                vbNewLine & "Please find the file attached: " & ThisWorkbook.Name & _
                vbNewLine & "Best regards"
        .Attachments.Add ThisWorkbook.Path & "\" & ThisWorkbook.Name 'add this file as attachment
        .Display 'with preview mode
        '.Send   'sending email directly
    End With
    
    Set OutMail = Nothing
    Set OutApp = Nothing
    
ElseIf question = vbNo Then

    Exit Sub
    
End If

End Sub

