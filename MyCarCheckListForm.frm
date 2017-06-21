VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MyCarCheckListForm 
   Caption         =   "User Form for Car Evaluation"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10515
   OleObjectBlob   =   "MyCarCheckListForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MyCarCheckListForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


        ' (C) 2017 VZ Home Experiments Vladimir Zhbanko //vz.home.experiments@gmail.com
        ' VBA code to make work with Excel User Forms easier
        ' More time to spend on more interesting stuff.
        
        ' ===================================
        ' Version history updates information:
        ' ===================================
        ' version 1: First Commit
        ' version 2: For Case Study
        ' version 3: Solution for Case Study 3
    
' declaring global variables for cross using in the other functions
Public Fail As String           ' 2 types Yes/No
Public picPath As String        ' string is containing the path to the picture file
Public lRow As Long             ' variable to pass row information
' =======================================================================================================
' this button closes the form
' =======================================================================================================
Private Sub buttonCancel_Click()
Unload Me
End Sub
' =======================================================================================================
' information about the program shown by clicking on the button "I am inspired"
' =======================================================================================================
Private Sub buttonHelp_Click()
MsgBox "User Form for Car Evaluation" & vbCrLf & "(C) 2017 VZ Home Experiments vz.home.experiments@gmail.com", vbOKOnly + vbInformation, "I am inspired!"
End Sub
' =======================================================================================================
' Case Study: Add Spin Buttons control to update UF
' =======================================================================================================
Private Sub SpinButton1_SpinUp()
If Me.tboxRow.Value <= 2 Then
    Exit Sub
End If
Worksheets(Me.tboxSheet.Text).Activate
UpdateInputs Me.tboxRow.Value - 1
End Sub
Private Sub SpinButton1_SpinDown()
Worksheets(Me.tboxSheet.Text).Activate
UpdateInputs Me.tboxRow.Value + 1
End Sub

' =======================================================================================================
' first form initialization bringing default values
' =======================================================================================================
Private Sub UserForm_Initialize()
' Not used; code below will be executed on form initialization

'' NOTE: Not optimal way as budget field will not be updated when arrows buttons are used!
' Bring data from budget field
'Dim Budget As Double: Dim SumCost As Double
'Budget = Worksheets("Summary").Range("B9").Value
'SumCost = Application.WorksheetFunction.Sum(ThisWorkbook.Sheets("Report").Range("G2:G500"))
'MyCarCheckListForm.tboxBudget.Value = Budget - SumCost

End Sub
' =======================================================================================================
' User Dialogue "Import Picture"
' =======================================================================================================
' this portion should point to the picture to enter to the userform
' user select picture browsing to the file and picture is grabbed inside the form
' path to the picture will be stored into Public variable so
' user will continue to write issue description and upon submitting picture is placed to the cell...
' Button "Insert Picture"
Private Sub buttonPicture_Click()

' File dialog to load picture into the form
With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .ButtonName = "Submit"
    .Title = "Select an image file"
    .Filters.Add "Image", "*.gif; *.jpg; *.jpeg", 1

    If .Show = -1 Then
        'file has been selected
        picPath = .SelectedItems(1) ' this will save path to the picture!

        'display preview image in image control
        Me.imageReport.PictureSizeMode = fmPictureSizeModeZoom
        Me.imageReport.Picture = LoadPicture(picPath)

    Else
        'user aborted the dialog

    End If
End With

' picture is now in the image box
' path of the picture picPath is saved into Global variable

End Sub
' =======================================================================================================
' Copy to the Report page
' =======================================================================================================
' This code will copy form data from UserForm to the Report page
' Also required to paste comment and score to the reference page if it was changed
' Report page should increase it's size by one row automatically
Private Sub buttonSubmit_Click()

Dim i As Integer: Dim lRow As Long: Dim lCol As Long: Dim nextRowValue As String
Dim wshDest As Worksheet: Set wshDest = Worksheets("Report")
Dim wshSource As Worksheet: Set wshSource = Worksheets(Me.tboxSheet.Text)

    ' =======================================
    ' code below will check position of radio buttons
    ' =======================================
    If (Me.optionYes.Value = True) Then
    Fail = "Yes"
    End If
    
    If (Me.optionNo.Value = True) Then
    Fail = "No"
    End If
    
    ' CASE STUDY 1 - Adding protection against incomplete entry
     If (Me.optionYes.Value = False) And (Me.optionNo.Value = False) Then
         Me.optionNo.SetFocus
         MsgBox "Check must either pass or fail, please choose at least one option"
         Exit Sub
     End If
       
    ' =======================================
    ' This portion refreshes the comment and the score on the source sheet
    ' =======================================
    ' refreshing data on the source sheet
    ' define the source sheet
    ' write the comment and score to the source sheet (it might be changed)
    wshSource.Cells(Me.tboxRow.Value, 6) = Fail               'score
    wshSource.Cells(Me.tboxRow.Value, 7) = Me.tboxComments.Value   'comment
    
    ' =======================================
    ' below portion will handle updating the Action page from the UserForm
    ' =======================================
    
     
    ' only if cboxNeedAction is true
    If Me.cboxNeedAction.Value = False Then
        ' exit sub if action is not needed
        MsgBox "Comment and Score are updated, No Action is created", vbOKOnly + vbInformation, "Source sheet is refreshed"
        Exit Sub
     
    Else
        ' find the next empty row in the destination sheet
        wshDest.Activate
        ' method below will fill  the next available empty row
        ' lRow will contain the last written row (ready to write)
        For i = 1 To 2000         ' There can not be more than a 2000 rows really!?
            currentRowValue = Cells(i, 3).Value
            nextRowValue = Cells(i + 1, 1).Value ' saving content of the next rows to add rows dynamically
    
           ' find where is the last available row in the table
            If IsEmpty(currentRowValue) Or currentRowValue = "" Then
                lRow = i
                If isDigit(Cells(i - 1, 1).Value) = False Then ' if the cell is not number it is a header
                    wshDest.Cells(i, 1).Value = 1 ' place the starting number
                Else
                    wshDest.Cells(i, 1).Value = wshDest.Cells(i - 1, 1).Value + 1 ' place the consecutive number
                    wshDest.Cells(i + 1, 1).Value = wshDest.Cells(i, 1).Value + 1 ' place the consecutive number
                End If
                    Exit For
            End If
    
        Next
        
        ' check for a completness of the form when gaps are identified
        ' logic behind: If Fail is 'Yes' then Comments and Actions are required!
        If (Me.optionYes.Value = True) And (Trim(Me.tboxComments.Value) = "") Then
            Me.tboxComments.SetFocus
            MsgBox "Please complete the Action and Comment fields of the form as gaps are identified"
                Exit Sub
        End If
    
        
        
        ' populate the Result sheet
            wshDest.Cells(lRow, 3).Value = Me.tboxCategory.Value        'Category
            wshDest.Cells(lRow, 4).Interior.ColorIndex = Me.tboxKey.Value        'Key color
            wshDest.Cells(lRow, 5).Value = Me.tboxComments.Value     'Comments
            wshDest.Cells(lRow, 6).Value = Me.tboxAction.Value           'Action
            wshDest.Cells(lRow, 7).Value = Me.tboxCost.Value          'Cost
            wshDest.Cells(lRow, 8).Value = picPath       'Path of the picture
            
        ' CASE STUDY 1 - Clear the data to be able to fill more again
            Me.tboxComments.Value = ""
            Me.tboxAction.Value = ""
            Me.tboxCost.Value = ""
            Me.cboxNeedAction.Value = False
            Me.optionNo = False
            Me.optionYes = False
           
        ' CASE STUDY 1 - Adding Budget field
          Call UpdateBudget


    End If
    
' =======================================
' Code will paste picture to the Result sheet
' =======================================
' exit if there was no picture added
  If picPath = "" Then
      Exit Sub
  Else
  ' add picture using function PastePicture (see module Functions)
    PastePicture picPath, lRow
  End If
 
End Sub


' =======================================================================================================
