VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MyCarCheckListFormR 
   Caption         =   "User Form for Car Evaluation - Report"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10515
   OleObjectBlob   =   "MyCarCheckListFormR.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MyCarCheckListFormR"
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
' user should be able to create powerpoint slide from the selected row
' =======================================================================================================
Private Sub buttonPPT_Click()
lRow = Me.tboxRow
WorkbooktoPowerPoint lRow
End Sub

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
UpdateInputsR Me.tboxRow.Value - 1
End Sub
Private Sub SpinButton1_SpinDown()
UpdateInputsR Me.tboxRow.Value + 1
End Sub

' =======================================================================================================
' first form initialization bringing default values
' =======================================================================================================
Private Sub UserForm_Initialize()
' CASE STUDY 2: Lock field comment not to be editable
Me.tboxComments.Locked = True

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
' path of the picture is saved into Global variable 'picPath'

End Sub
' =======================================================================================================
' Copy/Update to the Report page
' =======================================================================================================
' This code will copy or rewrite form data from UserForm to the Report page
Private Sub buttonSubmit_Click()

Dim i As Integer: Dim lRow As Long: Dim lCol As Long
Dim wshDest As Worksheet: Set wshDest = Worksheets("Report")

    ' =======================================
    ' below portion will handle updating the Action page from the UserForm
    ' =======================================

        ' find the next empty row in the destination sheet
        wshDest.Activate

        ' set specific row in the table that requires the update
        lRow = Me.tboxRow.Value
        
        ' re-populate the Result sheet
            wshDest.Cells(lRow, 3).Value = Me.tboxCategory.Value        'Category
            wshDest.Cells(lRow, 4).Interior.ColorIndex = Me.tboxKey.Value        'Key color
            wshDest.Cells(lRow, 5).Value = Me.tboxComments.Value     'Comments
            wshDest.Cells(lRow, 6).Value = Me.tboxAction.Value           'Action
            wshDest.Cells(lRow, 7).Value = Me.tboxCost.Value          'Cost
            ' Update Picture Path only if picture was selected by user
            If Not picPath = "" Then
            wshDest.Cells(lRow, 8).Value = picPath       'Path of the picture
            End If
        ' CASE STUDY 1 - Adding Budget field
          Call UpdateBudget
            
' =======================================
' Code will paste picture to the Result sheet (note: picture will be placed on top of previous one :)))
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
