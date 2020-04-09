VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PlanningForm 
   Caption         =   "User Form for Goals Setting"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10515
   OleObjectBlob   =   "PlanningForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PlanningForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''' (C) 2020 VZ Home Experiments Vladimir Zhbanko https://vladdsm.github.io/myblog_attempt/
''''''''' VBA code to use User Forms to add data to Excel Spreadsheets
''''''''' More time to spend on more interesting stuff.
''''''''' This is a really crazy attempt to develop really complicated user interface!
''''''''' For donations and support: https://www.paypal.me/Zhbanko
' =======================================================================================================
' declaring global variables for cross using in the programs of the UserForm
' =======================================================================================================
Public Status As String         ' Temporal status of the activity e.g. Planned, Done, In Progress, etc
Public StatusKey As Long
Public Situation As String      ' Situation e.g. Delay, Roadblock, Progress, etc
Public SituationKey As Long
Public picPath As String        ' Path to the picture file
Public lRow As Long             ' Row information
Public itemID As String
Public LTGoal As String

Private Sub buttonPivot_Click()
Call CreatePivotTable
End Sub

Private Sub buttonPPT_Click()
WorkbooktoPowerPoint (Me.tboxRow.Value)
End Sub

' =======================================================================================================
' Form initialization bringing default values
' =======================================================================================================
Private Sub UserForm_Initialize()
' Code below will be executed on form initialization
' add here code to initialize form fields e.g. uncomment line below to display help message once user opens the form:
' MsgBox "User Form to Set Goals" & vbCrLf & "(C) 2020 Vladimir Zhbanko https://vladdsm.github.io/myblog_attempt/", vbOKOnly + vbInformation, "I am inspired!"
' initialize lists in the comboboxes:
cboxStatus.List = [Summary!D2:D10].Value
cboxSituation.List = [Summary!H2:H6].Value
cboxLTGoal.List = [Summary!B2:B10].Value
cboxSituation.List = [Summary!H2:H10].Value
cboxOwner.List = [Summary!C2:C10].Value
' initialize radio buttons:
optionYes.Value = False
optionNo.Value = True
' initialize picture:
PlanningForm.imageReport.Picture = LoadPicture([Summary!J2].Value)
End Sub
' =======================================================================================================
' Button to close the Form
' =======================================================================================================
Private Sub buttonCancel_Click()
Unload Me
End Sub
' =======================================================================================================
' Information about the program shown by clicking on the button "Instructions"
' =======================================================================================================
Private Sub buttonInstructions_Click()

MsgBox "User Form to Set Goals" & vbCrLf & _
" " & vbCrLf & _
"Please read it carefully!!!" & vbCrLf & _
" ", vbYes + vbInformation, "I am inspired!"

MsgBox "User Form to Set Goals" & vbCrLf & _
" " & vbCrLf & _
"Use this form to add detailed information about the activity" & vbCrLf & _
"Elaborate information about the Activity to Achieve the Goal" & vbCrLf & _
"Worksheet Report contains detailed information about the activity" & vbCrLf & _
"Worksheet Planning contains summary information", vbOKOnly + vbInformation, "I am inspired!"

MsgBox "Important usage information!" & vbCrLf & _
" " & vbCrLf & _
"Do not rename, add, remove or hide columns of the Worksheets!!!" & vbCrLf & _
" " & vbCrLf & _
"Important: Using Arrows Up/Down will not save records!!!" & vbCrLf & _
"You can only change relevant information on the Sheet Summary", vbOK + vbExclamation, "I am inspired!"

MsgBox "User Form to Set Goals" & vbCrLf & _
" " & vbCrLf & _
"UserForm can be invoked by double clicking on the rows of worksheets Planning and Report" & vbCrLf & _
"Worksheet Planning is used to have a brief overview of the situation and plan goals" & vbCrLf & _
"Worksheet Report is used to add detailed information and track progress" & vbCrLf & _
"Make sure to Archive some records when they are completed and too old", vbOKOnly + vbInformation, "I am inspired!"

End Sub

' =======================================================================================================
' Information about the program shown by clicking on the button "Help"
' =======================================================================================================
Private Sub buttonHelp_Click()
MsgBox "User Form to Set Goals" & vbCrLf & "(C) 2020 Vladimir Zhbanko https://vladdsm.github.io/myblog_attempt/", vbOKOnly + vbInformation, "I am inspired!"
End Sub
' =======================================================================================================
' Spin Buttons control
' =======================================================================================================
' @Description: Spin-buttons allow to easily scroll through the records
Private Sub SpinButton1_SpinUp()
If Me.tboxRow.Value <= 2 Then
    Exit Sub
End If
Worksheets(Me.tboxSheet.Text).Activate
UpdateInputs Me.tboxRow.Value - 1
End Sub
' @Description: Spin-buttons allow to easily scroll through the records
Private Sub SpinButton1_SpinDown()
If Me.tboxRow.Value >= 199 Then
    MsgBox "This list may not contain too many records, please archive some of them!"
    Exit Sub
End If
If Me.tboxItem.Value = "" Then
    MsgBox "User must define numeric and unique id first! Closing the Form!"
    Unload Me
End If
Worksheets(Me.tboxSheet.Text).Activate
UpdateInputs Me.tboxRow.Value + 1
End Sub
' =======================================================================================================
' Store value of combobox to the variable used in the form
' =======================================================================================================
Private Sub cboxStatus_Change()
' store value of combobox to the global variable
Status = cboxStatus.Value
' refresh value of the color key box
StatusKey = getStatusKey(Status)
Me.tboxStatusKey = StatusKey
Me.tboxStatusKey.BackColor = getvbColor(Status)
End Sub
' =======================================================================================================
' Store value of combobox to the variable used in the form
' =======================================================================================================
Private Sub cboxSituation_Change()
' store value of combobox to the global variable
Situation = cboxSituation.Value
' refresh value of the color key box
SituationKey = getStatusKey(Situation)
Me.tboxSituationKey = SituationKey
Me.tboxSituationKey.BackColor = getvbColor(Situation)
End Sub
' =======================================================================================================
' Store value of combobox to the variable used in the form
' =======================================================================================================
Private Sub cboxLTGoal_Change()
' store value of combobox to the global variable
LTGoal = cboxLTGoal.Value
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
        ' add picture path to the text box field
        Me.tboxPath.Value = picPath

    Else
        ' executed when nothing was selected

    End If
End With

' picture is now in the image box
' path of the picture picPath is saved into Global variable

End Sub
' =======================================================================================================
' Write all the information to the Planning and Report Worksheets
' =======================================================================================================
' This code will write form data from fields to the Planning and Report page
' User Form fields are acting like a master!
' Note: Using Arrows Up/Down will not save records!!!
' Whenever there is no matching id on both sheets Planning and Report a new records are created on both sheets
' Report page should increase it's size by one row automatically if such record is new
Private Sub buttonSave_Click()

' Define needed variables
Dim lRowRep As Long: Dim lRowPln As Long: Dim idRowRep As Long: Dim idRowPnl As Long
Dim Archive As Boolean: Dim isMatchingIdOnPlanning As Boolean: Dim isMatchingIdOnReport As Boolean

Dim wshRep As Worksheet: Set wshRep = Worksheets("Report")
Dim wshPln As Worksheet: Set wshPln = Worksheets("Planning")
' =======================================
' Algorithm:
' Step 1. Check position of Archive button
' Step 2. Determine where to save the records
' Step 2.0. Abort saving if data is not complete?
' Step 2.1. Save new record on the Report Page
' Step 2.2. Update existing records on Planning and Report Page
' Step 1.
' Step 1.

    ' =======================================
    ' 1. code below will check position of radio buttons
    ' =======================================
    If (Me.optionYes.Value = True) Then
        Archive = True
        ' Abort saving!
        Me.optionYes.SetFocus
            MsgBox "Selected option 'Yes' to Archive this record! Use button 'Archive' instead!"
        Exit Sub
    End If
    
    If (Me.optionNo.Value = True) Then
        Archive = False
    End If

    
    ' =======================================
    ' 2. Determine where to save the records
    ' =======================================
    ' Find if this is the new or existing record
    ' Find if matching data [id] are present on the different worksheet Report/Planning
    ' True when Item id in the UF correspond to existing record in WS Planning
    ' False when Item id in the UF is not present on existing record
    isMatchingIdOnPlanning = isExistingID(Me.tboxItem.Value, "Planning")
    'True when Item id in the UF correspond to existing record in WS Report
    isMatchingIdOnReport = isExistingID(Me.tboxItem.Value, "Report")
    
    ' Find the next available rows
    lRowRep = getLastEmptyRow("Report")
    lRowPln = getLastEmptyRow("Planning")
    
    ' =======================================
    ' Step 2.0. Abort saving if data is not complete
    ' =======================================
    ' Fail safe - data is not complete
    If (Trim(Me.tboxActivity.Value) = "") Then
        MsgBox "Aborting saving, data field Activity is empty"
        Exit Sub
    ' Fail safe - abort saving as user started without identifying id
    ElseIf (Trim(Me.tboxItem.Value) = "") Then
        MsgBox "No id is present... aborting saving... close the Form and add numeric and unique id first!"
        Exit Sub
    ElseIf isMatchingIdOnPlanning = False And isMatchingIdOnReport = False Then
        MsgBox "No id on Planing WS and nothing on Report...aborting saving...add numeric and unique id first!"
        Exit Sub
    
    ' New record to write on the WS Report
    ElseIf isMatchingIdOnPlanning = True And isMatchingIdOnReport = False Then
    ' =======================================
    ' Step 2.1. Save new record on the Report Page
    ' =======================================
        MsgBox "Numbered id on Planing WS and nothing yet on Report...creating new record on Report"
        
        idRowPln = getRowID(Me.tboxItem.Value, "Planning")
        
        wshPln.CELLS(idRowPln, 1) = Me.tboxItem.Value                       'Item
        wshPln.CELLS(idRowPln, 2) = Me.cboxOwner.Value                      'Owner
        wshPln.CELLS(idRowPln, 3) = Me.tboxSTGoal.Value                     'STGoal
        wshPln.CELLS(idRowPln, 4) = Me.cboxLTGoal.Value                     'LTGoal
        wshPln.CELLS(idRowPln, 5) = Me.tboxActivity.Value                   'Activity
        wshPln.CELLS(idRowPln, 6) = Me.cboxStatus.Value                     'Status
        wshPln.CELLS(idRowPln, 7) = Me.cboxSituation.Value                  'Situation
        wshPln.CELLS(idRowPln, 8) = Me.tboxComments.Value                   'Comments
        
        wshRep.CELLS(lRowRep, 1).Value = Me.tboxItem.Value                          'Item
        wshRep.CELLS(lRowRep, 2).Value = Me.tboxStartDate                    'StartDate
        wshRep.CELLS(lRowRep, 3).Value = Me.tboxEndDate
        wshRep.CELLS(lRowRep, 4).Value = Me.cboxOwner.Value                         'Owner
        wshRep.CELLS(lRowRep, 5).Value = Me.tboxActivity.Value                      'Activity
        wshRep.CELLS(lRowRep, 6).Value = Me.tboxComments.Value                      'Comments
        wshRep.CELLS(lRowRep, 7).Value = Me.cboxStatus.Value                        'Status
        wshRep.CELLS(lRowRep, 7).Interior.ColorIndex = Me.tboxStatusKey.Value       'StatusKey
        wshRep.CELLS(lRowRep, 8).Value = Me.cboxSituation.Value                     'Situation
        wshRep.CELLS(lRowRep, 8).Interior.ColorIndex = Me.tboxSituationKey.Value    'SituationKey
        wshRep.CELLS(lRowRep, 9).Value = Me.tboxSTGoal.Value                        'STGoal
        wshRep.CELLS(lRowRep, 10).Value = Me.cboxLTGoal.Value                       'LTGoal
        wshRep.CELLS(lRowRep, 11).Value = Me.tboxExpense.Value                      'Expense
        wshRep.CELLS(lRowRep, 12).Value = Me.tboxHrsSpend.Value                     'HrsSpend
        wshRep.CELLS(lRowRep, 13).Value = Me.tboxValueAdd.Value                     'ValueAdd
        wshRep.CELLS(lRowRep, 14).Value = Me.tboxPath.Value                         'Path of the picture
'       This block is not really robust, visualizing pictures in excel rows is a bad practice
'        If picPath = "" Then
'            Exit Sub
'        Else
'            wshRep.Activate
'            PastePicture picPath, lRowRep
'        End If
        
    ' User clicked from the Worksheet 'Planning' or 'Report'
    ElseIf isMatchingIdOnPlanning = True And isMatchingIdOnReport = True Then
    ' =======================================
    ' Step 2.2. Update existing records on Planning and Report Page
    ' =======================================
        MsgBox "Numbered id on Planing WS and corresponding on Report...updating fields on both WS"
        
        idRowPln = getRowID(Me.tboxItem.Value, "Planning")
        idRowRep = getRowID(Me.tboxItem.Value, "Report")
        
        wshPln.CELLS(idRowPln, 1) = Me.tboxItem.Value                       'Item
        wshPln.CELLS(idRowPln, 2) = Me.cboxOwner.Value                      'Owner
        wshPln.CELLS(idRowPln, 3) = Me.tboxSTGoal.Value                     'STGoal
        wshPln.CELLS(idRowPln, 4) = Me.cboxLTGoal.Value                     'LTGoal
        wshPln.CELLS(idRowPln, 5) = Me.tboxActivity.Value                   'Activity
        wshPln.CELLS(idRowPln, 6) = Me.cboxStatus.Value                     'Status
        wshPln.CELLS(idRowPln, 7) = Me.cboxSituation.Value                  'Situation
        wshPln.CELLS(idRowPln, 8) = Me.tboxComments.Value                   'Comments
        
        
        wshRep.CELLS(idRowRep, 1).Value = Me.tboxItem.Value                          'Item
        wshRep.CELLS(idRowRep, 2).Value = Me.tboxStartDate                     'StartDate
        wshRep.CELLS(idRowRep, 3).Value = Me.tboxEndDate                       'EndDate
        wshRep.CELLS(idRowRep, 4).Value = Me.cboxOwner.Value                         'Owner
        wshRep.CELLS(idRowRep, 5).Value = Me.tboxActivity.Value                      'Activity
        wshRep.CELLS(idRowRep, 6).Value = Me.tboxComments.Value                      'Comments
        wshRep.CELLS(idRowRep, 7).Value = Me.cboxStatus.Value                        'Status
        wshRep.CELLS(idRowRep, 7).Interior.ColorIndex = Me.tboxStatusKey.Value       'StatusKey
        wshRep.CELLS(idRowRep, 8).Value = Me.cboxSituation.Value                     'Situation
        wshRep.CELLS(idRowRep, 8).Interior.ColorIndex = Me.tboxSituationKey.Value    'SituationKey
        wshRep.CELLS(idRowRep, 9).Value = Me.tboxSTGoal.Value                        'STGoal
        wshRep.CELLS(idRowRep, 10).Value = Me.cboxLTGoal.Value                       'LTGoal
        wshRep.CELLS(idRowRep, 11).Value = Me.tboxExpense.Value                      'Expense
        wshRep.CELLS(idRowRep, 12).Value = Me.tboxHrsSpend.Value                     'HrsSpend
        wshRep.CELLS(idRowRep, 13).Value = Me.tboxValueAdd.Value                     'ValueAdd
        wshRep.CELLS(idRowRep, 14).Value = Me.tboxPath.Value                         'Path of the picture
'       This block is not really robust, visualizing pictures in excel rows is a bad practice
'        If picPath = "" Then
'            Exit Sub
'        Else
'            wshRep.Activate
'            PastePicture picPath, lRowRep
'        End If

    ' User clicked from the Worksheet 'Report'
    ElseIf isMatchingIdOnPlanning = False And isMatchingIdOnReport = True Then
    ' =======================================
    ' Step 2.3. Save new record on the Planning Page
    ' =======================================
        MsgBox "No id on Planning WS and numbered id on Report...creating new record on Planning"
        
        idRowRep = getRowID(Me.tboxItem.Value, "Report")
        
        wshPln.CELLS(lRowPln, 1) = Me.tboxItem.Value                       'Item
        wshPln.CELLS(lRowPln, 2) = Me.cboxOwner.Value                      'Owner
        wshPln.CELLS(lRowPln, 3) = Me.tboxSTGoal.Value                     'STGoal
        wshPln.CELLS(lRowPln, 4) = Me.cboxLTGoal.Value                     'LTGoal
        wshPln.CELLS(lRowPln, 5) = Me.tboxActivity.Value                   'Activity
        wshPln.CELLS(lRowPln, 6) = Me.cboxStatus.Value                     'Status
        wshPln.CELLS(lRowPln, 7) = Me.cboxSituation.Value                  'Situation
        wshPln.CELLS(lRowPln, 8) = Me.tboxComments.Value                   'Comments
        
        wshRep.CELLS(idRowRep, 1).Value = Me.tboxItem.Value                          'Item
        wshRep.CELLS(idRowRep, 2).Value = Me.tboxStartDate                     'StartDate
        wshRep.CELLS(idRowRep, 3).Value = Me.tboxEndDate                       'EndDate
        wshRep.CELLS(idRowRep, 4).Value = Me.cboxOwner.Value                         'Owner
        wshRep.CELLS(idRowRep, 5).Value = Me.tboxActivity.Value                      'Activity
        wshRep.CELLS(idRowRep, 6).Value = Me.tboxComments.Value                      'Comments
        wshRep.CELLS(idRowRep, 7).Value = Me.cboxStatus.Value                        'Status
        wshRep.CELLS(idRowRep, 7).Interior.ColorIndex = Me.tboxStatusKey.Value       'StatusKey
        wshRep.CELLS(idRowRep, 8).Value = Me.cboxSituation.Value                     'Situation
        wshRep.CELLS(idRowRep, 8).Interior.ColorIndex = Me.tboxSituationKey.Value    'SituationKey
        wshRep.CELLS(idRowRep, 9).Value = Me.tboxSTGoal.Value                        'STGoal
        wshRep.CELLS(idRowRep, 10).Value = Me.cboxLTGoal.Value                       'LTGoal
        wshRep.CELLS(idRowRep, 11).Value = Me.tboxExpense.Value                      'Expense
        wshRep.CELLS(idRowRep, 12).Value = Me.tboxHrsSpend.Value                     'HrsSpend
        wshRep.CELLS(idRowRep, 13).Value = Me.tboxValueAdd.Value                     'ValueAdd
        wshRep.CELLS(idRowRep, 14).Value = Me.tboxPath.Value                         'Path of the picture
'       This block is not really robust, visualizing pictures in excel rows is a bad practice
'        If picPath = "" Then
'            Exit Sub
'        Else
'            wshRep.Activate
'            PastePicture picPath, lRowRep
'        End If

    End If
    

 
End Sub
' =======================================================================================================
' Command to Archive records
' =======================================================================================================
Private Sub buttonArchive_Click()
' =======================================
' Code below will archive data entries
' =======================================
' Define needed variables
Dim lRowRep As Long: Dim lRowPln As Long: Dim idRowRep As Long: Dim idRowPnl As Long
Dim iRowArc As Long
Dim Archive As Boolean: Dim isMatchingIdOnPlanning As Boolean: Dim isMatchingIdOnReport As Boolean
Dim wshRep As Worksheet: Set wshRep = Worksheets("Report")
Dim wshPln As Worksheet: Set wshPln = Worksheets("Planning")
Dim wshArc As Worksheet: Set wshArc = Worksheets(".Archive")
Dim Message As String
'
idRowPln = getRowID(Me.tboxItem.Value, "Planning") 'row id on Planning
idRowRep = getRowID(Me.tboxItem.Value, "Report")   'row id on Report

    ' =======================================
    ' 1. code below will check position of radio buttons
    ' =======================================
    If (Me.optionNo.Value = True) Then
        Archive = False
        ' Abort arching!
        Me.optionNo.SetFocus
            MsgBox "Selected option 'No' to not Archive this record! Select option 'Yes' and press button 'Archive' again!"
        Exit Sub
    End If
    
    If (Me.optionYes.Value = True) Then
        Archive = True
    End If
    
Message = MsgBox("Archiving these records!", vbYesNoCancel)

If Message = vbNo Then
    MsgBox "Abort Archiving"
    
ElseIf Message = vbYes And Archive = True And (idRowRep = 0 Or idRowPln = 0) Then
    
    MsgBox "Attempt to archive not complete record! Aborting Archive as record is not present on both Planning and Report worksheets!"
    Exit Sub
    
ElseIf Message = vbYes And Archive = True Then

    '' Archiving
    
    ' Find where to paste this record on the .Archive WS?
    lRowArc = getLastEmptyRow(".Archive")
    
    ' Cut all data from worksheet Report
    wshRep.Range("A" & idRowRep & ":O" & idRowRep).Cut wshArc.Range("A" & lRowArc & ":O" & lRowArc)
    ' Delete all data from worksheet Planning!
    wshPln.Range("A" & idRowPln & ":H" & idRowPln).Delete Shift:=xlUp
    wshRep.Range("A" & idRowRep & ":O" & idRowRep).Delete Shift:=xlUp
    
    ' Remove the row
    MsgBox "Records are Archived and inputs are deleted on both Worksheets!"


End If

End Sub
