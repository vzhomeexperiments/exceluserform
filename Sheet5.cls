VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' ==================================================================================================
' Purpose of this code is to invoke the User Form when user double click the line on the spreadsheet
' ==================================================================================================

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

' Define variables needed
Dim iRow As Integer: Dim itemID As String
' Initialize variables
iRow = Target.Row
itemID = CELLS(Target.Row, 1).Value

' Add fail safe to insure numeric id is present
If itemID = "" Then
    MsgBox "Error, Add numeric and unique id in the column A!"
    Exit Sub
End If

' Update Inputs in the form
UpdateInputs iRow

' Cancel = false means form can not be modified together with cell
Cancel = True

'possible to modify cell when the form is up
PlanningForm.Show False

End Sub



