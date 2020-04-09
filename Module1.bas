Attribute VB_Name = "Module1"
Option Explicit

Dim wb As Workbook
Dim wsData As Worksheet, wsPT As Worksheet

Sub Create_Pivot_Table()
Dim LastRow As Long, LastColumn As Long
Dim DataRange As Range
Dim PTCache As PivotCache
Dim PT As PivotTable

'On Error GoTo errHandle

Set wb = ThisWorkbook
Set wsData = wb.Worksheets("Data")

'// Delete Pivot Table sheet
Call Delete_PT_Sheet

'// Create Data Range variable
With wsData
LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
LastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column

    Set DataRange = .Range(.Cells(1, 1), .Cells(LastRow, LastColumn))
End With

'// Create Pivot Table Worksheet
Set wsPT = wb.Worksheets.Add
wsPT.Name = "Pivot Table"

'// Storing Pivot Table Cache
Set PTCache = wb.PivotCaches.Create(xlDatabase, DataRange)

'// Create Pivot Table
Set PT = PTCache.CreatePivotTable(wsPT.Range("B5"), "PT_SalesSummary")


'// Adding Columns, Rows and Data to pivot table
With PT

    '// Pivot Table Layout
    .RowAxisLayout xlTabularRow
    .ColumnGrand = False 'Optional (Column Grand Total)
    .RowGrand = False 'Optional (Row Grand Total)
    
    .TableStyle2 = "PivotStyleMedium9"
    .HasAutoFormat = False 'Re-Format Pivot Table when refresh
    .SubtotalLocation xlAtTop 'Position SubTotal on the top or bottom
    
    ' Filters
    With .PivotFields("Retailer Country")
        .Orientation = xlPageField
        .EnableMultiplePageItems = True 'Allow multi-selection
    End With
    
    ' Row Section (Layer 1)
    With .PivotFields("Order method type")
        .Orientation = xlRowField
        .Position = 1
        .LayoutBlankLine = False 'True if a blank row is inserted after the specified row field in a PivotTable report.
                                'The default value is False. Read/write Boolean.
        .Subtotals(1) = False
        .LayoutForm = xlTabular 'xlOutline
        .LayoutCompactRow = True
        .Subtotals(1) = True
    End With
    
    ' Row Section (Layer 2)
    With .PivotFields("Product line")
        .Orientation = xlRowField
        .LayoutBlankLine = True
    End With
    
    ' Column Section
    With .PivotFields("Year")
        .Orientation = xlColumnField
        .Position = 1
    End With
    
    
    ' Values section
    With .PivotFields("Revenue")
        .Orientation = xlDataField
        .Position = 1
        .Function = xlSum
        .NumberFormat = "$#,##;($#,##);-"
        .Caption = "Revenue Total"
    End With
    
    With .PivotFields("Revenue")
        .Orientation = xlDataField
        .Position = 1
        .Function = xlAverage
        .NumberFormat = "$#,##;($#,##);-"
        .Caption = "Revenue (Average)"
    End With
    
    With .PivotFields("Quantity")
        .Orientation = xlDataField
        .Position = 1
        .Function = xlCount
        .NumberFormat = "#,##;(#,##);-"
        .Caption = "QTY"
    End With
    
End With

'// Resize Column Width
wsPT.Cells.EntireColumn.AutoFit

ClearObjects:
Set PTCache = Nothing
Set PT = Nothing
Set DataRange = Nothing

Call Clear_Objects

Exit Sub

errHandle:
MsgBox "Error: " & Err.Description, vbExclamation
GoTo ClearObjects

End Sub

Private Sub Delete_PT_Sheet()

On Error Resume Next
Application.DisplayAlerts = False
wb.Worksheets("Pivot Table").Delete
Application.DisplayAlerts = True

End Sub

Private Sub Clear_Objects()

'// Release Memory
Set wsData = Nothing
Set wb = Nothing

End Sub











