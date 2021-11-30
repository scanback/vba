Attribute VB_Name = "ExpandPivot"
'The subroutine adds columns to a Pivot table when there are too many columns to add manually
'Add one column, put the cursor in it and then run the subroutine
'Does not eork if the pivot table is in the data model
'Found on the web https://www.extendoffice.com/documents/excel/2246-excel-pivot-table-add-multiple-fields.html

Sub AddAllFieldsValues()
Dim pt As PivotTable
Dim iCol As Long
Dim iColEnd As Long

Set pt = ActiveSheet.PivotTables(1)

With pt
   iCol = 1
   iColEnd = .PivotFields.Count

    For iCol = 1 To iColEnd
        With .PivotFields(iCol)
          If .Orientation = 0 Then
              .Orientation = xlDataField
          End If
        End With
    Next iCol
End With

End Sub
