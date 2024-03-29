'This Excel VBA sub fills the highlighted cells with the RGB color code in each cell, expressed in hex (#xxyyzz)
' The main use is to pre-display shades of color for shapefiles

Sub CellColor()
    Dim cell As Object
    Dim ColorCell As String

    Dim count As Integer
    count = 0
    For Each cell In Selection
        count = count + 1
        ColorCell = cell.Value
'MsgBox cell
reh = Mid(ColorCell, 2, 2)
grh = Mid(ColorCell, 4, 2)
blh = Mid(ColorCell, 6, 2)

red = CInt("&H" & reh)
grd = CLng("&H" & grh)
bld = CLng("&H" & blh)

'MsgBox red & " " & grd & " " & bld
        
        
        cell.Interior.Color = RGB(red, grd, bld)
        
    Next cell
    'MsgBox count & " item(s) selected"
End Sub