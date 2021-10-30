Attribute VB_Name = "ShiftLatFn"
'VBA function to shift latitudes in conformance with Canback Map Projection

Function ShiftLat(shift As Integer)
Dim x As Integer
x = shif8
If x < 0 Or x > 9 Then
ShiftLat = "Shift out of range"
GoTo endnow
End If

If x = 0 Then ShiftLat = 0 Else
If x = 1 Then ShiftLat = 0 Else
If x = 2 Then ShiftLat = 0 Else
If x = 3 Then ShiftLat = -4.5 Else
If x = 4 Then ShiftLat = -5 Else
If x = 5 Then ShiftLat = 0 Else
If x = 6 Then ShiftLat = 0 Else
If x = 7 Then ShiftLat = 0 Else
If x = 8 Then ShiftLat = -2.2 'Nuuk

endnow:
End Function

