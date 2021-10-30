Attribute VB_Name = "ShiftLongFn"
' VBA function to shift longitudes in conformance with Canback MAp Projection

Function ShiftLong(shift As Integer)
Dim x As Integer
x = shift
If x < 0 Or x > 8 Then
ShiftLong = "Shift out of range"
GoTo endnow
End If

If x = 0 Then ShiftLong = -20 Else
If x = 1 Then ShiftLong = 10 Else
If x = 2 Then ShiftLong = 30 Else
If x = 3 Then ShiftLong = -7 Else
If x = 4 Then ShiftLong = -20 Else
If x = 5 Then ShiftLong = -70 Else
If x = 6 Then ShiftLong = 35 Else
If x = 7 Then ShiftLong = -33 Else
If x = 8 Then ShiftLong = -6.6 'Nuuk

endnow:
End Function
