Attribute VB_Name = "ShiftLongFn"
' VBA function to shift longitudes in conformance with Canback MAp Projection

Function ShiftLong(shift As Integer)

  If shift < 0 Or shift > 8 Then
ShiftLong = "Shift out of range"
GoTo endnow
End If

If shift = 0 Then ShiftLong = -20 Else
If shift = 1 Then ShiftLong = 0 Else
If shift = 2 Then ShiftLong = 30 Else
If shift = 3 Then ShiftLong = -7 Else
If shift = 4 Then ShiftLong = -20 Else
If shift = 5 Then ShiftLong = -70 Else
If shift = 6 Then ShiftLong = 35 Else
If shift = 7 Then ShiftLong = -33 Else
If shift = 8 Then ShiftLong = -6.6

endnow:
End Function
