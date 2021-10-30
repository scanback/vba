Attribute VB_Name = "ShiftLatFn"
'VBA function to shift latitudes in conformance with Canback Map Projection

Function ShiftLat(shift As Integer)

  If shift < 0 Or shift > 9 Then
ShiftLat = "Shift out of range"
GoTo endnow
End If

If shift = 0 Then ShiftLat = 0 Else
If shift = 1 Then ShiftLat = 0 Else
If shift = 2 Then ShiftLat = 0 Else
If shift = 3 Then ShiftLat = -4.5 Else
If shift = 4 Then ShiftLat = -5 Else
If shift = 5 Then ShiftLat = 0 Else
If shift = 6 Then ShiftLat = 0 Else
If shift = 7 Then ShiftLat = 0 Else
If shift = 8 Then ShiftLat = -2.2

endnow:
End Function

