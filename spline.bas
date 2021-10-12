'https://www.business-spreadsheets.com/forum.asp?t=120

Function SPLINE(periodcol As Range, ratecol As Range, x As Range)

Dim period_count As Integer
Dim rate_count As Integer

period_count = periodcol.Rows.Count
rate_count = ratecol.Rows.Count

If period_count = 1 Then period_count = periodcol.Columns.Count
If rate_count = 1 Then rate_count = ratecol.Columns.Count

If period_count <> rate_count Then
    SPLINE = "Error: Range count dos not match"
    GoTo endnow
End If
 
ReDim xin(period_count) As Double
ReDim yin(period_count) As Double

Dim c As Integer

For c = 1 To period_count
xin(c) = periodcol(c)
yin(c) = ratecol(c)
Next c

Dim n As Integer
Dim i, k As Integer
Dim p, qn, sig, un As Double
ReDim u(period_count - 1) As Double
ReDim yt(period_count) As Double

n = period_count
yt(1) = 0
u(1) = 0

For i = 2 To n - 1
    sig = (xin(i) - xin(i - 1)) / (xin(i + 1) - xin(i - 1))
    p = sig * yt(i - 1) + 2
    yt(i) = (sig - 1) / p
    u(i) = (yin(i + 1) - yin(i)) / (xin(i + 1) - xin(i)) - (yin(i) - yin(i - 1)) / (xin(i) - xin(i - 1))
    u(i) = (6 * u(i) / (xin(i + 1) - xin(i - 1)) - sig * u(i - 1)) / p
   
    Next i
   
qn = 0
un = 0

yt(n) = (un - qn * u(n - 1)) / (qn * yt(n - 1) + 1)

For k = n - 1 To 1 Step -1
    yt(k) = yt(k) * yt(k + 1) + u(k)
Next k

Dim klo, khi As Integer
Dim h, b, a As Double

klo = 1
khi = n
Do
k = khi - klo
If xin(k) > x Then
khi = k
Else
klo = k
End If
k = khi - klo
Loop While k > 1
h = xin(khi) - xin(klo)
a = (xin(khi) - x) / h
b = (x - xin(klo)) / h
y = a * yin(klo) + b * yin(khi) + ((a ^ 3 - a) * yt(klo) + (b ^ 3 - b) * yt(khi)) * (h ^ 2) / 6


SPLINE = y

endnow:
End Function
