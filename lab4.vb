
Public Sub lab4_1()
ans = 0
For i = 3 To 99 Step 3
    ans = ans + i
    ans = Sqr(ans)
Next
MsgBox ans
End Sub

Public Sub lab4_3()
k = CInt(InputBox("k"))
x = CInt(InputBox("x"))
yi = 5 / 6
y = yi
For i = 1 To k
    yi = yi * ((5 * x) / (i + 2))
    y = y + yi
Next
MsgBox y
End Sub

Public Sub lab4_4()
For i = 1 To 20
    k = 0
    n = 0
    a = CInt(InputBox(a))
    While n < a
        k = k + 1
        n = 3 ^ k
    Wend
    k = k - 1
    Cells(i, 1) = a
    Cells(i, 2) = k
Next
End Sub
