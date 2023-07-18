Public Sub lab3_1()
Max = -10000000000#
For i = 1 To 20
    n = CSng(InputBox(i))
    If n > Max Then Max = n
Next
MsgBox "Максимальный из отрицательных: " & Max
End Sub

Public Sub lab3_2()
n = 1
ans = Sin(Tan(n))
While ans > 0
    n = n + 1
    ans = Sin(Tan(n))
Wend
MsgBox ans
End Sub

Public Sub lab3_3()
ans = 0
For x = -15 To 10
    y = Sin(x) + 4 * Cos(x - 2)
    If y > 0 Then ans = ans + y
Next
End Sub

Public Sub lab3_4()
k = 0
n = CInt(InputBox("n"))
While n <> 0
    If n Mod 3 = 0 Then k = k + 1
    n = CInt(InputBox("n"))
Wend
End Sub