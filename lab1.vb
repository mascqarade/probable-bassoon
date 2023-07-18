Public Sub lab1_1()
Dim A, B, C, x, z As Single
A = Cells(1, 1)
B = Cells(2, 1)
C = Cells(3, 1)
x = CSng(InputBox("x"))
y = CSng(InputBox("y"))
y1 = ((A - 3 * B) / (A + B)) * (A + C) * Cos(x) ^ 2 - z
y2 = (x * A / B) * Sqr(x ^ (C) + 1) * (1 / (A ^ (x) + 1))
MsgBox ("a) " + CStr(y1))
MsgBox ("b) " + CStr(y2))
End Sub

Public Sub lab1_2()
x = CSng(InputBox("Длина ребра"))
V = x ^ 3
S = 6 * (x ^ 2)
MsgBox ("Объем: " + CStr(V))
MsgBox ("Площадь поверхности: " + CStr(S))
End Sub

Public Sub lab1_3()
x = CSng(InputBox("Первое число"))
y = CSng(InputBox("Второе число"))
z = CSng(InputBox("Третье число"))
MsgBox ("Среднее арифметическое: " + CStr((x + y + z) / 3))
MsgBox ("Среднее геометрическое: " + CStr((x * y * z) ^ (1 / 3)))
End Sub

Public Sub lab1_4()
A = CSng(InputBox("A"))
S = CSng(InputBox("S"))
h = (-A + Sqr(A ^ (2) + 8 * S)) / 2
MsgBox h
End Sub