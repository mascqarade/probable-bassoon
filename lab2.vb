Public Sub lab2_1()
Dim A, B, C As Single
A = CSng(InputBox("A"))
B = CSng(InputBox("B"))
C = CSng(InputBox("C"))
Cells(2, 2) = A
Cells(3, 2) = B
Cells(4, 2) = C
A = Cells(2, 2)
B = Cells(3, 2)
C = Cells(4, 2)
If A >= B Then
    If B >= C Then
        MsgBox (CStr(A) + " " + CStr(B) + " " + CStr(C))
    ElseIf A >= C Then
        MsgBox (CStr(A) + " " + CStr(C) + " " + CStr(B))
    Else
        MsgBox (CStr(C) + " " + CStr(A) + " " + CStr(B))
    End If
Else
    If A >= C Then
        MsgBox (CStr(B) + " " + CStr(A) + " " + CStr(C))
    ElseIf C >= B Then
        MsgBox (CStr(C) + " " + CStr(B) + " " + CStr(A))
    Else
        MsgBox (CStr(B) + " " + CStr(C) + " " + CStr(A))
    End If
End If
End Sub

Public Sub lab2_2()
x = CSng(InputBox("x"))
y = CSng(InputBox("y"))
z = CSng(InputBox("z"))
If x + y > z Then
    If x = z Or y = z Or x = y Then
        MsgBox ("Треугольник существует. Равнобедренный")
    Else
        MsgBox ("Треугольник существует. Неравнобедренный")
    End If
Else
    MsgBox ("Треугольник не существует")
End If
End Sub

Public Sub lab2_3()
k = CInt(InputBox("Номер клиента"))
n = Cells(k, 1)
If n < 5000 Then
    MsgBox "Налог:" & n * 0.13
ElseIf n < 40000 Then
    MsgBox "Налог: " & n * 0.2
Else
    MsgBox "Налог: " & n * 0.3
End If
End Sub

Public Sub lab2_4()
Dim Min As Single
x = CSng(InputBox("x"))
y = CSng(InputBox("y"))
z = CSng(InputBox("z"))
If Abs(x) < Abs(y) Then
    If Abs(x) < Abs(z) Then
        Min = x
    Else: Min = z
    End If
Else
    If Abs(y) < Abs(z) Then
        Min = y
    Else: Min = z
    End If
End If
MsgBox "Минимальный по модулю: " & Min
End Sub