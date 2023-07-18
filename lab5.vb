Public Sub lab5_1()
n = CInt(InputBox("N"))
ReDim a(n) As Integer
Min = 101
k1 = 0
k2 = 0
For I = 1 To n
    a(I) = Rnd * 100 + Rnd * (-100)
    If a(I) < Min Then Min = a(I): k1 = I: k2 = I
    If a(I) = Min Then k2 = I
Next
MsgBox k1 & " " & k2
End Sub

Public Sub lab5_2()
Dim a(10) As Integer
Min = 15
For I = 1 To 10
    a(I) = Rnd * 20 + Rnd * (-20)
    Cells(I, 1) = a(I)
    If a(I) > 2 And a(I) < 14 And a(I) < Min Then Min = a(I)
Next
MsgBox Min
End Sub

Public Sub lab5_3()
Dim a(15) As Integer
n = 0
For I = 1 To 15
    a(I) = Rnd * 20 + Rnd * (-20)
    If a(I) < 1 Then n = n + 1
Next
ReDim B(n) As Integer
j = 1
For I = 1 To 15
    If a(I) < 1 Then B(j) = a(I): j = j + 1
Next
End Sub

Public Sub lab5_4()
B = CInt(InputBox("B"))
flag = 0
k = 0
Max = 0
Dim a(10) As Integer
For I = 1 To 10
    a(I) = Rnd * 20
    If a(I) > B And a(I) > Max Then Max = a(I): flag = 1: k = I
Next
If flag = 0 Then
    MsgBox "00"
Else
    MsgBox Max & " " & k
End If
End Sub

