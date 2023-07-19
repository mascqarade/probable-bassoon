Public Sub lab8_1()
s = LCase(InputBox("Введите строку"))
For i = 1 To Len(s)
    If Mid(s, i, 2) = "ма" Then Mid(s, i, 2) = "ле"
Next
MsgBox s
End Sub

Public Sub lab8_2()
c = 0
s = LCase(InputBox("Введите строку"))
For i = 1 To Len(s)
    If Mid(s, i, 3) = "кот" Then c = c + 1
Next
MsgBox c
End Sub

Public Sub lab8_3()
s = InputBox("Введите строку")
pr = 0
Dim ost As String
For i = 1 To Len(s)
    If Mid(s, i, 1) = " " Then
        pr = pr + 1
    End If
Next
k = 0
Dim nach As String
Dim kon As String
For i = 1 To Len(s)
    If Mid(s, i, 1) = " " Then
        k = k + 1
        If k > 0 Then
            ost = ost + Mid(s, i, 1)
        End If
    ElseIf k = 0 Then
        nach = nach + Mid(s, i, 1)
    ElseIf k = pr Then
        kon = kon + Mid(s, i, 1)
    Else
        ost = ost + Mid(s, i, 1)
    End If
Next
MsgBox (kon + ost + nach)
End Sub

Public Sub lab8_4()

Dim s, c As String
s = InputBox("Введите строку")
s = LCase(s)
c = StrReverse(s)
MsgBox c

End Sub
