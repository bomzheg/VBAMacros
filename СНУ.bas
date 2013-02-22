Attribute VB_Name = "СНУ"
Option Base 1
Option Explicit
Option Private Module
Sub СНУ()
Dim list$, x!(2), fx!(2), mJ!(2, 2), Job!(2, 2), eps!, b!(2, 1), i%, k%, fx2!(2, 1), B2!(2), j%
list = InputBox("Введите название листа:", "Выполнение макроса.", "сну", 5000, 5000)
If Not [Table].IsWorkSheetExist(list) Then
    MsgBox ("Нет листа " + list)
    Exit Sub
End If
eps = Worksheets(list).Cells(4, 3)
Call [Matrix].Ввод_Вектора(list, x, 1, 3, 0, 2)
j = 0
Do
fx(1) = f1(x)
fx(2) = f2(x)
For i = 1 To 2
fx2(i, 1) = fx(i)
Next i
mJ(1, 1) = d1f1(x)
mJ(1, 2) = d2f1(x)
mJ(2, 1) = d1f2(x)
mJ(2, 2) = d2f2(x)

Call [Matrix].обращение(mJ, Job)
Call [Matrix].Матр_произв(Job, fx2, b)
Call [Matrix].Вывод_Вектора(list, x, 4 + j, 6, 0)
Call [Matrix].Вывод_Вектора(list, fx, 4 + j, 7, 0)
Call [Matrix].Вывод_Матрицы(list, mJ, 4 + j, 7)
Call [Matrix].Вывод_Матрицы(list, Job, 4 + j, 9)
For i = 1 To 2
B2(i) = b(i, 1)
Next i
Call [Matrix].Вывод_Вектора(list, B2, 4 + j, 12, 0)
For k = 1 To 2
    x(k) = x(k) - B2(k)
Next k
Call [Matrix].Вывод_Вектора(list, x, 4 + j, 13, 0)
j = j + 2
Loop While i < 50 And ([Matrix].Норма_вектора(B2)) > eps
For i = j + 1 To j + 6
For k = 1 To 8
Worksheets(list).Cells(4 + i, 5 + k) = Empty
Next k
Next i
End Sub
Private Function f1!(x!())
f1 = 28.3 ^ 2 / Cells(6, 3) ^ 2 * x(1)
End Function
Private Function f2!(x!())
f2 = (0.16 - (x(1) ^ 2) ^ (1 / 2)) + x(2)
End Function
Private Function d1f1!(x!())
d1f1 = Cos(x(1) + 1.5)
End Function
Private Function d2f1!(x!())
d2f1 = -1
End Function
Private Function d1f2!(x!())
d1f2 = (-1 * x(1)) / (Sqr(0.16 - x(1) ^ 2))
End Function
Private Function d2f2!(x!())
d2f2 = 1
End Function






















