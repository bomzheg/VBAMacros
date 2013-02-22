Attribute VB_Name = "нелин_ур"
Option Base 1
Option Explicit
Option Private Module
Sub nelin_ur()
Dim i%, x!, dx!, n%, buf!, fx!, A!, b!, x1!, x2!, fa!, fb!, eps!, B2!, list$
list = InputBox("Введите название листа:", "Выполнение макроса.", "НУ", 5000, 5000)
If Not [Table].IsWorkSheetExist(list) Then
    MsgBox ("Нет листа " + list)
    Exit Sub
End If
n = Worksheets(list).Cells(1, 3)
x1 = Worksheets(list).Cells(2, 3)
x2 = Worksheets(list).Cells(3, 3)
dx = Abs(x1 - x2) / n
x = x1
buf = x1
B2 = f(x)
eps = Worksheets(list).Cells(4, 3)
For i = 1 To n
    x = x + dx
    Worksheets(list).Cells(2, 6 + i) = x
    fx = f(x)
    Worksheets(list).Cells(3, 6 + i) = fx
    If B2 * fx < 0 Then
      A = buf
      b = x
      Exit For
    Else
      buf = x
      B2 = f(x)
    End If
Next i
i = 0
Do
    i = i + 1
    x = (b + A) / 2
    fa = f(A)
    fb = f(b)
    fx = f(x)
    Worksheets(list).Cells(5 + i, 6) = A
    Worksheets(list).Cells(5 + i, 7) = x
    Worksheets(list).Cells(5 + i, 8) = b
    Worksheets(list).Cells(5 + i, 9) = Abs(b - A)
    Worksheets(list).Cells(5 + i, 10) = fa
    
    Worksheets(list).Cells(5 + i, 11) = fx
    Worksheets(list).Cells(5 + i, 12) = fb
    If fa * fx < 0 Then
        b = x
    Else: A = x
    End If
    Range(Worksheets(list).Cells(6 + i, 6), Worksheets(list).Cells(6 + i + 3, 12)) = Empty
Loop While Abs(b - A) > eps
End Sub
Private Function f!(x!)
f = (1 / (x + 1.5)) - x - 1
End Function
Private Function ff!(x!)
ff = Cos(x) + x - 0.5
End Function
Private Function f1!(x!)
f1 = -Sin(x) + 1
End Function
Sub Нелин_ур_произ()
Dim i%, fx!, fx1!, h!, x!, eps!, list$, j%
list = InputBox("Введите название листа:", "Выполнение макроса.", "ну", 5000, 5000)
If Not [Table].IsWorkSheetExist(list) Then
    MsgBox ("Нет листа " + list)
    Exit Sub
End If
x = Worksheets(list).Cells(2, 3)
eps = Worksheets(list).Cells(4, 3)
i = 1
Do
Worksheets(list).Cells(10 + i, 19) = x
fx = ff(x)
fx1 = f1(x)
h = fx / fx1
x = x - h
Worksheets(list).Cells(10 + i, 20) = fx
Worksheets(list).Cells(10 + i, 21) = fx1
Worksheets(list).Cells(10 + i, 22) = h
Worksheets(list).Cells(10 + i, 23) = x
i = i + 1
Loop While Abs(fx) > eps And i < 1000 ' или or?
For i = i To i + 7
For j = 1 To 5
Worksheets(list).Cells(10 + i, 18 + j) = Empty
Next j
Next i
End Sub
