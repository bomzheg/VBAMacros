Attribute VB_Name = "Matrix"
Option Base 1
Option Explicit
Option Private Module
Sub Произведение_вектора_на_число(Вектор!(), Число!, Результат!(), Optional n%)
Dim i%
If n = 0 Then
    n = UBound(Вектор, 1)
End If
For i = 1 To n
    Результат(i) = Вектор(i) * Число
Next
End Sub
Function Норма_вектора!(A!(), Optional Число_строк%)
Dim S!, i%
S = 0
If Число_строк = 0 Then
Число_строк = UBound(A, 1)
End If
For i = 1 To Число_строк
    S = S + A(i) ^ 2
Next i
Норма_вектора = Sqr(S)
End Function
Sub Транспонирование(A!(), At!(), Optional Число_строк%, Optional Число_столбцов%)
Dim i%, j%
If Число_строк = 0 And Число_столбцов = 0 Then
Число_строк = UBound(A, 1)
Число_столбцов = UBound(A, 2)
ReDim At(Число_столбцов, Число_строк)
End If
For i = 1 To Число_строк
    For j = 1 To Число_столбцов
        At(j, i) = A(i, j)
    Next j
Next i
End Sub
Function норма_матрицы!(A!(), Optional Число_строк%, Optional Число_столбцов%)
Dim i%, j%, S!
If Число_строк = 0 And Число_столбцов = 0 Then
Число_строк = UBound(A, 1)
Число_столбцов = UBound(A, 2)
End If
S = 0
For i = 1 To Число_строк
    For j = 1 To Число_столбцов
      S = S + A(i, j) ^ 2
    Next j
Next i
норма_матрицы = Sqr(S)
End Function
Sub векторн_сумма(Число_строк%, A!(), b!(), R!())
Dim i%
For i = 1 To Число_строк
R(i) = A(i) + b(i)
Next i
End Sub
Sub сумма_матриц(A!(), b!(), c!(), Optional Число_строк%, Optional Число_столбцов%)
Dim i%, j%
If Число_строк = 0 And Число_столбцов = 0 Then
Число_строк = UBound(A, 1)
Число_столбцов = UBound(A, 2)
End If
ReDim c(Число_строк, Число_столбцов)
For i = 1 To Число_строк
    For j = 1 To Число_столбцов
      c(i, j) = A(i, j) + b(i, j)
    Next j
Next i
End Sub
Sub Разность_матриц(A!(), b!(), c!(), Optional Число_строк%, Optional Число_столбцов%)
If Число_строк = 0 And Число_столбцов = 0 Then
Число_строк = UBound(A, 1)
Число_столбцов = UBound(A, 2)
End If
Dim i%, j%
For i = 1 To Число_строк
    For j = 1 To Число_столбцов
      c(i, j) = A(i, j) - b(i, j)
    Next j
Next i
End Sub
Sub Разность_векторов(A!(), b!(), c!(), Optional Число_элементов%)
If Число_элементов = 0 Then
Число_элементов = UBound(A, 1)
End If
Dim i%
For i = 1 To Число_элементов
      c(i) = A(i) - b(i)
Next i
End Sub
Sub Матр_произв(A!(), b!(), c!(), Optional Число_строкA%, Optional Число_столбцовA%, Optional Число_столбцовB%)
Dim i%, j%, S!, l%
If Число_строкA = 0 And Число_столбцовA = 0 And Число_столбцовB = 0 Then
Число_строкA = UBound(A, 1)
Число_столбцовA = UBound(A, 2)
Число_столбцовB = UBound(b, 2)
End If
ReDim c(Число_строкA, Число_столбцовB)
For i = 1 To Число_строкA
    For j = 1 To Число_столбцовB
        S = 0
        For l = 1 To Число_столбцовA
            S = S + A(i, l) * b(l, j)
        Next l
        c(i, j) = S
   Next j
Next i
End Sub
Sub обращение(A!(), d!())
Dim i%, j%, k%, f!(), S!, Число_строк%
Число_строк = UBound(A, 1)
ReDim f(Число_строк, 2 * Число_строк)
For i = 1 To Число_строк
    For j = 1 To 2 * Число_строк
        If j <= Число_строк Then
        f(i, j) = A(i, j)
        Else
        If i = j - Число_строк Then
        f(i, j) = 1
        Else
        f(i, j) = 0
        End If
        End If
    Next j
Next i
For k = 1 To Число_строк
  S = f(k, k)
  For j = k To 2 * Число_строк
    f(k, j) = f(k, j) / S
  Next j
  For i = 1 To Число_строк
      If i <> k Then
         S = f(i, k)
         For j = k To 2 * Число_строк
            f(i, j) = f(i, j) - f(k, j) * S
         Next j
      End If
  Next i
Next k
For i = 1 To Число_строк
    For j = 1 To Число_строк
     d(i, j) = f(i, j + Число_строк)
    Next j
Next i
End Sub
Sub swap(c!, d!)
Dim buf!
buf = c
c = d
d = buf
End Sub
Function Определитель(A!())
Dim i%, j%, n%, opr!
n = UBound(A, 1)
opr = 0
For j = 1 To n
opr = opr + A(1, j) * (-1) ^ (1 + j) * det(A(), j)
Определитель = opr
Next j
End Function
Function det(A!(), k%)
 Dim b!(), i%, j%, Nstr%      'перевести из стадии "блин, почему не компилится?" в "блин, почему не работает?" - готово
 Nstr = UBound(A, 1)            ' и всёже почему не работает? - тоже готово - прога работает
 ReDim b(Nstr - 1, Nstr - 1)
 det = 0
 If k = 1 Then
 For j = 2 To Nstr
    For i = 2 To Nstr
     b(i - 1, j - 1) = A(i, j)
    Next i
 Next j
 Else
 If k = Nstr Then
  For i = 2 To Nstr
    For j = 1 To Nstr - 1
     b(i - 1, j) = A(i, j)
    Next j
  Next i
 Else
 For i = 2 To Nstr
    For j = 1 To k - 1
     b(i - 1, j) = A(i, j)
    Next j
    For j = k + 1 To Nstr
     b(i - 1, j - 1) = A(i, j)
    Next j
 Next i
 End If
 End If
 Nstr = Nstr - 1
 If Nstr = 3 Then
       det = Det3_3(b)
       Exit Function
       Else
       If Nstr > 3 Then
 For j = 1 To Nstr
       det = det + (-1) ^ (1 + j) * b(1, j) * det(b(), j)
 Next j
       Else: det = b(1, 1) * b(2, 2) - b(1, 2) * b(2, 1)
       End If
 End If
End Function
Sub SLAU_2_matrix(МатрицаКоэффициентов!(), ВекторСвободных!(), x!())
Dim n%, i%, j%, c!()
n = UBound(МатрицаКоэффициентов, 1)
ReDim c(n, n + 1)
For i = 1 To n
    For j = 1 To n
        c(i, j) = МатрицаКоэффициентов(i, j)
    Next j
    c(i, j) = ВекторСвободных(i)
Next i
Call СЛАУ(c, x)
End Sub
Sub СЛАУ(c!(), x!(), Optional n%)
Dim i%, j%, k%, bet!, S!
n = UBound(c, 1)
ReDim x(n)
For i = 1 To n
    For j = 1 To n + 1
     If i = j And c(i, j) = 0 Then
        MsgBox (" ошибка, диагональный элемент - нулевой. выполнение макроса будет прервано")
    Exit Sub
    End If
    Next j
Next i
For k = 1 To n - 1
    For i = k + 1 To n
        bet = -c(i, k) / c(k, k)
        For j = k To n + 1
            c(i, j) = c(i, j) + bet * c(k, j)
        Next j
    Next i
Next k
x(n) = c(n, n + 1) / c(n, n)
For i = n - 1 To 1 Step -1
    S = 0
    For j = i + 1 To n
      S = S + c(i, j) * x(j)
    Next j
    x(i) = (c(i, n + 1) - S) / c(i, i)
Next i
End Sub
Sub SLAU_iter(Matrix!(), vector!(), Otvet!(), Epsylon!)
Dim i%, j%, n%, m%, c!(), d!(), xk1!(), xk!(), temp!(), max!, buf!, an%
n = UBound(vector, 1)
an = SLAU_analiz
'Select Case an
'    Case 1
'        MsgBox ("неоднородная система, нет смысла решать")
ReDim c(n, n), d(n, 1), xk(n, 1), xk1(n, 1), temp(n, 1), Otvet(n)
For i = 1 To n
    For j = 1 To n
        If i <> j Then
            c(i, j) = -Matrix(i, j) / Matrix(i, i)
        Else
            c(i, j) = 0
        End If
    Next j
    d(i, 1) = vector(i) / Matrix(i, i)
Next i
If норма_матрицы(c) >= 1 Then
      i = MsgBox("Условие сходимости не выполняется =(", vbOKOnly Or vbCritical, "Решение слау")
      Exit Sub
End If
Call Матр_произв(c, d, temp)
Call сумма_матриц(temp, d, xk1)
j = 0
Do
    For i = 1 To n
        xk(i, 1) = xk1(i, 1)
    Next i
    Call Матр_произв(c, xk, temp)
    Call сумма_матриц(temp, d, xk1)
    j = j + 1
    max = Abs(xk(1, 1) - xk1(1, 1))
    For i = 2 To n
        buf = Abs(xk(i, 1) - xk1(i, 1))
        If buf > max Then max = buf
    Next i
Loop While max > Epsylon Or j < 100
For i = 1 To n
    Otvet(i) = xk1(i, 1)
Next i
End Sub
Function SLAU_analiz%(Matrix!(), vector!())
Dim i%, j%, n%, norm!, det!
norm = Норма_вектора(vector)
det = Определитель(Matrix)
If Abs(det) < 0.00001 Then
    SLAU_analiz = SLAU_analiz + 1 ' - несовместная система
End If
If norm > 0.0001 Then
    SLAU_analiz = SLAU_analiz + 10 '- однородная система
End If
End Function
Sub Ввод_Матрицы(list$, A!(), ном_строки%, ном_столбца%, Optional n%, Optional m%)  '
Dim i%, j%
If n = 0 And m = 0 Then
    Call Размерность_матрицы(list, ном_строки, ном_столбца, n, m)
End If
ReDim A(n, m)  ' а точно ли эта строка не перед end if?
For i = 1 To n
    For j = 1 To m
        A(i, j) = Worksheets(list).Cells(i + ном_строки - 1, j + ном_столбца - 1)
    Next j
Next i
End Sub
Sub Вывод_Матрицы(list$, A$(), ном_строки%, ном_столбца%)
Dim i%, j%, n%, m%
n = UBound(A, 1)
m = UBound(A, 2)
For i = 1 To n
    For j = 1 To m
       Worksheets(list).Cells(i + ном_строки - 1, j + ном_столбца - 1) = A(i, j)
    Next j
Next i
End Sub
Sub Ввод_Вектора(list$, A!(), ном_строки%, ном_столбца%, из_строки_вопр%, Optional n%)
Dim i%
If n = 0 Then
    Call Размерность_вектора(list, ном_строки, ном_столбца, n, из_строки_вопр)
End If
ReDim A(n)
For i = 1 To n
        If из_строки_вопр = 1 Then
            A(i) = Worksheets(list).Cells(ном_строки, i + ном_столбца - 1)
        Else
            A(i) = Worksheets(list).Cells(i + ном_строки - 1, ном_столбца)
        End If
Next i
End Sub
Sub Вывод_Вектора(list$, A As Variant, ном_строки%, ном_столбца%, В_строку_вопр%, Optional n%)
Dim i%
If n = 0 Then n = UBound(A, 1)
For i = 1 To n
         If В_строку_вопр = 1 Then
             Worksheets(list).Cells(ном_строки, i + ном_столбца - 1) = A(i)
        Else
              Worksheets(list).Cells(i + ном_строки - 1, ном_столбца) = A(i)
        End If
Next i
End Sub
Sub Размерность_матрицы(list$, Ном_стр%, Ном_столб%, n%, Optional m%) ' чёрт её знает проверить не мешает
Dim i%, j%
Do
    j = j + 1
    m = j
Loop While Worksheets(list).Cells(Ном_стр + 1, Ном_столб + j - 1) <> ""
Do
    i = i + 1
    n = i
Loop While Worksheets(list).Cells(Ном_стр + i - 1, Ном_столб + 1) <> ""
End Sub
Sub Размерность_вектора(list$, Ном_стр%, Ном_столб%, n%, из_строки_вопр%) ' чёрт её знает проверить не мешает
Dim i%
i = 1
If из_строки_вопр = 1 Then
Do
    n = i
    i = i + 1
Loop While Worksheets(list).Cells(Ном_стр, Ном_столб + i - 1) <> ""
Else
Do
    n = i
    i = i + 1
Loop While Worksheets(list).Cells(Ном_стр + i - 1, Ном_столб) <> ""
End If
End Sub
Sub Boards(board() As Variant)
board(1) = xlEdgeTop
board(2) = xlEdgeBottom
board(3) = xlEdgeRight
board(4) = xlEdgeLeft
board(5) = xlInsideVertical
board(6) = xlInsideHorizontal
End Sub
