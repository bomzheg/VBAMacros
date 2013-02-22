Attribute VB_Name = "апроксимация"
Option Base 1
Option Explicit
Sub Апр()
    Dim n%, m%, x!(), y!(), i%, j%, A!(), fx$, msg%, R!, bett!()
    Const list$ = "приближение ф-ции"
    If Not [Table].IsWorkSheetExist(list) Then
        MsgBox ("Нет листа " + list)
        Exit Sub
    End If
    Call [Matrix].Ввод_Вектора(list, x, 4, 5, 1, n)
    Call [Matrix].Ввод_Вектора(list, y, 5, 5, 1, n)
    m = InputBox("Ведите количство коэффициентов а, которые необходимо вычислить", , "5", 5000, 5000)
20  ReDim A(m), S(m, 1), bett(1, n)
    Call апроксимация(x, y, m, A, R)
    fx = CStr(A(1))
    For i = 2 To m
        If A(i) >= 0 Then
            fx = fx + " +" + CStr(A(i)) + "*x^" + CStr(i - 1)
        Else: fx = fx + " " + CStr(A(i)) + "*x^" + CStr(i - 1)
        End If
    Next i
    For i = 1 To n
        bett(1, i) = func(A, x, m, i)
    Next i
    msg = MsgBox(fx + ". Её квадратичный критерий R=" + CStr(R) + "." + vbCrLf + "Нажмите да, что бы вывести вектор-столбец коэффициентов в столбец 2," _
    + " начиная с третьей строки," + vbCrLf + "или нет, чтобы увеличить количество коэффициентов на 1. Нажмите отмена, если Вы не хотите выводить коэффициенты", 3, "Функция:")
    Select Case msg
        Case 6
            For i = m To m + 10
                Worksheets(list).Cells(2 + i, 2) = ""
            Next i
            Call [Matrix].Вывод_Вектора(list, A, 3, 2, 0)
            Call [Matrix].Вывод_Матрицы(list, bett, 6, 5)
            Worksheets(list).Cells(10, 5) = R
        Case 7
            m = m + 1
            GoTo 20
    End Select
End Sub
Function func!(A!(), x!(), m%, j%)
    Dim f!, i%
    f = 0
    For i = 1 To m
        f = f + A(i) * x(j) ^ (i - 1)
    Next i
    func = f
End Function
Sub апроксимация(x!(), yv!(), число_коэф%, A!(), Optional R!)
    Dim i%, j%, Ф!(), ФТ!(), MatrN!(), b!(), yT!(), SLAU!(), n%, y!()
    n = UBound(x, 1)
        ReDim Ф(n, число_коэф), ФТ(число_коэф, n), MatrN(число_коэф, число_коэф), b(число_коэф, 1), yT(n, 1), SLAU(число_коэф, число_коэф + 1), y(1, n)
        For i = 1 To n
            y(1, i) = yv(i)
            For j = 1 To число_коэф
                Ф(i, j) = x(i) ^ (j - 1)
            Next j
        Next i
    Call [Matrix].Транспонирование(y, yT)
    Call [Matrix].Транспонирование(Ф, ФТ)
    Call [Matrix].Матр_произв(ФТ, Ф, MatrN)
    Call [Matrix].Матр_произв(ФТ, yT, b)
    For i = 1 To число_коэф
        For j = 1 To число_коэф + 1
            If j <= число_коэф Then
            SLAU(i, j) = MatrN(i, j)
            Else: SLAU(i, j) = b(i, 1)
            End If
        Next j
    Next i
    Call [Matrix].СЛАУ(SLAU, A, число_коэф)
    R = 0
    For j = 1 To n
        R = R + (y(1, j) - func(A, x, число_коэф, j)) ^ 2
    Next j
End Sub
Sub ЭкспоненциальнаяАпроксимация(x!(), y!(), число_коэф%, A!(), R!)
    Dim n%, i%, lny!()
    n = UBound(x, 1)
    ReDim lny(n)
    For i = 1 To n
        lny(i) = Log(y(i)) ' натуральный логарифм
    Next i
    Call апроксимация(x, lny, число_коэф%, A, R)
End Sub
Sub ExponentzialApromaksim()
    Dim n%, x!(), y!(), i%, j%, A!(), fx!, R!, apr!(), lny!(), xkas!, ykas!, k!, b!
    Const list$ = "приближение ф-ции"
    If Not [Table].IsWorkSheetExist(list) Then
        MsgBox ("Нет листа " + list)
        Exit Sub
    End If
    Const m% = 2
    Call [Matrix].Ввод_Вектора(list, x, 4, 5, 1, n)
    Call [Matrix].Ввод_Вектора(list, y, 5, 5, 1, n)
    ReDim apr(n), A(m)
    Call ЭкспоненциальнаяАпроксимация(x, y, m, A, R)
    For i = 1 To n
        apr(i) = Exp(A(1)) * Exp(A(2) * x(i))
    Next i
    Call [Matrix].Вывод_Вектора(list, apr, 7, 5, 1)
    Call [Matrix].Вывод_Вектора(list, A, 17, 8, 1)
    Worksheets(list).Cells(17, 8) = Exp(A(1))
    Worksheets(list).Cells(17, 9) = A(2)
    Worksheets(list).Cells(17, 12) = R
    xkas = InputBox("Задайте координату точки, к кторой нужно провести касательную", "Выполнение макроса", "0", 5000, 5000)
    ykas = Exp(A(1)) * Exp(A(2) * xkas)
    k = A(2) * Exp(A(1)) * Exp(A(2) * xkas)
    b = ykas - k * xkas
    Worksheets(list).Cells(20, 8) = k
    Worksheets(list).Cells(21, 8) = b
    Worksheets(list).Cells(22, 8) = xkas
End Sub
Sub КасательнаяПроизвольнойФункции()
Dim n%, m%, x!(), y!(), i%, A!(), R!, bett!(), xkas!, ykas!, k!, b!
    Const list$ = "приближение ф-ции"
    If Not [Table].IsWorkSheetExist(list) Then
        MsgBox ("Нет листа " + list)
        Exit Sub
    End If
    Call [Matrix].Ввод_Вектора(list, x, 4, 5, 1, n)
    Call [Matrix].Ввод_Вектора(list, y, 5, 5, 1, n)
    m = 7
    ReDim A(m), bett(1, n)
    Call апроксимация(x, y, m, A, R)
    For i = 1 To n
        bett(1, i) = func(A, x, m, i)
    Next i
    xkas = InputBox("Задайте координату точки, к которой нужно провести касательную", "Выполнение макроса", CStr(0.5 * (x(1) + x(2))), 5000, 5000)
    k = 0
    ykas = A(1) * xkas ^ (1 - 1)
    For i = 2 To m
        ykas = ykas + A(i) * xkas ^ (i - 1)
        k = k + (i - 1) * A(i) * xkas ^ (i - 2)
    Next i
    b = ykas - k * xkas
    Worksheets(list).Cells(20, 8) = k
    Worksheets(list).Cells(21, 8) = b
    Worksheets(list).Cells(22, 8) = xkas
    Worksheets(list).Cells(23, 8) = ykas
    Worksheets(list).Cells(23, 7) = "ykas"
    Call [Matrix].Вывод_Вектора(list, A, 3, 2, 0)
    Call [Matrix].Вывод_Матрицы(list, bett, 6, 5)
End Sub








