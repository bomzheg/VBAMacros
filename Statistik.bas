Attribute VB_Name = "Statistik"
Option Base 1
Option Explicit
Sub Predstawlenie()
Dim x!(), dx!(), rez$(), pm$, i%, j%, n%, k%, l%, list$, dxs$(), xs$()
list = "статобработка"
Call [Matrix].Ввод_Вектора(list, x, 15, 9, 0, n)
Call [Matrix].Ввод_Вектора(list, dx, 15, 10, 0)
pm = Worksheets(list).Cells(1, 1)
ReDim rez(n), dxs(n), xs(n)
For i = 1 To n
    j = 0
    Do While (dx(i) > 10)
        dx(i) = dx(i) / 10
        j = j + 1
    Loop
    Do While (dx(i) < 1)
            dx(i) = dx(i) * 10
            j = j - 1
    Loop
    dx(i) = CInt(dx(i))
    dxs(i) = dx(i)
    If j <> 0 Then
        If Abs(j) < 3 Then
            dx(i) = dx(i) * 10 ^ j
            dxs(i) = dx(i)
        Else
            dxs(i) = dxs(i) & "E" & CStr(j)
        End If
    End If
    x(i) = x(i) * (10 ^ (-j))
    x(i) = CInt(x(i))
    k = 0
    Do While (x(i) Mod 10 = 0)
        k = k + 1
        x(i) = x(i) \ 10
    Loop
    x(i) = x(i) * 10 ^ k
    x(i) = x(i) * 10 ^ j
    xs(i) = x(i)
    If k > 0 Then
        xs(i) = xs(i) & "."
    End If
    For k = k To 1 Step -1
        xs(i) = xs(i) & "0"
    Next k
Next i
For i = 1 To n
    rez(i) = xs(i) & pm & dxs(i)
Next i
Call [Matrix].Вывод_Вектора(list, rez, 15, 11, 0)
End Sub
Sub Статистика()
10 Dim p!, n%, x!(), dx!, Sx!, Sx2!, i%, j%, dMax!, S!, Urasch!, k%, Ut!, eps!, board(6) As Variant
Const list = "статобработка" ' имя листа задаём констанотой
If Not [Table].IsWorkSheetExist(list) Then
    MsgBox ("Нет листа " + list)
    Exit Sub
End If
p = Worksheets(list).Cells(1, 6) ' задаём вероятность
Call [Matrix].Ввод_Вектора(list, x, 2, 4, 1, n)  ' наш массив чисел который будем обрабатывать
Call [Matrix].Boards(board) 'в этот массив записываю названия границ+)
' АХТУНГ сложный элемент
Worksheets(list).Range(Cells(7, 11), Cells(7 + 10, 11 + n + 10)).Select ' выделяем фрагмент в который предыдущие
'запуски программы записали ответы
Selection.Clear ' ... и удаляем оттуда всё
' стираем выделенные границы
For i = 1 To 6
Selection.Borders(board(i)).LineStyle = xlNone ' так оно короче чем расписывать, каждую границу
Next i
' дада, но без готу совсем криво выглядит. один готу можно и пережить
50
   dx = 0
    For i = 1 To n
        dx = dx + x(i) ' суммируем элементы нашего вектора
    Next i
    dx = dx / n ' получили среднее значение
    Sx2 = 0
    For i = 1 To n
        Sx2 = Sx2 + (x(i) - dx) ^ 2
    Next i
    Sx2 = Sx2 / (n - 1) ' получили дисперсию
    Sx = Sqr(Sx2)
    k = 1
    dMax = Abs(x(k) - dx) ' начинаем искать наибольшее отклонение от среднего c к-атого элемента (первого)
    For i = 2 To n
        S = Abs(x(i) - dx) ' записываем во временную переменную отклонение и-того элемента
        If S > dMax Then ' если отклонение и-того компонента больше чем к-ого
            dMax = S    ' записываем новое отклонение ...
            k = i       ' ... и новый номер отклоняющегося элемента
        End If
    Next i
    Urasch = dMax / (Sx * Sqr(((n - 1) / n))) ' стандартная формула расчёта у-критерия
    Ut = U(p, n - 2)                            ' у-табличное берём из таблицы ' спасибо кеп+)
    If Urasch > Ut Then                         ' если расчётное больше табличного - к-ое измерение содержит грубую
                                                ' или систематичесую ошибку с вероятностью (1-р)
        Call [Matrix].Вывод_Вектора(list, x, 7 + j, 12, 1, n)        ' для удобочитаемости - выводим вектор
                                                                   ' который был на обрабатываемой итерации
        ' АХТУНГ сложный код
        Worksheets(list).Range(Cells(7 + j, 12), Cells(7 + j, 11 + n)).Select ' выделяем только что выведенный вектор
        For i = 1 To 5                                  ' и делаем рамку
            With Selection.Borders(board(i))
                .LineStyle = xlContinuous               'стиль
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
        Next i
        Worksheets(list).Cells(7 + j, 11 + k).Select ' выделение красным отброшенного значения
        With Selection.Interior
            .ColorIndex = 3
            .Pattern = xlSolid
        End With
        x(k) = x(n)     ' в к-ое место записываем последний элемент
        n = n - 1       ' уменьшаем число элементов ( фактически обрубаем возможность обратиться к последнему элементу)
        j = j + 1       ' счётчик используется для вывода по строкам
        GoTo 50
    End If
eps = t(p, n - 1) * Sqr(Sx2 / n)
Call [Matrix].Вывод_Вектора(list, x, 7 + j, 12, 1, n)
Worksheets(list).Range(Cells(7 + j, 12), Cells(7 + j, 11 + n)).Select


For i = 1 To 5
    With Selection.Borders(board(i))
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
Next i
Worksheets(list).Cells(7, 4) = t(p, n - 1)
Worksheets(list).Cells(7, 3) = "Tpf"
Worksheets(list).Cells(3, 4) = dx
Worksheets(list).Cells(3, 3) = "dx"
Worksheets(list).Cells(4, 4) = Sx2
Worksheets(list).Cells(4, 3) = "Sx2"
Worksheets(list).Cells(5, 4) = eps
Worksheets(list).Cells(5, 3) = "eps"
Worksheets(list).Cells(8, 4) = n
Worksheets(list).Cells(8, 3) = "n"
Worksheets(list).Range(Cells(7 + j + 1, 11), Cells(7 + j + 10, 11 + n + 10)).Clear
Worksheets(list).Range("A1").Select
End Sub
Private Function U!(p!, f%)
Dim i%, j%, k%, tempi%, tempj%, modp1!, modf1%, modf2%, modp2!, tp1!, tp2!, tf1%, tf2%
Const strList$ = "критерий стьюдента", n% = 20, m% = 3, kn% = 4, km% = 18
tempi = 1: tempj = 1
If Not [Table].IsWorkSheetExist(strList) Then
    MsgBox ("Нет листа " + strList)
    Exit Function
End If


tf1 = Worksheets(strList).Cells(kn + 1, km)
modf1 = Abs(tf1 - f)
For i = 2 To n
    tf2 = Worksheets(strList).Cells(kn + i, km)
    modf2 = Abs(tf2 - f)
    If modf2 < modf1 Then tempi = i: tf1 = tf2: modf1 = modf2
    If modf1 < 0.001 Then Exit For
Next i

tp1 = Worksheets(strList).Cells(kn, km + 1)
modp1 = Abs(tp1 - p)
For j = 2 To m
    tp2 = Worksheets(strList).Cells(kn, km + j)
    modp2 = Abs(tp2 - p)
    If modp2 < modp1 Then tempj = j: tp1 = tp2: modp1 = modp2
    If modp1 < 0.001 Then Exit For
Next j
U = Worksheets(strList).Cells(tempi + kn, tempj + km)
End Function
Private Function t!(p!, f%)
Dim i%, j%, k%, tempi%, tempj%, modp1!, modf1%, modf2%, modp2!, tp1!, tp2!, tf1%, tf2%
Const strList$ = "критерий стьюдента", n% = 54, m% = 8, kn% = 4, km% = 5
tempi = 1: tempj = 1
If Not [Table].IsWorkSheetExist(strList) Then
    MsgBox ("Нет листа " + strList)
    Exit Function
End If
tf1 = Worksheets(strList).Cells(kn + 1, km)
modf1 = Abs(tf1 - f)
For i = 2 To n
    tf2 = Worksheets(strList).Cells(kn + i, km)
    modf2 = Abs(tf2 - f)
    If modf2 < modf1 Then tempi = i: tf1 = tf2: modf1 = modf2
    If modf1 < 0.001 Then Exit For
Next i

tp1 = Worksheets(strList).Cells(kn, km + 1)
modp1 = Abs(tp1 - p)
For j = 2 To m
    tp2 = Worksheets(strList).Cells(kn, km + j)
    modp2 = Abs(tp2 - p)
    If modp2 < modp1 Then tempj = j: tp1 = tp2: modp1 = modp2
    If modp1 < 0.001 Then Exit For
Next j
t = Worksheets(strList).Cells(tempi + kn, tempj + km)
End Function








