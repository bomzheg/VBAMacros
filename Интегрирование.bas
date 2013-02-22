Attribute VB_Name = "Интегрирование"
Option Base 1
Option Explicit
Option Private Module
Sub integr()
Attribute integr.VB_Description = "численное интегрирование несколькими методами"
Attribute integr.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i%, x!(), fx!(), dSw!(), dSn!(), n%, list$, h!, Sw!, Sn!, Ssr!, dSssr!(), Sssr!, dStr!(), Str!, A!(), R!
list = "integr"
If Not [Table].IsWorkSheetExist(list) Then
    MsgBox ("Нет листа " + list)
    Exit Sub
End If
Call [Matrix].Ввод_Вектора(list, x, 2, 4, 1, n)
Call [Matrix].Ввод_Вектора(list, fx, 3, 4, 1, n)
n = n - 1
h = (x(n + 1) - x(1)) / n
ReDim dSw(n), dSn(n), dSssr(n), dStr(n)
Sw = 0
Sn = 0
Sssr = 0
Str = 0
Do
    Call [апроксимация].апроксимация(x, fx, 2 + i, A, R)
    i = i + 1
Loop While R > 0.1 And i < 3
For i = 1 To n
dSw(i) = fx(i) * h
Sw = Sw + dSw(i)
'''
dSn(i) = fx(i + 1) * h
Sn = Sn + dSn(i)
'''
dSssr(i) = func(A, x(i) + 0.5 * h) * h
Sssr = Sssr + dSssr(i)
''
dStr(i) = func(A, ((x(i) + x(i + 1)) / 2) * h)
Str = Str + dStr(i)
''
Next i
Call [Matrix].Вывод_Вектора(list, dSw, 4, 4, 1)
Call [Matrix].Вывод_Вектора(list, dSn, 5, 5, 1)
Call [Matrix].Вывод_Вектора(list, dSssr, 6, 4, 1)
Call [Matrix].Вывод_Вектора(list, dStr, 7, 4, 1)
Worksheets(list).Cells(11, 4) = Sw
Worksheets(list).Cells(11, 5) = Sn
Ssr = (Sw + Sn) / 2
Worksheets(list).Cells(11, 6) = Ssr
Worksheets(list).Cells(11, 7) = Sssr
Worksheets(list).Cells(11, 8) = Str
End Sub
Function func!(A!(), x!)
Dim n%, i%
n = UBound(A, 1)
For i = 1 To n
    func = func + A(i) * x ^ (i - 1)
Next i
End Function
Sub integr_sim()
Dim i%, x!(), fx!(), dS!(), n%, list$, h!, S!, y!
list = "integr"
If Not [Table].IsWorkSheetExist(list) Then
    MsgBox ("Нет листа " + list)
    Exit Sub
End If
n = Worksheets(list).Cells(1, 2)
ReDim x(n + 1), fx(n + 1), dS(n)
x(1) = Worksheets(list).Cells(2, 1)
x(1 + n) = Worksheets(list).Cells(2, 2)
fx(1) = f(x(1))
fx(n + 1) = f(x(1 + n))
h = (x(n + 1) - x(1)) / n
For i = 2 To n
x(i) = x(i - 1) + h
fx(i) = f(x(i))  ' обращение к функции f
Next i
Call [Matrix].Вывод_Вектора(list, fx, 3, 3, 1)
S = 0
For i = 1 To n
y = f(x(i) + 0.5 * h)
dS(i) = h * (fx(i) + 4 * y + fx(i + 1)) / 6

S = S + dS(i)
Next i
Call [Matrix].Вывод_Вектора(list, dS, 9, 3, 1)
Worksheets(list).Cells(9, 33) = S
End Sub
