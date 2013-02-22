Attribute VB_Name = "Мног_опт"
Option Base 1
Option Explicit
Option Private Module
Dim Str_list$, i%, j%, k%
Sub Mn_opt_shag()
Dim x!(), y!(), Fxy!(), n%, dx!, dy, eps!, min!, i1%, j1%, ax!, bx!, ay!, by!
Str_list = "мног.опт"
If Not [Table].IsWorkSheetExist(Str_list) Then
    MsgBox ("Нет листа " + Str_list)
    Exit Sub
End If

n = Worksheets(Str_list).Cells(1, 2)
ReDim x(n), y(n), Fxy(n, n)
ax = Worksheets(Str_list).Cells(2, 2)
bx = Worksheets(Str_list).Cells(3, 2)
ay = Worksheets(Str_list).Cells(4, 2)
by = Worksheets(Str_list).Cells(5, 2)
eps = Worksheets(Str_list).Cells(6, 2)
k = 0
Do
    k = k + 1
    dx = (bx - ax) / n
    dy = (by - ay) / n
    x(1) = ax
    y(1) = ay
    For i = 2 To n
        x(i) = x(i - 1) + dx
        y(i) = y(i - 1) + dy
    Next i
    For i = 1 To n
        For j = 1 To n
            Fxy(i, j) = f(x(i), y(j))
        Next j
    Next i
    Call Вывод_Вектора(Str_list, x, 6, 6, 1)
    Call Вывод_Вектора(Str_list, y, 6, 6, 0)
    Call Вывод_Матрицы(Str_list, Fxy, 6, 6)
    min = Fxy(1, 1)
    For i = 1 To n
        For j = 1 To n
            If min >= Fxy(i, j) Then
                min = Fxy(i, j)
                i1 = i
                j1 = j
            End If
        Next j
    Next i

    Select Case i1
    Case 1
        ax = x(i1) - dx
        bx = x(i1 + 1)
    Case n
        ax = x(i1 - 1)
        bx = x(i1) + dx
    Case Else
        ax = x(i1 - 1)
        bx = x(i1 + 1)
    End Select

    Select Case j1
    Case 1
        ay = y(j1) - dy
        by = y(j1 + 1)
    Case n
        ay = y(j1 - 1)
        by = y(j1) + dy
    Case Else
        ay = y(j1 - 1)
        by = y(j1 + 1)
    End Select
    If k >= 20 Then Exit Do
Loop While (Abs(bx - ax) > eps Or Abs(by - ay) > eps)
Worksheets(Str_list).Cells(8, 3) = min
Worksheets(Str_list).Cells(9, 2) = x(i1)
Worksheets(Str_list).Cells(9, 3) = y(j1)
'Call диаграмма(Fxy, x, y, Str_list)
End Sub
Private Function f!(x!, y!)
f = x ^ 2 - 3 * y ^ 2 + x * y - 6
End Function
Private Function f2!(msng_x!())
f2 = 8 * msng_x(1) ^ 2 + 4 * msng_x(1) * msng_x(2) + 5 * msng_x(2) ^ 2
End Function
Private Function df2x1!(x!())
df2x1 = 16 * x(1) + 4 * x(2)
End Function
Private Function df2x2!(x!())
df2x2 = 4 * x(1) + 10 * x(2)
End Function
Sub gradient(msng_x!(), msng_g!())
msng_g(1) = df2x1(msng_x)
msng_g(2) = df2x2(msng_x)
End Sub
Sub Многомерн_опт_град()
Dim int_i%, int_j%, msng_x!(), int_k%, msng_g!(), sng_Fx!, msng_V!(), sng_h!, msng_z!(), sng_Fz!, sng_eps!, int_n%, msng_temp1!()
Str_list = "мног.опт"
If Not [Table].IsWorkSheetExist(Str_list) Then
    MsgBox ("Нет листа " + Str_list)
    Exit Sub
End If
Call [Matrix].Ввод_Вектора(Str_list, msng_x, 25, 1, 0, int_n)
ReDim msng_g(int_n), msng_V(int_n), msng_z(int_n), msng_temp1(int_n)
sng_h = Worksheets(Str_list).Cells(24, 2)
sng_eps = Worksheets(Str_list).Cells(23, 2)
Do
    sng_Fx = f2(msng_x)
    Call gradient(msng_x, msng_g)
    For int_i = 1 To int_n
        msng_V(int_i) = msng_g(int_i) / [Matrix].Норма_вектора(msng_g, int_n)
    Next int_i
    Call [Matrix].Произведение_вектора_на_число(msng_V, sng_h, msng_temp1)
    Call [Matrix].Разность_векторов(msng_x, msng_temp1, msng_z, int_n)
    sng_Fz = f2(msng_z)
    Call [Matrix].Вывод_Вектора(Str_list, msng_x, 24 + int_k, 5, 0)
    Call [Matrix].Вывод_Вектора(Str_list, msng_g, 24 + int_k, 7, 0)
    Call [Matrix].Вывод_Вектора(Str_list, msng_V, 24 + int_k, 9, 0)
    Call [Matrix].Вывод_Вектора(Str_list, msng_z, 24 + int_k, 11, 0)
    Worksheets(Str_list).Cells(25 + int_k, 6) = sng_Fx
    Worksheets(Str_list).Cells(25 + int_k, 8) = [Matrix].Норма_вектора(msng_g, int_n)
    Worksheets(Str_list).Cells(25 + int_k, 10) = sng_h
    Worksheets(Str_list).Cells(25 + int_k, 12) = sng_Fz
    int_k = int_k + 2
    If sng_Fz < sng_Fx Then
        For int_i = 1 To int_n
            msng_x(int_i) = msng_z(int_i)
        Next int_i
        Worksheets(Str_list).Cells(25 - 2 + int_k, 13) = "Да"
    Else
        Worksheets(Str_list).Cells(25 - 2 + int_k, 13) = "Нет"
        If Abs(sng_h) <= sng_eps Then Exit Do
        sng_h = sng_h / 3
    End If
Loop While int_k <= 50
Worksheets(Str_list).Range(Cells(25 + int_k, 5), Cells(25 + 51, 13)).Clear
End Sub
Private Sub диаграмма(Fxy!(), x!(), y!(), Str_list$)
Dim n%
n = UBound(Fxy, 1)
    Worksheets(Str_list).Range(Cells(7, 7), Cells(16, 16)).Select
    Charts.Add
    ActiveChart.ChartType = xlSurface
    ActiveChart.SetSourceData Source:=Sheets("мног.опт").Range("F6:P16"), PlotBy _
        :=xlRows
'For i = 1 To n
'    ActiveChart.SeriesCollection(i).XValues = x(i)
'    ActiveChart.SeriesCollection(i).Name = y(i)
'Next i
    ActiveChart.SeriesCollection(1).XValues = "=мног.опт!R6C7:R6C16"
    ActiveChart.SeriesCollection(1).Name = "=мног.опт!R7C6"
    ActiveChart.SeriesCollection(2).XValues = "=мног.опт!R6C7:R6C16"
    ActiveChart.SeriesCollection(2).Name = "=мног.опт!R8C6"
    ActiveChart.SeriesCollection(3).XValues = "=мног.опт!R6C7:R6C16"
    ActiveChart.SeriesCollection(3).Name = "=мног.опт!R9C6"
    ActiveChart.SeriesCollection(4).XValues = "=мног.опт!R6C7:R6C16"
    ActiveChart.SeriesCollection(4).Name = "=мног.опт!R10C6"
    ActiveChart.SeriesCollection(5).XValues = "=мног.опт!R6C7:R6C16"
    ActiveChart.SeriesCollection(5).Name = "=мног.опт!R11C6"
    ActiveChart.SeriesCollection(6).XValues = "=мног.опт!R6C7:R6C16"
    ActiveChart.SeriesCollection(6).Name = "=мног.опт!R12C6"
    ActiveChart.SeriesCollection(7).XValues = "=мног.опт!R6C7:R6C16"
    ActiveChart.SeriesCollection(7).Name = "=мног.опт!R13C6"
    ActiveChart.SeriesCollection(8).XValues = "=мног.опт!R6C7:R6C16"
    ActiveChart.SeriesCollection(8).Name = "=мног.опт!R14C6"
    ActiveChart.SeriesCollection(9).XValues = "=мног.опт!R6C7:R6C16"
    ActiveChart.SeriesCollection(9).Name = "=мног.опт!R15C6"
    ActiveChart.SeriesCollection(10).XValues = "=мног.опт!R6C7:R6C16"
    ActiveChart.SeriesCollection(10).Name = "=мног.опт!R16C6"
    ActiveChart.Location Where:=xlLocationAsObject, Name:="мног.опт"
    With ActiveChart
        .HasTitle = False
        .Axes(xlCategory).HasTitle = False
        .Axes(xlSeries).HasTitle = False
        .Axes(xlValue).HasTitle = False
    End With
End Sub
