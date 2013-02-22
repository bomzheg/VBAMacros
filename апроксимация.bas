Attribute VB_Name = "������������"
Option Base 1
Option Explicit
Sub ���()
    Dim n%, m%, x!(), y!(), i%, j%, A!(), fx$, msg%, R!, bett!()
    Const list$ = "����������� �-���"
    If Not [Table].IsWorkSheetExist(list) Then
        MsgBox ("��� ����� " + list)
        Exit Sub
    End If
    Call [Matrix].����_�������(list, x, 4, 5, 1, n)
    Call [Matrix].����_�������(list, y, 5, 5, 1, n)
    m = InputBox("������ ��������� ������������� �, ������� ���������� ���������", , "5", 5000, 5000)
20  ReDim A(m), S(m, 1), bett(1, n)
    Call ������������(x, y, m, A, R)
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
    msg = MsgBox(fx + ". Ÿ ������������ �������� R=" + CStr(R) + "." + vbCrLf + "������� ��, ��� �� ������� ������-������� ������������� � ������� 2," _
    + " ������� � ������� ������," + vbCrLf + "��� ���, ����� ��������� ���������� ������������� �� 1. ������� ������, ���� �� �� ������ �������� ������������", 3, "�������:")
    Select Case msg
        Case 6
            For i = m To m + 10
                Worksheets(list).Cells(2 + i, 2) = ""
            Next i
            Call [Matrix].�����_�������(list, A, 3, 2, 0)
            Call [Matrix].�����_�������(list, bett, 6, 5)
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
Sub ������������(x!(), yv!(), �����_����%, A!(), Optional R!)
    Dim i%, j%, �!(), ��!(), MatrN!(), b!(), yT!(), SLAU!(), n%, y!()
    n = UBound(x, 1)
        ReDim �(n, �����_����), ��(�����_����, n), MatrN(�����_����, �����_����), b(�����_����, 1), yT(n, 1), SLAU(�����_����, �����_���� + 1), y(1, n)
        For i = 1 To n
            y(1, i) = yv(i)
            For j = 1 To �����_����
                �(i, j) = x(i) ^ (j - 1)
            Next j
        Next i
    Call [Matrix].����������������(y, yT)
    Call [Matrix].����������������(�, ��)
    Call [Matrix].����_������(��, �, MatrN)
    Call [Matrix].����_������(��, yT, b)
    For i = 1 To �����_����
        For j = 1 To �����_���� + 1
            If j <= �����_���� Then
            SLAU(i, j) = MatrN(i, j)
            Else: SLAU(i, j) = b(i, 1)
            End If
        Next j
    Next i
    Call [Matrix].����(SLAU, A, �����_����)
    R = 0
    For j = 1 To n
        R = R + (y(1, j) - func(A, x, �����_����, j)) ^ 2
    Next j
End Sub
Sub ����������������������������(x!(), y!(), �����_����%, A!(), R!)
    Dim n%, i%, lny!()
    n = UBound(x, 1)
    ReDim lny(n)
    For i = 1 To n
        lny(i) = Log(y(i)) ' ����������� ��������
    Next i
    Call ������������(x, lny, �����_����%, A, R)
End Sub
Sub ExponentzialApromaksim()
    Dim n%, x!(), y!(), i%, j%, A!(), fx!, R!, apr!(), lny!(), xkas!, ykas!, k!, b!
    Const list$ = "����������� �-���"
    If Not [Table].IsWorkSheetExist(list) Then
        MsgBox ("��� ����� " + list)
        Exit Sub
    End If
    Const m% = 2
    Call [Matrix].����_�������(list, x, 4, 5, 1, n)
    Call [Matrix].����_�������(list, y, 5, 5, 1, n)
    ReDim apr(n), A(m)
    Call ����������������������������(x, y, m, A, R)
    For i = 1 To n
        apr(i) = Exp(A(1)) * Exp(A(2) * x(i))
    Next i
    Call [Matrix].�����_�������(list, apr, 7, 5, 1)
    Call [Matrix].�����_�������(list, A, 17, 8, 1)
    Worksheets(list).Cells(17, 8) = Exp(A(1))
    Worksheets(list).Cells(17, 9) = A(2)
    Worksheets(list).Cells(17, 12) = R
    xkas = InputBox("������� ���������� �����, � ������ ����� �������� �����������", "���������� �������", "0", 5000, 5000)
    ykas = Exp(A(1)) * Exp(A(2) * xkas)
    k = A(2) * Exp(A(1)) * Exp(A(2) * xkas)
    b = ykas - k * xkas
    Worksheets(list).Cells(20, 8) = k
    Worksheets(list).Cells(21, 8) = b
    Worksheets(list).Cells(22, 8) = xkas
End Sub
Sub ������������������������������()
Dim n%, m%, x!(), y!(), i%, A!(), R!, bett!(), xkas!, ykas!, k!, b!
    Const list$ = "����������� �-���"
    If Not [Table].IsWorkSheetExist(list) Then
        MsgBox ("��� ����� " + list)
        Exit Sub
    End If
    Call [Matrix].����_�������(list, x, 4, 5, 1, n)
    Call [Matrix].����_�������(list, y, 5, 5, 1, n)
    m = 7
    ReDim A(m), bett(1, n)
    Call ������������(x, y, m, A, R)
    For i = 1 To n
        bett(1, i) = func(A, x, m, i)
    Next i
    xkas = InputBox("������� ���������� �����, � ������� ����� �������� �����������", "���������� �������", CStr(0.5 * (x(1) + x(2))), 5000, 5000)
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
    Call [Matrix].�����_�������(list, A, 3, 2, 0)
    Call [Matrix].�����_�������(list, bett, 6, 5)
End Sub








