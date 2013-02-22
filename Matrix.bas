Attribute VB_Name = "Matrix"
Option Base 1
Option Explicit
Option Private Module
Sub ������������_�������_��_�����(������!(), �����!, ���������!(), Optional n%)
Dim i%
If n = 0 Then
    n = UBound(������, 1)
End If
For i = 1 To n
    ���������(i) = ������(i) * �����
Next
End Sub
Function �����_�������!(A!(), Optional �����_�����%)
Dim S!, i%
S = 0
If �����_����� = 0 Then
�����_����� = UBound(A, 1)
End If
For i = 1 To �����_�����
    S = S + A(i) ^ 2
Next i
�����_������� = Sqr(S)
End Function
Sub ����������������(A!(), At!(), Optional �����_�����%, Optional �����_��������%)
Dim i%, j%
If �����_����� = 0 And �����_�������� = 0 Then
�����_����� = UBound(A, 1)
�����_�������� = UBound(A, 2)
ReDim At(�����_��������, �����_�����)
End If
For i = 1 To �����_�����
    For j = 1 To �����_��������
        At(j, i) = A(i, j)
    Next j
Next i
End Sub
Function �����_�������!(A!(), Optional �����_�����%, Optional �����_��������%)
Dim i%, j%, S!
If �����_����� = 0 And �����_�������� = 0 Then
�����_����� = UBound(A, 1)
�����_�������� = UBound(A, 2)
End If
S = 0
For i = 1 To �����_�����
    For j = 1 To �����_��������
      S = S + A(i, j) ^ 2
    Next j
Next i
�����_������� = Sqr(S)
End Function
Sub �������_�����(�����_�����%, A!(), b!(), R!())
Dim i%
For i = 1 To �����_�����
R(i) = A(i) + b(i)
Next i
End Sub
Sub �����_������(A!(), b!(), c!(), Optional �����_�����%, Optional �����_��������%)
Dim i%, j%
If �����_����� = 0 And �����_�������� = 0 Then
�����_����� = UBound(A, 1)
�����_�������� = UBound(A, 2)
End If
ReDim c(�����_�����, �����_��������)
For i = 1 To �����_�����
    For j = 1 To �����_��������
      c(i, j) = A(i, j) + b(i, j)
    Next j
Next i
End Sub
Sub ��������_������(A!(), b!(), c!(), Optional �����_�����%, Optional �����_��������%)
If �����_����� = 0 And �����_�������� = 0 Then
�����_����� = UBound(A, 1)
�����_�������� = UBound(A, 2)
End If
Dim i%, j%
For i = 1 To �����_�����
    For j = 1 To �����_��������
      c(i, j) = A(i, j) - b(i, j)
    Next j
Next i
End Sub
Sub ��������_��������(A!(), b!(), c!(), Optional �����_���������%)
If �����_��������� = 0 Then
�����_��������� = UBound(A, 1)
End If
Dim i%
For i = 1 To �����_���������
      c(i) = A(i) - b(i)
Next i
End Sub
Sub ����_������(A!(), b!(), c!(), Optional �����_�����A%, Optional �����_��������A%, Optional �����_��������B%)
Dim i%, j%, S!, l%
If �����_�����A = 0 And �����_��������A = 0 And �����_��������B = 0 Then
�����_�����A = UBound(A, 1)
�����_��������A = UBound(A, 2)
�����_��������B = UBound(b, 2)
End If
ReDim c(�����_�����A, �����_��������B)
For i = 1 To �����_�����A
    For j = 1 To �����_��������B
        S = 0
        For l = 1 To �����_��������A
            S = S + A(i, l) * b(l, j)
        Next l
        c(i, j) = S
   Next j
Next i
End Sub
Sub ���������(A!(), d!())
Dim i%, j%, k%, f!(), S!, �����_�����%
�����_����� = UBound(A, 1)
ReDim f(�����_�����, 2 * �����_�����)
For i = 1 To �����_�����
    For j = 1 To 2 * �����_�����
        If j <= �����_����� Then
        f(i, j) = A(i, j)
        Else
        If i = j - �����_����� Then
        f(i, j) = 1
        Else
        f(i, j) = 0
        End If
        End If
    Next j
Next i
For k = 1 To �����_�����
  S = f(k, k)
  For j = k To 2 * �����_�����
    f(k, j) = f(k, j) / S
  Next j
  For i = 1 To �����_�����
      If i <> k Then
         S = f(i, k)
         For j = k To 2 * �����_�����
            f(i, j) = f(i, j) - f(k, j) * S
         Next j
      End If
  Next i
Next k
For i = 1 To �����_�����
    For j = 1 To �����_�����
     d(i, j) = f(i, j + �����_�����)
    Next j
Next i
End Sub
Sub swap(c!, d!)
Dim buf!
buf = c
c = d
d = buf
End Sub
Function ������������(A!())
Dim i%, j%, n%, opr!
n = UBound(A, 1)
opr = 0
For j = 1 To n
opr = opr + A(1, j) * (-1) ^ (1 + j) * det(A(), j)
������������ = opr
Next j
End Function
Function det(A!(), k%)
 Dim b!(), i%, j%, Nstr%      '��������� �� ������ "����, ������ �� ����������?" � "����, ������ �� ��������?" - ������
 Nstr = UBound(A, 1)            ' � ���� ������ �� ��������? - ���� ������ - ����� ��������
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
Sub SLAU_2_matrix(��������������������!(), ���������������!(), x!())
Dim n%, i%, j%, c!()
n = UBound(��������������������, 1)
ReDim c(n, n + 1)
For i = 1 To n
    For j = 1 To n
        c(i, j) = ��������������������(i, j)
    Next j
    c(i, j) = ���������������(i)
Next i
Call ����(c, x)
End Sub
Sub ����(c!(), x!(), Optional n%)
Dim i%, j%, k%, bet!, S!
n = UBound(c, 1)
ReDim x(n)
For i = 1 To n
    For j = 1 To n + 1
     If i = j And c(i, j) = 0 Then
        MsgBox (" ������, ������������ ������� - �������. ���������� ������� ����� ��������")
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
'        MsgBox ("������������ �������, ��� ������ ������")
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
If �����_�������(c) >= 1 Then
      i = MsgBox("������� ���������� �� ����������� =(", vbOKOnly Or vbCritical, "������� ����")
      Exit Sub
End If
Call ����_������(c, d, temp)
Call �����_������(temp, d, xk1)
j = 0
Do
    For i = 1 To n
        xk(i, 1) = xk1(i, 1)
    Next i
    Call ����_������(c, xk, temp)
    Call �����_������(temp, d, xk1)
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
norm = �����_�������(vector)
det = ������������(Matrix)
If Abs(det) < 0.00001 Then
    SLAU_analiz = SLAU_analiz + 1 ' - ������������ �������
End If
If norm > 0.0001 Then
    SLAU_analiz = SLAU_analiz + 10 '- ���������� �������
End If
End Function
Sub ����_�������(list$, A!(), ���_������%, ���_�������%, Optional n%, Optional m%)  '
Dim i%, j%
If n = 0 And m = 0 Then
    Call �����������_�������(list, ���_������, ���_�������, n, m)
End If
ReDim A(n, m)  ' � ����� �� ��� ������ �� ����� end if?
For i = 1 To n
    For j = 1 To m
        A(i, j) = Worksheets(list).Cells(i + ���_������ - 1, j + ���_������� - 1)
    Next j
Next i
End Sub
Sub �����_�������(list$, A$(), ���_������%, ���_�������%)
Dim i%, j%, n%, m%
n = UBound(A, 1)
m = UBound(A, 2)
For i = 1 To n
    For j = 1 To m
       Worksheets(list).Cells(i + ���_������ - 1, j + ���_������� - 1) = A(i, j)
    Next j
Next i
End Sub
Sub ����_�������(list$, A!(), ���_������%, ���_�������%, ��_������_����%, Optional n%)
Dim i%
If n = 0 Then
    Call �����������_�������(list, ���_������, ���_�������, n, ��_������_����)
End If
ReDim A(n)
For i = 1 To n
        If ��_������_���� = 1 Then
            A(i) = Worksheets(list).Cells(���_������, i + ���_������� - 1)
        Else
            A(i) = Worksheets(list).Cells(i + ���_������ - 1, ���_�������)
        End If
Next i
End Sub
Sub �����_�������(list$, A As Variant, ���_������%, ���_�������%, �_������_����%, Optional n%)
Dim i%
If n = 0 Then n = UBound(A, 1)
For i = 1 To n
         If �_������_���� = 1 Then
             Worksheets(list).Cells(���_������, i + ���_������� - 1) = A(i)
        Else
              Worksheets(list).Cells(i + ���_������ - 1, ���_�������) = A(i)
        End If
Next i
End Sub
Sub �����������_�������(list$, ���_���%, ���_�����%, n%, Optional m%) ' ���� � ����� ��������� �� ������
Dim i%, j%
Do
    j = j + 1
    m = j
Loop While Worksheets(list).Cells(���_��� + 1, ���_����� + j - 1) <> ""
Do
    i = i + 1
    n = i
Loop While Worksheets(list).Cells(���_��� + i - 1, ���_����� + 1) <> ""
End Sub
Sub �����������_�������(list$, ���_���%, ���_�����%, n%, ��_������_����%) ' ���� � ����� ��������� �� ������
Dim i%
i = 1
If ��_������_���� = 1 Then
Do
    n = i
    i = i + 1
Loop While Worksheets(list).Cells(���_���, ���_����� + i - 1) <> ""
Else
Do
    n = i
    i = i + 1
Loop While Worksheets(list).Cells(���_��� + i - 1, ���_�����) <> ""
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
