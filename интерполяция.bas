Attribute VB_Name = "������������"
Option Base 1
Option Explicit
Option Private Module
Sub ������������_����������_��������_�������()
Dim x!(), y!(), n%, i%, j%, list$, c!(), l!(), xras!, yras!
n = Cells(2, 8)
list = InputBox("������� �������� �����:", "���������� �������.", "����������� �-���", 5000, 5000)
If Not [Table].IsWorkSheetExist(list) Then
    MsgBox ("��� ����� " + list)
    Exit Sub
End If
xras = InputBox("������� �������� �," + vbCrLf + "��� �������� ���������� ��������� Y", "������� �������", , 5000, 5000)
ReDim x(1, n), y(1, n), c(1, n), l(1, n)
Call [Matrix].����_�������(list, x, 3, 4, 1, n)
Call [Matrix].����_�������(list, y, 4, 4, 1, n)
For i = 1 To n
    l(1, i) = Ci(y, x, i)
        For j = 1 To i - 1
            l(1, i) = l(1, i) * (xras - x(1, j))
        Next j
        For j = 1 + i To n
            l(1, i) = l(1, i) * (xras - x(1, j))
        Next j
Next i
For i = 1 To n
yras = yras + l(1, i)
Next i
i = MsgBox("�������� ������� � ����� " + CStr(xras) + " �����: " + CStr(yras) + ". ������� �������� � ������ (8,6)?", vbYesNo, "���������")
If i = 6 Then
Worksheets(list).Cells(8, 6) = yras
End If
End Sub
Private Function Ci!(y!(), x!(), i%)
Dim j%, n%
n = UBound(x, 2)
Ci = y(1, i)
For j = 1 To i - 1
    Ci = Ci / (x(1, i) - x(1, j))
Next j
For j = i + 1 To n
    Ci = Ci / (x(1, i) - x(1, j))
Next j
End Function
Sub ������������_����������_�������������()
Dim x!(), y!(), n%, i%, j%, b%, list$, A!(), l!(), k%, yap!() '' ��������� �������� ������ ���������
list = "����������� �-���"
If Not [Table].IsWorkSheetExist(list) Then
    MsgBox ("��� ����� " + list)
    Exit Sub
End If
If list = "" Then
Exit Sub
End If
n = Worksheets(list).Cells(2, 8)
ReDim x(1, n), y(1, n), A(n, 1), l(1, n), yap(1, n)
Call [Matrix].����_�������(list, x, 3, 4, 1, n)
Call [Matrix].����_�������(list, y, 4, 4, 1, n)
''''
A(n, 1) = (n - 1) * Ci2(y, x, n)
For b = 1 To n - 1
  ' a(n - b, 1) = Ci2(y, x)  '�������� ������������ � � ���� "a(n - b, 1) = Ci(y, x, (n - b)) * (-1) ^ b"
  '  a(n - b, 1) = a(n - b, 1) * (-1) ^ b
        l(1, n - b) = 0
        For i = 1 To n
            k = 0
            l(1, n - b) = l(1, n - b) + x(1, i) * Interpol(x, i, b, k) ' ������� ������
        Next i
        A(n - b, 1) = l(1, n - b) * Ci2(y, x, n - b) * (-1) ^ b
Next b
Call [Matrix].�����_�������(list, A, 2, 0)
For i = 1 To n
    yap(1, i) = 0
    For j = 1 To n
        yap(1, i) = yap(1, i) + A(j, 1) * x(1, i) ^ (j - 1)
    Next j
    Worksheets(list).Cells(7, 4 + i) = yap(1, i)
Next i
End Sub
Function Interpol!(x!(), i%, b%, k%)
Dim n%, j%, S!
n = UBound(x, 2)
k = k + 1
S = 0
Select Case k
Case b - 2
    For j = 1 + i To n
        S = S + x(1, j)
    Next j
Case b - 1
    S = 1
Case b
    S = 0
Case Else
    For j = 1 + i To n
        S = S + x(1, j) * Interpol(x, j, b, k)
    Next j
End Select
Interpol = S * (n - k - 1)
End Function
Private Function Ci2!(y!(), x!(), k%)
Dim j%, n%, i%, c!()
n = UBound(x, 2)
ReDim c(n)
Ci2 = 0
Select Case k
Case 1
For i = 2 To n
    c(i) = y(1, i)
    Select Case i
        Case 1
            For j = 2 To n
                c(i) = c(i) / (x(1, i) - x(1, j))
            Next j
        Case n
            For j = 1 To n - 1
                c(i) = c(i) / (x(1, i) - x(1, j))
            Next j
        Case Else
            For j = 1 To i - 1
                c(i) = c(i) / (x(1, i) - x(1, j))
            Next j
            For j = i + 1 To n
                c(i) = c(i) / (x(1, i) - x(1, j))
            Next j
    End Select
    Ci2 = Ci2 + c(i)
Next i
Case 2 To n - 1
For i = 1 To k - 1
    c(i) = y(1, i)
    Select Case i
        Case 1
            For j = 2 To n
                c(i) = c(i) / (x(1, i) - x(1, j))
            Next j
        Case n
            For j = 1 To n - 1
                c(i) = c(i) / (x(1, i) - x(1, j))
            Next j
        Case Else
            For j = 1 To i - 1
                c(i) = c(i) / (x(1, i) - x(1, j))
            Next j
            For j = i + 1 To n
                c(i) = c(i) / (x(1, i) - x(1, j))
            Next j
    End Select
    Ci2 = Ci2 + c(i)
Next i
For i = k + 1 To n
    c(i) = y(1, i)
    Select Case i
        Case 1
            For j = 2 To n
                c(i) = c(i) / (x(1, i) - x(1, j))
            Next j
        Case n
            For j = 1 To n - 1
                c(i) = c(i) / (x(1, i) - x(1, j))
            Next j
        Case Else
            For j = 1 To i - 1
                c(i) = c(i) / (x(1, i) - x(1, j))
            Next j
            For j = i + 1 To n
                c(i) = c(i) / (x(1, i) - x(1, j))
            Next j
    End Select
    Ci2 = Ci2 + c(i)
Next i
Case n
For i = 1 To n - 1
    c(i) = y(1, i)
    Select Case i
        Case 1
            For j = 2 To n
                c(i) = c(i) / (x(1, i) - x(1, j))
            Next j
        Case n
            For j = 1 To n - 1
                c(i) = c(i) / (x(1, i) - x(1, j))
            Next j
        Case Else
            For j = 1 To i - 1
                c(i) = c(i) / (x(1, i) - x(1, j))
            Next j
            For j = i + 1 To n
                c(i) = c(i) / (x(1, i) - x(1, j))
            Next j
    End Select
    Ci2 = Ci2 + c(i)
Next i
End Select
End Function

Sub ������������_4�_��������������()
Dim x!(), y!(), n%, i%, j%, b%, list$, A!(), l!(), k%, yap!() '' ��������� �������� ������ ���������
' ����������� ���� ��������
list = InputBox("������� �������� �����:", "���������� �������.", "����������� �-���", 5000, 5000)
If Not [Table].IsWorkSheetExist(list) Then
    MsgBox ("��� ����� " + list)
    Exit Sub
End If
n = Worksheets(list).Cells(2, 8)
ReDim x(1, n), y(1, n), A(n, 1), l(1, n), yap(1, n)
Call [Matrix].����_�������(list, x, 3, 4, 1, n)
Call [Matrix].����_�������(list, y, 4, 4, 1, n)
''''
A(n, 1) = (n) * Ci2(y, x)
For b = 1 To n - 1
        l(1, n - b) = 0
        For i = 1 To n
            k = 0
            l(1, n - b) = l(1, n - b) + x(1, i) * Interpol(x, i, b, k) ' ������� ������
        Next i
        A(n - b, 1) = l(1, n - b) * Ci2(y, x) * (-1) ^ b
Next b
End Sub










