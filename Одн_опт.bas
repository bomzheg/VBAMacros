Attribute VB_Name = "���_���"
Option Base 1
Option Explicit
Option Private Module
Dim list$
Sub ����������_���()
Dim i%, fx!(), n%, eps!, k%, x!(), dx!, A!, b!, c!
list = "���. ���."
If Not [Table].IsWorkSheetExist(list) Then
    MsgBox ("��� ����� " + list)
    Exit Sub
End If
n = Worksheets(list).Cells(1, 2)
ReDim x(n), fx(n)
A = Worksheets(list).Cells(2, 2)
b = Worksheets(list).Cells(3, 2)
eps = Worksheets(list).Cells(4, 2)
Do
    x(1) = A
    x(n) = b
    dx = (b - A) / n
    fx(1) = f(x(1))
    fx(n) = f(x(n))
    For i = 2 To n - 1                                  ' � ������ �������� �������� ����� ��� ���, ������� ��������� ����������.( ��� ������ � bash.org.ru)
        x(i) = x(i - 1) + dx
        fx(i) = f(x(i))                                 'webster89: �������� � ��������� �� ��������� ������
    Next i                                              'webster89: ��� � ������ :D
                                                        ''''''''''
    For i = 2 To n                                      'xxx: ����, ��� ���� �������, ������� � ���������� ������ ����� )
        If fx(i) < fx(i - 1) Then                       'xxx: ���������� ������...
            c = x(i)                                    ''''''''''
            A = x(i - 1)                                ' ��� ������� ����������, ��� �������,
            If i = n Then                               '��������� ������� ��������� ������������� ���������� �����
                b = x(i) + dx                           '� ����������� ������ �� ��� �� ������ ������ �������� � ������ �������.
            Else
                b = x(i + 1)
            End If
        Else
            If i = 2 Then
                c = x(1)
                A = x(1) - dx
                b = x(2)
            End If
    Exit For
        End If
    Next i
    Call �����_�������(list, x, 2 + k, 8, 1)
    Call �����_�������(list, fx, 3 + k, 8, 1)
    Worksheets(list).Cells(2 + k, 8) = "x"
    Worksheets(list).Cells(3 + k, 8) = "y"
    k = k + 3
Loop While Abs(b - A) > eps And k < 300
Call formating(k, n, list) ' �������������� - ������ ������� � ��.
Worksheets(list).Cells(6, 3) = c
Worksheets(list).Cells(7, 3) = f(c)
Worksheets(list).Range(Cells(k + 1, 8), Cells(301, 18)).Clear
End Sub
Private Function f!(x!)
f = x ^ 3 - 8 * x ^ 2 + 14 * x + 6
End Function
Sub formating(k%, n%, list$)
Dim i%
For i = 0 To k - 3 Step 3
    Worksheets(list).Range(Cells(2 + i, 8), Cells(3 + i, 8 + n)).Select ' ��������� ������� �����
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone ' ������������ ������� - �����������
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone ' ����
    With Selection.Borders(xlEdgeLeft) ' ������ ���� - ���������� ����� �������
        .LineStyle = xlContinuous ' ����� �����
        .Weight = xlThin    ' ��
        .ColorIndex = xlAutomatic ' ������ - ����
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideHorizontal)
         .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone ' �����.������ ���
    
    Worksheets(list).Range(Cells(2 + i, 8), Cells(3 + i, 8)).Select
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlDouble
        .Weight = xlThick
        .ColorIndex = xlAutomatic
    End With
Next i

Worksheets(list).Range(Cells(2 + i, 8), Cells(301, 8 + n)).Select ' �������� ������ � ����������� �����
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
Worksheets(list).Range(Cells(6, 3), Cells(7, 3)).Select
'    With Selection.Interior ' ������� ���� � ����� ����
'        .ColorIndex = 15
'        .Pattern = xlSolid
'    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
        With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
        With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideHorizontal)
         .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End Sub
