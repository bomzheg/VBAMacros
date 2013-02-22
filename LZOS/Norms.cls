VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Norms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Private list$, IsNorma%, PeredelName$, NumPeredel%, altNorm%, dontSearch As Boolean
Sub initialize(plist$, pIsNorma As Boolean, pPeredelName$)
list = plist
If pIsNorma Then IsNorma = 0 Else IsNorma = 6
altNorm = AltNormPointer
Me.Peredel = pPeredelName
End Sub
Sub GetAll(plist$, pIsNorma%, pPeredelName$)
plist = list
pIsNorma = IsNorma
pPeredelName = PeredelName
End Sub
Property Get PeredelNum%()
PeredelNum = NumPeredel
End Property
Property Let dntSearch(pSearch As Boolean)
dontSearch = pSearch
End Property
Property Let Peredel(pPeredel$)
PeredelName = pPeredel
NumPeredel = 0
With Worksheets(list)
    For NumPeredel = 0 To 20
        If (LCase(.Cells(4, 5 + NumPeredel)) = LCase(PeredelName)) Then
            Exit For
        End If
    Next NumPeredel
End With
If (NumPeredel > 20) Then NumPeredel = -1
End Property
Property Get Peredel$()
Peredel = PeredelName
End Property
Function GetPrice#(Shifr$, Detal$)
Dim nStr&, tmpPrice#, tmpDetal$
tmpDetal = TrimZero(Detal)
With Worksheets(list)
   If "4" = Mid(Shifr, 1, 1) Or "8" = Mid(Shifr, 1, 1) Then
      Mid(Shifr, 1, 1) = "0"
   End If
   If SearchByShifrAndDetal(Shifr, Detal, nStr) Then
      tmpPrice = .Cells(5 + nStr, 5 + NumPeredel + IsNorma)
      If tmpPrice = 0 Then
         GetPrice = -1#
      Else
         GetPrice = tmpPrice
      End If
   Else
      GetPrice = -1#
   End If
End With
End Function
Function TrimZeroOLD$(Detal$)
Dim tempstr$, lenD%
tempstr = Detal
lenD = Len(Detal)
Do While Mid(tempstr, lenD, 1) = "0"
   tempstr = Mid(tempstr, 1, lenD - 1)
   lenD = Len(tempstr)
Loop
Do While Mid(tempstr, 1, 1) = "0"
   tempstr = Mid(tempstr, 2, lenD - 1)
   lenD = Len(tempstr)
Loop
TrimZero = tempstr
End Function
Function TrimZero$(Detal$)
Dim tempstr$, lenD%, i%
tempstr = Detal
lenD = Len(Detal)
Do While Mid(tempstr, i + 1, 1) = "0"
   i = i + 1
Loop
i = 0
tempstr = Mid(tempstr, i + 1, Len(tempstr) - i)
Do While Mid(tempstr, lenD - i, 1) = "0"
   i = i + 1
Loop
TrimZero = Mid(tempstr, 1, Len(tempstr) - i)
End Function
Sub SetDescript(Shifr$, Detal$, Description$)
Dim nStr%
With Worksheets(list)
    If Not SearchByShifrAndDetal(Shifr$, Detal$, nStr%) Then
        .Cells(5 + nStr, 3) = Shifr
        .Cells(5 + nStr, 4) = Detal
    End If
      .Cells(5 + nStr, 2) = Description
End With
End Sub
Sub SetPrice(Shifr$, Detal$, Price#)
    Call SetCustomPrice(Shifr$, Detal$, Price#, IsNorma)
End Sub
Sub SetAlternativePrice(Shifr$, Detal$, AltPrice#)
    Call SetCustomPrice(Shifr$, Detal$, AltPrice#, altNorm)
End Sub
Sub SetPriceAndNormOld(Shifr$, Detal$, Price#, Norm#, Optional Description$) ' возможно тоже больit не требуется
    Call SetCustomPrice(Shifr, Detal, Norm, IsNorma, Description)
    Call SetCustomPrice(Shifr, Detal, Price, altNorm, Description)
End Sub
Sub SetPriceAndNorm(Shifr$, Detal$, Price#, Norm#, Optional Description$)
Static nStr&, NewDetal As Boolean
If nStr = 0 Then nStr = GetNormsCount
With Worksheets(list)
   If dontSearch Then
      NewDetal = True
   ElseIf Not (SearchByShifrAndDetal(Shifr, Detal, nStr)) Then
         NewDetal = True
   End If
   If NewDetal Then
      .Cells(5 + nStr, 3) = Shifr
      .Cells(5 + nStr, 4) = Detal
      .Cells(5 + nStr, 2) = Description
   End If
   Call SetNewPrice(5 + nStr, 5 + NumPeredel + altNorm, Price)
   Call SetNewPrice(5 + nStr, 5 + NumPeredel + IsNorma, Norm)
End With
End Sub
Private Sub SetCustomPrice(Shifr$, Detal$, Price#, NormPointer%, Optional Description$) ' возможно больше не требуется
Dim nStr&
With Worksheets(list)
    If Not (SearchByShifrAndDetal(Shifr, Detal, nStr)) Then
        .Cells(5 + nStr, 3) = Shifr
        .Cells(5 + nStr, 4) = Detal
        .Cells(5 + nStr, 2) = Description
    End If
    Call SetNewPrice(5 + nStr, 5 + NumPeredel + NormPointer, Price)
End With
End Sub
Private Sub SetNewPrice(Row&, Col&, Price#)
With Worksheets(list)
    If Price = -1# Or Price = 0 Then
        .Cells(Row, Col) = ""
    Else
        .Cells(Row, Col) = Price
    End If
End With
End Sub
Private Function AltNormPointer%()
AltNormPointer = Abs(IsNorma - 6)
End Function
Private Function GetNormsCount&()
Dim CountDetal&
Call [Matrix].Размерность_вектора(list, 5, 3, CountDetal, 0)
GetNormsCount = CountDetal
End Function
Private Function SearchByShifrAndDetal(Shifr$, Detal$, nStr&) As Boolean
Dim CountDetal&
CountDetal = GetNormsCount
With Worksheets(list)
For nStr = 0 To CountDetal - 1
    If (.Cells(5 + nStr, 3) = Shifr) And (TrimZero(.Cells(5 + nStr, 4)) = Detal) Then
        SearchByShifrAndDetal = True
        Exit Function
    End If
Next nStr
End With
SearchByShifrAndDetal = False
End Function
