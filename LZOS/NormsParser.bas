Attribute VB_Name = "NormsParser"
Option Explicit
Option Base 1
Function PriceFromStr#(str$)
Attribute PriceFromStr.VB_ProcData.VB_Invoke_Func = " \n14"
Dim tmpstr$
   tmpstr = Mid(str, 84, 7)
   If Trim(tmpstr) = "" Then
      PriceFromStr = 0
   Else
      PriceFromStr = ConvDblForm(tmpstr)
   End If
End Function
Function NormFromStr#(str$)
Attribute NormFromStr.VB_ProcData.VB_Invoke_Func = " \n14"
Dim tmpstr$
   tmpstr = Mid(str, 71, 7)
   If Trim(tmpstr) = "" Then
      NormFromStr = 0
   Else
      NormFromStr = ConvDblForm(tmpstr)
   End If
End Function
Function ConvDblForm#(str$) ' Заменяет точку на запятую и переводит в формат дабл
Attribute ConvDblForm.VB_ProcData.VB_Invoke_Func = " \n14"
Dim tmpstr$
   tmpstr = str
   Mid(tmpstr, 2, 1) = ","
   ConvDblForm = CDbl(tmpstr)
End Function
Function PerFromStr$(str$)
Attribute PerFromStr.VB_ProcData.VB_Invoke_Func = " \n14"
   PerFromStr = Trim(Mid(str, 12, 21)) 'передел
End Function
Function PerNumFromStr%(str$)
Attribute PerNumFromStr.VB_ProcData.VB_Invoke_Func = " \n14"
   PerNumFromStr = Trim(Mid(str, 8, 3)) '№передела
End Function
Function ShObFromStr$(str$)
Attribute ShObFromStr.VB_ProcData.VB_Invoke_Func = " \n14"
   ShObFromStr = Trim(Mid(str, 35, 10)) 'шифр оборудования
End Function
Function ProfFromStr%(str$)
Attribute ProfFromStr.VB_ProcData.VB_Invoke_Func = " \n14"
   ProfFromStr = Trim(Mid(str, 52, 3)) 'профессия
End Function
Function ShifrFromHeader%(str$, Shifr$())
Attribute ShifrFromHeader.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i%
Const n% = 3
ReDim Shifr(n)
For i = 0 To n - 1
   Shifr(i + 1) = Mid(str, 36 + 5 * i, 5) 'профессия
   If Shifr(i + 1) = "00000" Then
      Exit For
   End If
Next i
ShifrFromHeader = i
End Function
Function DetalFromHeader$(str$)
Attribute DetalFromHeader.VB_ProcData.VB_Invoke_Func = " \n14"
   DetalFromHeader = Mid(str, 3, 11) 'Деталь
End Function
Function DescrFromHeader$(str$)
Attribute DescrFromHeader.VB_ProcData.VB_Invoke_Func = " \n14"
   DescrFromHeader = Trim(Mid(str, 15, 20)) 'SetDescript
End Function
