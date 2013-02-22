VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Peredel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Option Base 1
Private PerGlanz%, SelectedPer%
Private Const AllPeredel As Boolean = False
Private Const PerSand% = 1, PerZink% = 2, PerAnod% = 3, PerVoron% = 4, PerChrom% = 5
Private PerCount%, Peredels$()
Private Const mSand$ = "Песок", mZink$ = "Цинк", mAnod$ = "Анодирование", mVoron$ = "Воронение", mChrom = "Хром", mGlanz = "Полировка"
Sub init()
Const AllPer% = 6, SmallPer% = 3
If AllPeredel Then
   PerCount = AllPer
   ReDim Peredels(PerCount)
   Peredels(PerAnod) = mAnod
   Peredels(PerVoron) = mVoron
   Peredels(PerChrom) = mChrom
   PerGlanz = PerCount
Else
   PerCount = SmallPer
   ReDim Peredels(PerCount)
   PerGlanz = PerCount
End If
   Peredels(PerSand) = mSand
   Peredels(PerZink) = mZink
   Peredels(PerGlanz) = mGlanz
End Sub
Property Get Count%()
   Count = PerCount
End Property
Property Get Pers() As Variant
   Pers = Peredels
End Property
Property Get Selected%()
   Selected = SelectedPer
End Property
Property Let Selected(pPerNum%)
   SelectedPer = pPerNum
End Property
Property Get Glanz%()
   Glanz = PerGlanz
End Property
Property Get Sand%()
   Sand = PerSand
End Property
Property Get Zink%()
   Zink = PerZink
End Property
Property Get Anod%()
   Anod = PerAnod
End Property
Property Get Voron%()
   Voron = PerVoron
End Property
Property Get Chrom%()
   Chrom = PerChrom
End Property
Property Get SelectedStr$()
   SelectedStr = Peredels(SelectedPer)
End Property
Property Let SelectedStr(pPerName$)
   For i = 1 To PerCount
      If UCase(pPerName) = UCase(Peredels(i)) Then
         SelectedPer = i
      End If
   Next i
End Property
