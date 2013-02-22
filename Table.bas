Attribute VB_Name = "Table"
Option Explicit
Option Base 1
Function IsWorkSheetExist(WsName$) As Boolean
Dim c As Object, temp$
 On Error GoTo errÍandle:
 temp = Worksheets(WsName).Cells(1, 1)
   IsWorkSheetExist = True
 Exit Function
errÍandle:
   IsWorkSheetExist = False
 End Function
Function IsWorkSheetWriteble(WsName$, Optional iErr As Long, Optional jErr As Long) As Boolean
 Dim c As Object, temp$, i As Long, j As Long
 Const n As Long = 65536
 Const m% = 256
 On Error GoTo errÍandle:
 For i = 1 To n
    For j = 1 To m
        temp = Worksheets(WsName).Cells(i, j)
        Worksheets(WsName).Cells(i, j) = "0"
        Worksheets(WsName).Cells(i, j) = temp
    Next j
 Next i
   IsWorkSheetWriteble = True
 Exit Function
errÍandle:
   iErr = i
   jErr = j
   IsWorkSheetWriteble = False
 End Function
