Attribute VB_Name = "mdlDate"
Option Explicit

Public YY As Integer, MM As Integer, DD As Integer, HH As Integer

Public Function DateAdd(Value As Integer)
    HH = HH + Value
DateAddTop:
    If HH >= 24 Then HH = HH - 24: DD = DD + 1
    If DD >= 30 Then DD = DD - 29: MM = MM + 1
    If MM >= 12 Then MM = MM - 11: YY = YY + 1
    If HH >= 24 Then GoTo DateAddTop
    DateLoad
End Function

Public Function DateLoad()
    Dim str As String
    If YY <> 0 Then str = str & YY & " 年 "
    If MM <> 0 Then str = str & MM & " 月 "
    If DD <> 0 Then str = str & DD & " 日 "
    str = str & HH & " 时"
    frm.labDate = str
End Function

