Attribute VB_Name = "mdlWorks"
Option Explicit

Public WorksLimit As Integer

Public Type GuguguWorks
    Name As String
    Type As String
    IndexName(30) As String
    indexTop As String
    sts_Index As Integer
    Locked As Boolean
End Type

Public WorksInfo(20) As GuguguWorks

Public Function WorksLoad()
    WorksLimit = 3
    
    WorksInfo(0).Name = "������벻��ʲô������"
    
    WorksInfo(1).Name = "��������Сð��"
    WorksInfo(1).Type = "t1"
    WorksInfo(1).Locked = True
    WorksInfo(1).IndexName(1) = ""
    WorksInfo(1).IndexName(2) = "2:��������"
    WorksInfo(1).IndexName(3) = "3:���ذ�����"
    WorksInfo(1).IndexName(4) = "4:�Ĳ���ɵ"
    WorksInfo(1).IndexName(5) = "5:����ȡ�˸����ֳ���"
    WorksInfo(1).IndexName(6) = "6:������"
    WorksInfo(1).indexTop = 6
    
    WorksInfo(2).Name = "��ħ��֮��"
    WorksInfo(1).Type = "t1"
    WorksInfo(2).Locked = True
    WorksInfo(2).IndexName(1) = ""
    WorksInfo(2).IndexName(2) = "2:Ӣ�۹���"
    WorksInfo(2).IndexName(3) = "3:���ٳ���"
    WorksInfo(2).IndexName(4) = "4:��ɽ����"
    WorksInfo(2).IndexName(5) = "5:���ط���"
    WorksInfo(2).IndexName(6) = "6:����֮ս"
    WorksInfo(2).IndexName(7) = "�⴫:������"
    WorksInfo(2).IndexName(8) = "�⴫:����"
    WorksInfo(2).indexTop = 8
    
    WorksInfo(3).Name = "�������"
    WorksInfo(3).Type = "t1"
    WorksInfo(3).Locked = True
    WorksInfo(3).IndexName(1) = ""
    WorksInfo(3).IndexName(2) = "2"
    WorksInfo(3).IndexName(3) = "3"
    WorksInfo(3).IndexName(4) = "������"
    WorksInfo(3).IndexName(5) = "���ð�"
    WorksInfo(3).IndexName(6) = "�����"
    WorksInfo(3).indexTop = 6
    
End Function

Public Function WorksRnd() As String
    Dim r As Integer, i As Integer
    
WorksRndTop:

    Randomize
    r = Int(Rnd * (WorksLimit - 1 + 1)) + 1
    If WorksInfo(r).Locked = False Then
        i = i + 1
        If i >= 10 Then r = 0: GoTo WorksRndEnd
        GoTo WorksRndTop
    End If
    
WorksRndEnd:
    
    WorksRnd = WorksUse(r)
End Function

Public Function WorksUse(Value As Integer) As String
    If Value <> 0 Then
        WorksInfo(Value).sts_Index = WorksInfo(Value).sts_Index + 1
        If WorksInfo(Value).sts_Index >= WorksInfo(Value).indexTop Then WorksInfo(Value).Locked = False
    End If
    WorksUse = WorksInfo(Value).Name & WorksInfo(Value).IndexName(WorksInfo(Value).sts_Index)
End Function
