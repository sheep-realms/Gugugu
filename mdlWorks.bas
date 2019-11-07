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
    
    WorksInfo(0).Name = "我真的想不出什么名字了"
    
    WorksInfo(1).Name = "不正常的小冒险"
    WorksInfo(1).Type = "t1"
    WorksInfo(1).Locked = True
    WorksInfo(1).IndexName(1) = ""
    WorksInfo(1).IndexName(2) = "2:卷土重来"
    WorksInfo(1).IndexName(3) = "3:三回啊三回"
    WorksInfo(1).IndexName(4) = "4:四不四傻"
    WorksInfo(1).IndexName(5) = "5:随手取了个名字出来"
    WorksInfo(1).IndexName(6) = "6:六翻了"
    WorksInfo(1).indexTop = 6
    
    WorksInfo(2).Name = "大魔王之剑"
    WorksInfo(1).Type = "t1"
    WorksInfo(2).Locked = True
    WorksInfo(2).IndexName(1) = ""
    WorksInfo(2).IndexName(2) = "2:英雄归来"
    WorksInfo(2).IndexName(3) = "3:兵临城下"
    WorksInfo(2).IndexName(4) = "4:东山再起"
    WorksInfo(2).IndexName(5) = "5:绝地反击"
    WorksInfo(2).IndexName(6) = "6:终焉之战"
    WorksInfo(2).IndexName(7) = "外传:阿尔法"
    WorksInfo(2).IndexName(8) = "外传:复兴"
    WorksInfo(2).indexTop = 8
    
    WorksInfo(3).Name = "畅玩红月"
    WorksInfo(3).Type = "t1"
    WorksInfo(3).Locked = True
    WorksInfo(3).IndexName(1) = ""
    WorksInfo(3).IndexName(2) = "2"
    WorksInfo(3).IndexName(3) = "3"
    WorksInfo(3).IndexName(4) = "精华版"
    WorksInfo(3).IndexName(5) = "重置版"
    WorksInfo(3).IndexName(6) = "传奇版"
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
