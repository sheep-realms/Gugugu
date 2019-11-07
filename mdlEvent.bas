Attribute VB_Name = "mdlEvent"
Option Explicit

Public Const EventLimit As Integer = 103

Public Type GuguguEvent
    Name As String
    Type As String
    Text As String
    Locked As Boolean
    doYes_hp As Integer
    doYes_mp As Integer
    doYes_mn As Long
    doYes_pt As Integer
    doYes_ep As Integer
    doYes_tm As Integer
    doNo_hp As Integer
    doNo_mp As Integer
    doNo_mn As Long
    doNo_pt As Integer
    doNo_ep As Integer
    doNo_tm As Integer
End Type

Public EventList(200) As GuguguEvent

Public Function EventLoad()
    EventList(1).Name = "开始咕咕咕"
    EventList(1).Text = "今天，你成为了一位内容创作者，你将要用一些零花钱开始你的内容创作。不过，你有一个糟糕的毛病，就是喜欢放鸽子。所以，在接下来的创作生涯中，你将要与咕咕咕为伴。你做好准备了吗？" & vbCrLf & "(备注：你，莫得选择！)"
    EventList(1).doYes_mn = 1500
    EventList(1).doNo_mn = 1500
    
    EventList(2).Name = "接手第一份订单"
    EventList(2).Text = "你收到了第一份订单，这是一份很简单的工作！"
    EventList(2).doYes_mn = 20
    EventList(2).doYes_mp = -2
    EventList(2).doYes_pt = 1
    EventList(2).doYes_ep = 10
    EventList(2).doNo_mp = 1
    
    EventList(3).Name = "学习新的创作技巧"
    EventList(3).Text = "你发现你还是太菜了，需要深入学习一下。"
    EventList(3).doYes_mp = -20
    EventList(3).doYes_ep = 30
    EventList(3).doNo_mp = 1
    EventList(3).Locked = True
    
    EventList(4).Name = "熬夜学习创作技巧"
    EventList(4).Text = "你开启了学习模式，学习使你快乐，你逐渐沉迷于其中，突然发现已是深夜。"
    EventList(4).doYes_hp = -5
    EventList(4).doYes_mp = -40
    EventList(4).doYes_ep = 50
    EventList(4).doNo_mp = 25
    
    EventList(5).Name = "迎接全新的开始吧！"
    EventList(5).Text = "你渐渐熟悉了创作者的工作，你看到了一些希望，这使你充满了信心，感觉人生已经抵达了高潮！不过话说回来，自己终究是个鸽子呢......"
    EventList(5).doYes_mp = 50
    EventList(5).doNo_mp = 50
    
    EventList(10).Name = "是时候该睡觉了"
    EventList(10).Text = "睡觉可以使你迅速恢复体力，但不睡觉将会导致你的健康状况堪忧。"
    EventList(10).doYes_hp = 2
    EventList(10).doYes_mp = 50
    EventList(10).doYes_tm = 9
    EventList(10).doNo_hp = -5
    
    EventList(20).Name = "一只鸽子失去了生命"
    EventList(20).Text = "一只鸽子由于某种来自东方的神秘力量，当场去世了。" & vbCrLf & "享年 " & YY & " 年 " & MM & " 月 " & DD & " 日 " & HH & " 时"
    
    EventList(21).Name = "一只鸽子失去了生命"
    EventList(21).Text = "一只鸽子认为自己可以升仙，连续" & NoSleepDays & "天没有睡觉。在这几天的修仙过程中，他体验到了无比的快乐。只是，他终究没有完成他的远大目标，倒在了修仙路上。" & vbCrLf & "享年 " & YY & " 年 " & MM & " 月 " & DD & " 日 " & HH & " 时"
    
    EventList(22).Name = "一只鸽子失去了生命"
    EventList(22).Text = "一只鸽子想尝试修仙，但是他的身体实在是太弱了，没能顶住修仙所带来的健康损失。" & vbCrLf & "享年 " & YY & " 年 " & MM & " 月 " & DD & " 日 " & HH & " 时"

    EventList(23).Name = "一只鸽子自爆了"
    EventList(23).Text = "《鸽子迷惑行为大赏》一书中记载道：一只鸽子不知道为什么购买了自爆道具并使用了，或许他已经看破了红尘，觉得这个世界没有意义，他的所作所为都没有意义。但是更有可能的一种情况就是，他把这当成了什么RPG模式的GalGame，被BOSS击败即可观看CG——当然他在这里不可能看到任何CG，因为这只是个文字游戏，甚至连图像都没有，最多给你整个字符画出来。这种沙雕行为无疑可以被列入鸽子迷惑行为大赏，这连傻瓜都知道这不会有什么好结果。什么？借一部说话？这种东西我有两部，一小一大，你要哪个？哪个都不给你！" & vbCrLf & "享年 " & YY & " 年 " & MM & " 月 " & DD & " 日 " & HH & " 时"
    
    Dim i As Integer
    For i = 100 To EventLimit
        EventList(i).Locked = True
    Next i
    
    EventList(100).Name = "接到了一份小订单"
    EventList(100).Text = "只是一些简单的要求，很快就能完成，不会花费太大功夫的！"
    EventList(100).doYes_mp = -2
    EventList(100).doYes_mn = 20
    EventList(100).doYes_pt = 1
    EventList(100).doYes_ep = 2
    EventList(100).doYes_tm = 1
    EventList(100).doNo_mp = 1
    EventList(100).doNo_tm = 1
    
    EventList(101).Name = "学习一些小技巧"
    EventList(101).Text = "心血来潮想学点创作小技巧。"
    EventList(101).doYes_mp = -5
    EventList(101).doYes_ep = 10
    EventList(101).doYes_tm = 1
    EventList(101).doNo_mp = 1
    EventList(101).doNo_tm = 1
    
    EventList(102).Name = "碰上了一份不小的订单"
    EventList(102).Text = "一个组织正在制作一个名为《%worksname%》的作品，如果你也去掺和一把，或许能捞到不小的收获！"
    EventList(102).doYes_mp = -20
    EventList(102).doYes_mn = 1000
    EventList(102).doYes_pt = 50
    EventList(102).doYes_ep = 200
    EventList(102).doYes_tm = 9
    EventList(102).doNo_mp = 1
    EventList(102).doNo_tm = 1
    
    EventList(103).Name = "要不要抄袭他人的作品？"
    EventList(103).Text = "你看到了一位大神的作品，突然动了邪念，想要“借鉴一下”——"
    EventList(103).doYes_mp = -10
    EventList(103).doYes_pt = -25
    EventList(103).doYes_ep = 2
    EventList(103).doYes_tm = 3
    EventList(103).doNo_mp = 1
    EventList(103).doNo_tm = 1
    
    
    EventList(0).Name = ""
    EventList(0).Text = ""
    EventList(0).Locked = False
End Function

Public Function EventLock(sltName As String)
    If sts_hp = 0 Then
        GameOver = True
        EventLoad
        If NoSleepDays > 1 Then EventSet 21: Exit Function
        If doEvent = 10 Then EventSet 22: Exit Function
        EventSet 20: Exit Function
    End If

    Select Case doEvent
    Case 1
        EventSet 2
        Exit Function
    Case 2
        EventSet 3
        Exit Function
    Case 3
        If sltName = "y" Then
            EventSet 4
            EventList(4).Locked = True
        Else
            EventSet 5
        End If
    Case 4
        EventSet 5
    Case 5
        DateAdd 30
        EventRnd
    Case 10
        If sltName = "y" Then
            doNightSleep = False
            NoSleepDays = 0
        Else
            doNightSleep = True
            NoSleepDays = NoSleepDays + 1
        End If
        EventRnd
    Case 21
        
    Case 103
        stc_CopyLimit = stc_CopyLimit + 1
        EventRnd
    Case Else
        EventRnd
    End Select
End Function

Public Function EventRnd()
    Dim r As Integer, i As Integer
    
    If (((HH >= 21) And (HH <= 23)) Or ((HH >= 0) And (HH <= 5))) And doNightSleep = False Then EventSet 10: Exit Function
    If ((HH >= 6) And (HH <= 20)) And doNightSleep = True Then doNightSleep = False
    
EventRndTop:

    Randomize
    r = Int(Rnd * (EventLimit - 100 + 1)) + 100
    If EventList(r).Locked = False Then
        i = i + 1
        If i >= 10 Then r = 100: EventSet r: Exit Function
        GoTo EventRndTop
    End If
    EventSet r
End Function




