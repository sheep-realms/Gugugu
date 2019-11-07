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
    EventList(1).Name = "��ʼ������"
    EventList(1).Text = "���죬���Ϊ��һλ���ݴ����ߣ��㽫Ҫ��һЩ�㻨Ǯ��ʼ������ݴ���������������һ������ë��������ϲ���Ÿ��ӡ����ԣ��ڽ������Ĵ��������У��㽫Ҫ�빾����Ϊ�顣������׼������" & vbCrLf & "(��ע���㣬Ī��ѡ��)"
    EventList(1).doYes_mn = 1500
    EventList(1).doNo_mn = 1500
    
    EventList(2).Name = "���ֵ�һ�ݶ���"
    EventList(2).Text = "���յ��˵�һ�ݶ���������һ�ݺܼ򵥵Ĺ�����"
    EventList(2).doYes_mn = 20
    EventList(2).doYes_mp = -2
    EventList(2).doYes_pt = 1
    EventList(2).doYes_ep = 10
    EventList(2).doNo_mp = 1
    
    EventList(3).Name = "ѧϰ�µĴ�������"
    EventList(3).Text = "�㷢���㻹��̫���ˣ���Ҫ����ѧϰһ�¡�"
    EventList(3).doYes_mp = -20
    EventList(3).doYes_ep = 30
    EventList(3).doNo_mp = 1
    EventList(3).Locked = True
    
    EventList(4).Name = "��ҹѧϰ��������"
    EventList(4).Text = "�㿪����ѧϰģʽ��ѧϰʹ����֣����𽥳��������У�ͻȻ����������ҹ��"
    EventList(4).doYes_hp = -5
    EventList(4).doYes_mp = -40
    EventList(4).doYes_ep = 50
    EventList(4).doNo_mp = 25
    
    EventList(5).Name = "ӭ��ȫ�µĿ�ʼ�ɣ�"
    EventList(5).Text = "�㽥����Ϥ�˴����ߵĹ������㿴����һЩϣ������ʹ����������ģ��о������Ѿ��ִ��˸߳���������˵�������Լ��վ��Ǹ�������......"
    EventList(5).doYes_mp = 50
    EventList(5).doNo_mp = 50
    
    EventList(10).Name = "��ʱ���˯����"
    EventList(10).Text = "˯������ʹ��Ѹ�ٻָ�����������˯�����ᵼ����Ľ���״�����ǡ�"
    EventList(10).doYes_hp = 2
    EventList(10).doYes_mp = 50
    EventList(10).doYes_tm = 9
    EventList(10).doNo_hp = -5
    
    EventList(20).Name = "һֻ����ʧȥ������"
    EventList(20).Text = "һֻ��������ĳ�����Զ�������������������ȥ���ˡ�" & vbCrLf & "���� " & YY & " �� " & MM & " �� " & DD & " �� " & HH & " ʱ"
    
    EventList(21).Name = "һֻ����ʧȥ������"
    EventList(21).Text = "һֻ������Ϊ�Լ��������ɣ�����" & NoSleepDays & "��û��˯�������⼸������ɹ����У������鵽���ޱȵĿ��֡�ֻ�ǣ����վ�û���������Զ��Ŀ�꣬����������·�ϡ�" & vbCrLf & "���� " & YY & " �� " & MM & " �� " & DD & " �� " & HH & " ʱ"
    
    EventList(22).Name = "һֻ����ʧȥ������"
    EventList(22).Text = "һֻ�����볢�����ɣ�������������ʵ����̫���ˣ�û�ܶ�ס�����������Ľ�����ʧ��" & vbCrLf & "���� " & YY & " �� " & MM & " �� " & DD & " �� " & HH & " ʱ"

    EventList(23).Name = "һֻ�����Ա���"
    EventList(23).Text = "�������Ի���Ϊ���͡�һ���м��ص���һֻ���Ӳ�֪��Ϊʲô�������Ա����߲�ʹ���ˣ��������Ѿ������˺쳾�������������û�����壬����������Ϊ��û�����塣���Ǹ��п��ܵ�һ��������ǣ������⵱����ʲôRPGģʽ��GalGame����BOSS���ܼ��ɹۿ�CG������Ȼ�������ﲻ���ܿ����κ�CG����Ϊ��ֻ�Ǹ�������Ϸ��������ͼ��û�У������������ַ�������������ɳ����Ϊ���ɿ��Ա���������Ի���Ϊ���ͣ�����ɵ�϶�֪���ⲻ����ʲô�ý����ʲô����һ��˵�������ֶ�������������һСһ����Ҫ�ĸ����ĸ��������㣡" & vbCrLf & "���� " & YY & " �� " & MM & " �� " & DD & " �� " & HH & " ʱ"
    
    Dim i As Integer
    For i = 100 To EventLimit
        EventList(i).Locked = True
    Next i
    
    EventList(100).Name = "�ӵ���һ��С����"
    EventList(100).Text = "ֻ��һЩ�򵥵�Ҫ�󣬺ܿ������ɣ����Ứ��̫�󹦷�ģ�"
    EventList(100).doYes_mp = -2
    EventList(100).doYes_mn = 20
    EventList(100).doYes_pt = 1
    EventList(100).doYes_ep = 2
    EventList(100).doYes_tm = 1
    EventList(100).doNo_mp = 1
    EventList(100).doNo_tm = 1
    
    EventList(101).Name = "ѧϰһЩС����"
    EventList(101).Text = "��Ѫ������ѧ�㴴��С���ɡ�"
    EventList(101).doYes_mp = -5
    EventList(101).doYes_ep = 10
    EventList(101).doYes_tm = 1
    EventList(101).doNo_mp = 1
    EventList(101).doNo_tm = 1
    
    EventList(102).Name = "������һ�ݲ�С�Ķ���"
    EventList(102).Text = "һ����֯��������һ����Ϊ��%worksname%������Ʒ�������Ҳȥ����һ�ѣ��������̵���С���ջ�"
    EventList(102).doYes_mp = -20
    EventList(102).doYes_mn = 1000
    EventList(102).doYes_pt = 50
    EventList(102).doYes_ep = 200
    EventList(102).doYes_tm = 9
    EventList(102).doNo_mp = 1
    EventList(102).doNo_tm = 1
    
    EventList(103).Name = "Ҫ��Ҫ��Ϯ���˵���Ʒ��"
    EventList(103).Text = "�㿴����һλ�������Ʒ��ͻȻ����а���Ҫ�����һ�¡�����"
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




