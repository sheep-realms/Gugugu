Attribute VB_Name = "mdlMain"
Option Explicit

Public GameOver As Boolean

Public sts_mp As Integer, sts_hp As Integer, sts_mn As Long, sts_pt As Integer, sts_Ep As Integer

Public sts_mp_max As Integer, sts_hp_max As Integer

Public doEvent As Integer

Public doNightSleep As Boolean
Public NoSleepDays As Integer

Public stc_CopyLimit As Integer

Public dY_mp As Integer, dY_hp As Integer, dY_mn As Long, dY_pt As Integer, dY_ep As Integer, dY_tm As Integer
Public dN_mp As Integer, dN_hp As Integer, dN_mn As Long, dN_pt As Integer, dN_ep As Integer, dN_tm As Integer

Public Function stsLoad()
    frm.sHp.Width = sts_hp
    frm.sMp.Width = sts_mp
    frm.labMn.Caption = sts_mn
    frm.labPt.Caption = sts_pt
    frm.labEp.Caption = sts_Ep
End Function

Public Function SetHP(Value As Integer)
    sts_hp = Value
    stsLoad
End Function

Public Function SetMP(Value As Integer)
    sts_mp = Value
    stsLoad
End Function

Public Function CgHP(Value As Integer)
    sts_hp = sts_hp + Value
    If sts_hp > sts_hp_max Then sts_hp = sts_hp_max
    If sts_hp < 0 Then sts_hp = 0
    stsLoad
End Function

Public Function CgMP(Value As Integer)
    sts_mp = sts_mp + Value
    If sts_mp > sts_mp_max Then sts_mp = sts_mp_max
    If sts_mp < 0 Then sts_mp = 0
    stsLoad
End Function

Public Function CgMN(Value As Long)
    sts_mn = sts_mn + Value
    stsLoad
End Function

Public Function CgPT(Value As Integer)
    sts_pt = sts_pt + Value
    stsLoad
End Function

Public Function CgEP(Value As Integer)
    sts_Ep = sts_Ep + Value
    If sts_Ep < 0 Then sts_Ep = 0
    stsLoad
End Function

Public Function EventSet(Value As Integer)
    Dim str As String
    str = EventList(Value).Text & vbCrLf & vbCrLf
    If InStr(str, "%worksname%") <> 0 Then
        str = Replace(str, "%worksname%", WorksRnd)
    End If

    frm.labEvT.Caption = EventList(Value).Name
    frm.labEvent.Caption = str
    
    str = ""
    
    doEvent = Value
    
    If GameOver = True Then Exit Function
    
    BuffLoad
    
    dY_hp = EventList(Value).doYes_hp + bf_hp + ebf_hp
    dY_mp = EventList(Value).doYes_mp + bf_mp + ebf_mp
    dY_mn = EventList(Value).doYes_mn + bf_mn + ebf_mn
    dY_pt = EventList(Value).doYes_pt + bf_pt + ebf_pt
    dY_ep = EventList(Value).doYes_ep + bf_ep + ebf_ep
    dY_tm = EventList(Value).doYes_tm
    dN_hp = EventList(Value).doNo_hp + dbf_hp + ebf_hp
    dN_mp = EventList(Value).doNo_mp + dbf_mp + ebf_mp
    dN_mn = EventList(Value).doNo_mn + dbf_mn + ebf_mn
    dN_pt = EventList(Value).doNo_pt + dbf_pt + ebf_pt
    dN_ep = EventList(Value).doNo_ep + dbf_ep + ebf_ep
    dN_tm = EventList(Value).doNo_tm
    
    
    If dY_hp <> 0 Then
        str = str & "健康 "
        If dY_hp > 0 Then str = str & "+"
        str = str & dY_hp & "  "
    End If
    If dY_mp <> 0 Then
        str = str & "体力 "
        If dY_mp > 0 Then str = str & "+"
        str = str & dY_mp & "  "
    End If
    If dY_mn <> 0 Then
        str = str & "资金 "
        If dY_mn > 0 Then str = str & "+"
        str = str & dY_mn & "  "
    End If
    If dY_pt <> 0 Then
        str = str & "声望 "
        If dY_pt > 0 Then str = str & "+"
        str = str & dY_pt & "  "
    End If
    If dY_ep <> 0 Then
        str = str & "资历 "
        If dY_ep > 0 Then str = str & "+"
        str = str & dY_ep & "  "
    End If
    If dY_tm <> 0 Then
        str = str & "耗时 "
        If dY_tm > 0 Then str = str & "+"
        str = str & dY_tm & "  "
    End If
    If str <> "" Then str = "不咕 = " & str
    
    frm.labEvent.Caption = frm.labEvent.Caption & str
    
    str = ""
    
    If dN_hp <> 0 Then
        str = str & "健康 "
        If dN_hp > 0 Then str = str & "+"
        str = str & dN_hp & "  "
    End If
    If dN_mp <> 0 Then
        str = str & "体力 "
        If dN_mp > 0 Then str = str & "+"
        str = str & dN_mp & "  "
    End If
    If dN_mn <> 0 Then
        str = str & "资金 "
        If dN_mn > 0 Then str = str & "+"
        str = str & dN_mn & "  "
    End If
    If dN_pt <> 0 Then
        str = str & "声望 "
        If dN_pt > 0 Then str = str & "+"
        str = str & dN_pt & "  "
    End If
    If dN_ep <> 0 Then
        str = str & "资历 "
        If dN_ep > 0 Then str = str & "+"
        str = str & dN_ep & "  "
    End If
    If dN_tm <> 0 Then
        str = str & "耗时 "
        If dN_tm > 0 Then str = str & "+"
        str = str & dN_tm & "  "
    End If
    If str <> "" Then str = "咕了 = " & str
    frm.labEvent.Caption = frm.labEvent.Caption & vbCrLf & str
End Function
