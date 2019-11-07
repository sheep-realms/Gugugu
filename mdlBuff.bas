Attribute VB_Name = "mdlBuff"
Option Explicit

Public ebf_hp As Integer, ebf_mp As Integer, ebf_mn As Integer, ebf_pt As Integer, ebf_ep As Integer
Public bf_hp As Integer, bf_mp As Integer, bf_mn As Integer, bf_pt As Integer, bf_ep As Integer
Public dbf_hp As Integer, dbf_mp As Integer, dbf_mn As Integer, dbf_pt As Integer, dbf_ep As Integer

Public Function BuffLoad()
    
    ebf_mp = -(NoSleepDays)
    If (doEvent = 10) And (sts_hp < 10) Then
        bf_hp = 10
        bf_mp = -20
    ElseIf (doEvent = 10) And (sts_hp < 25) Then
        bf_hp = 10
        bf_mp = -15
    ElseIf (doEvent = 10) And (sts_hp < 50) Then
        bf_hp = 5
        bf_mp = -10
    ElseIf (doEvent = 10) And (sts_hp < 80) Then
        bf_mp = -5
    Else
        bf_hp = 0
        bf_mp = 0
    End If
    If (doEvent = 10) And (NoSleepDays <> 0) Then
        dbf_hp = -(NoSleepDays * 15)
    Else
        dbf_hp = 0
    End If
End Function


