VERSION 5.00
Begin VB.Form frm 
   Caption         =   "Form1"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11055
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   11055
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Frame fraEvent 
      Caption         =   "ÊÂ¼þ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   10335
      Begin VB.CommandButton cmdS3 
         Caption         =   "Ñ¡ÔñÈý"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         TabIndex        =   10
         Top             =   4560
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton cmdS2 
         Caption         =   "Ñ¡Ôñ¶þ"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   9
         Top             =   4560
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton cmdS1 
         Caption         =   "Ñ¡ÔñÒ»"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   4560
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton cmdNo 
         Caption         =   "¹¾"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6480
         TabIndex        =   7
         Top             =   4560
         Width           =   1815
      End
      Begin VB.CommandButton cmdYes 
         Caption         =   "²»¹¾"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8400
         TabIndex        =   6
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label labEvT 
         Caption         =   "ÊÂ¼þ±êÌâ"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   9855
      End
      Begin VB.Label labEvent 
         Caption         =   "ÕâÀï½«»áÃèÊöÊÂ¼þÄÚÈÝ¡£"
         Height          =   3615
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   9855
      End
   End
   Begin VB.PictureBox picHp 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      ScaleHeight     =   1
      ScaleMode       =   0  'User
      ScaleWidth      =   99
      TabIndex        =   1
      Top             =   240
      Width           =   2655
      Begin VB.Shape sHp 
         BackColor       =   &H0033AE2D&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   1311
      End
   End
   Begin VB.PictureBox picMp 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      ScaleHeight     =   1
      ScaleMode       =   0  'User
      ScaleWidth      =   99
      TabIndex        =   0
      Top             =   720
      Width           =   2655
      Begin VB.Shape sMp 
         BackColor       =   &H00DBAE53&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   1311
      End
   End
   Begin VB.Label labMn 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1500"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label labPt 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   375
      Left            =   8520
      TabIndex        =   19
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Áé¸Ð"
      Height          =   375
      Left            =   7560
      TabIndex        =   20
      Top             =   720
      Width           =   975
   End
   Begin VB.Label labDate 
      Height          =   375
      Left            =   360
      TabIndex        =   18
      Top             =   6960
      Width           =   10335
   End
   Begin VB.Label labEp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   375
      Left            =   8520
      TabIndex        =   17
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label labEpT 
      Caption         =   "×ÊÀú"
      Height          =   375
      Left            =   7560
      TabIndex        =   16
      Top             =   240
      Width           =   975
   End
   Begin VB.Label labPtT 
      Caption         =   "ÉùÍû"
      Height          =   375
      Left            =   4200
      TabIndex        =   15
      Top             =   720
      Width           =   975
   End
   Begin VB.Label labMnT 
      Caption         =   "×Ê½ð"
      Height          =   375
      Left            =   4200
      TabIndex        =   14
      Top             =   240
      Width           =   975
   End
   Begin VB.Label labMpT 
      Caption         =   "ÌåÁ¦"
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   720
      Width           =   975
   End
   Begin VB.Label labHpT 
      Caption         =   "½¡¿µ"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdYes_Click()
    If GameOver = True Then Exit Sub
    CgHP dY_hp
    CgMP dY_mp
    CgMN dY_mn
    CgPT dY_pt
    CgEP dY_ep
    DateAdd dY_tm
    stsLoad
    EventLock "y"
End Sub

Private Sub cmdNo_Click()
    If GameOver = True Then Exit Sub
    CgHP dN_hp
    CgMP dN_mp
    CgMN dN_mn
    CgPT dN_pt
    CgEP dN_ep
    DateAdd dN_tm
    stsLoad
    EventLock "n"
End Sub

Private Sub Form_Load()
    sts_mp_max = 100
    sts_hp_max = 100
    sts_mp = 100
    sts_hp = 100
    sts_mn = 0
    sts_pt = 0
    
    sHp.Width = sts_hp
    sMp.Width = sts_mp
    
    stsLoad
    EventLoad
    WorksLoad
    EventSet 1
End Sub

