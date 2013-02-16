VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QuickClip"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11715
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   11715
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdStartMonitor 
      Caption         =   "开始监视"
      Height          =   435
      Left            =   7680
      TabIndex        =   8
      Top             =   300
      Width           =   1035
   End
   Begin VB.CommandButton CmdQuit 
      Caption         =   "退出"
      Height          =   435
      Left            =   10980
      TabIndex        =   7
      Top             =   300
      Width           =   615
   End
   Begin VB.CommandButton CmdAbout 
      Caption         =   "关于"
      Height          =   435
      Left            =   10260
      TabIndex        =   6
      Top             =   300
      Width           =   615
   End
   Begin VB.CommandButton CmdHide 
      Caption         =   "隐藏"
      Height          =   435
      Left            =   9540
      TabIndex        =   5
      Top             =   300
      Width           =   615
   End
   Begin VB.CommandButton CmdSetting 
      Caption         =   "设置"
      Height          =   435
      Left            =   8820
      TabIndex        =   4
      Top             =   300
      Width           =   615
   End
   Begin VB.Frame FrmLog 
      BackColor       =   &H00FFFFFF&
      Caption         =   "调试日志"
      Height          =   5115
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   11655
      Begin VB.TextBox Txtlog 
         Height          =   4590
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   420
         Width           =   11355
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   12780
      ScaleHeight     =   4275
      ScaleWidth      =   7215
      TabIndex        =   0
      Top             =   3480
      Visible         =   0   'False
      Width           =   7275
   End
   Begin VB.Label LblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QuickClip"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3315
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CmdAbout_Click()
FrmAbout.Show
End Sub

Private Sub CmdQuit_Click()
Unload Me
End
End Sub

Private Sub CmdSetting_Click()
FrmSetting.Show
End Sub

Private Sub CmdStartMonitor_Click()
Static onMonitor As Boolean
If onMonitor = True Then StopCapture Else StartCapture
onMonitor = Not (onMonitor)
End Sub

Private Sub Form_Load()
    Dim ver As String
    ver = GetFileVerInfo(APPPath("QuickClip.exe"))(0)
    log "QuickClip 版本" & ver
    If IsRunningOnRemovableDevice Then log "可移动磁盘优化已启用"
    LblTitle.Caption = "QuickClip " & ver
    Load FrmCatchMsg
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Unload FrmCatchMsg
    log "QuickClip退出"
    Log_Close
    End
End Sub


