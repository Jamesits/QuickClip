VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QuickClip"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11700
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
   ScaleHeight     =   990
   ScaleWidth      =   11700
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdStartMonitor 
      Caption         =   "开始监视"
      Height          =   435
      Left            =   7680
      TabIndex        =   7
      Top             =   300
      Width           =   1035
   End
   Begin VB.CommandButton CmdQuit 
      Caption         =   "退出"
      Height          =   435
      Left            =   10980
      TabIndex        =   6
      Top             =   300
      Width           =   615
   End
   Begin VB.CommandButton CmdAbout 
      Caption         =   "关于"
      Height          =   435
      Left            =   10260
      TabIndex        =   5
      Top             =   300
      Width           =   615
   End
   Begin VB.CommandButton CmdHide 
      Caption         =   "隐藏"
      Enabled         =   0   'False
      Height          =   435
      Left            =   9540
      TabIndex        =   4
      Top             =   300
      Width           =   615
   End
   Begin VB.CommandButton CmdSetting 
      Caption         =   "设置"
      Height          =   435
      Left            =   8820
      TabIndex        =   3
      Top             =   300
      Width           =   615
   End
   Begin VB.Frame FrmLog 
      BackColor       =   &H00FFFFFF&
      Caption         =   "调试日志"
      Height          =   5115
      Left            =   120
      TabIndex        =   1
      Top             =   1020
      Width           =   11475
      Begin VB.TextBox Txtlog 
         Height          =   4650
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   300
         Width           =   11115
      End
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
      TabIndex        =   0
      Top             =   0
      Width           =   3315
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'主窗体
Option Explicit
Public ver As String
Private Sub CmdAbout_Click()
FrmAbout.Show vbModal
End Sub

Private Sub CmdQuit_Click()
Quit
End Sub

Private Sub CmdSetting_Click()
FrmSetting.Show vbModal
End Sub

Private Sub CmdStartMonitor_Click()
Static onMonitor As Boolean
If onMonitor = True Then StopCapture Else StartCapture
onMonitor = Not (onMonitor)
End Sub

Private Sub Form_Load()
    
    ver = App.Major & "." & Format(App.Minor, "00") & "." & Format(App.Revision, "0000") 'GetFileVerInfo(APPPath(App.EXEName))(0)
    Log "QuickClip 版本" & ver
    'If IsRunningOnRemovableDevice Then Log "可移动磁盘优化已启用"
    LblTitle.Caption = "QuickClip " & ver
    Load FrmCatchMsg
    If Common_AutoCapture Then CmdStartMonitor_Click
    If Log_ShowLogAtFrmMain Then Me.Height = 6645 Else Me.Height = 1410
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Unload FrmCatchMsg
    Log "QuickClip退出"
    Log_Close
    End
End Sub


