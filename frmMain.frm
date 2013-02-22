VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QuickClip"
   ClientHeight    =   765
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4170
   BeginProperty Font 
      Name            =   "΢���ź�"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   4170
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton CmdStartMonitor 
      Caption         =   "��ʼ����"
      Height          =   435
      Left            =   120
      TabIndex        =   5
      Top             =   180
      Width           =   1035
   End
   Begin VB.CommandButton CmdQuit 
      Caption         =   "�˳�"
      Height          =   435
      Left            =   3420
      TabIndex        =   4
      Top             =   180
      Width           =   615
   End
   Begin VB.CommandButton CmdAbout 
      Caption         =   "����"
      Height          =   435
      Left            =   2700
      TabIndex        =   3
      Top             =   180
      Width           =   615
   End
   Begin VB.CommandButton CmdHide 
      Caption         =   "����"
      Enabled         =   0   'False
      Height          =   435
      Left            =   1980
      TabIndex        =   2
      Top             =   180
      Width           =   615
   End
   Begin VB.CommandButton CmdSetting 
      Caption         =   "����"
      Height          =   435
      Left            =   1260
      TabIndex        =   1
      Top             =   180
      Width           =   615
   End
   Begin VB.Label LblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QuickClip"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   6060
      TabIndex        =   0
      Top             =   -120
      Visible         =   0   'False
      Width           =   3315
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'������
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
    Log "QuickClip �汾" & ver
    'If IsRunningOnRemovableDevice Then Log "���ƶ������Ż�������"
    LblTitle.Caption = "QuickClip " & ver
    Load FrmCatchMsg
    If Common_AutoCapture Then CmdStartMonitor_Click
    If Log_ShowLogAtFrmMain Then FrmLog.Show
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Unload FrmCatchMsg
    Log "QuickClip�˳�"
    Log_Close
    End
End Sub


