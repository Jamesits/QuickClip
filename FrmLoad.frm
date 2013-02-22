VERSION 5.00
Begin VB.Form FrmLoad 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "QuickClip加载中……"
   ClientHeight    =   1875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6645
   Icon            =   "FrmLoad.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "QuickClip"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6600
   End
End
Attribute VB_Name = "FrmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'显示加载窗体
Option Explicit
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Sub Form_Initialize()
InitCommonControls '应用XP风格控件
End Sub

Private Sub Form_Load()
Me.Show
Me.Refresh
GetSystemInf
GetDeviceInf APPPath(), , IsRunningOnRemovableDevice
iniFileName = APPPath("QuickClip.ini")
LoadFromINI
processSettingsOnStartup
Log_Open
frmMain.Show
Unload Me
End Sub
