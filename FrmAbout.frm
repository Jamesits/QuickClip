VERSION 5.00
Begin VB.Form FrmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于QuickClip"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton CmdClose 
      Cancel          =   -1  'True
      Caption         =   "关闭"
      Height          =   435
      Left            =   3120
      TabIndex        =   0
      Top             =   2340
      Width           =   915
   End
   Begin VB.Label LblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "QuickClip 1.00.0000"
      Height          =   255
      Left            =   660
      TabIndex        =   3
      Top             =   300
      Width           =   1710
   End
   Begin VB.Label LblAuthor 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) 2013 James Swineson.All rights reserved."
      Height          =   255
      Left            =   660
      TabIndex        =   2
      Top             =   1320
      Width           =   4755
   End
   Begin VB.Label LblDescription 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "简洁、快速、方便的剪贴板自动记录软件。"
      Height          =   255
      Left            =   660
      TabIndex        =   1
      Top             =   720
      Width           =   3420
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
LblTitle.Caption = "QuickClip " & GetFileVerInfo(APPPath("QuickClip.exe"))(0)
Me.Show
End Sub

