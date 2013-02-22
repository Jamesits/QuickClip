VERSION 5.00
Begin VB.Form FrmLog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QuickClip日志"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11610
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   11610
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame FrmLog 
      BackColor       =   &H00FFFFFF&
      Caption         =   "调试日志"
      Height          =   5115
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   11475
      Begin VB.TextBox Txtlog 
         Height          =   4650
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   300
         Width           =   11115
      End
   End
End
Attribute VB_Name = "FrmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

