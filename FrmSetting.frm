VERSION 5.00
Begin VB.Form FrmSetting 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "QuickClip设置"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13665
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   13665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame FrmRecInf 
      BackColor       =   &H00FFFFFF&
      Caption         =   "同时写入日志信息"
      Height          =   1215
      Left            =   240
      TabIndex        =   41
      Top             =   3540
      Width           =   2415
      Begin VB.OptionButton OptTxtRecInf 
         BackColor       =   &H00FFFFFF&
         Caption         =   "不写入"
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   44
         Top             =   240
         Width           =   1635
      End
      Begin VB.OptionButton OptTxtRecInf 
         BackColor       =   &H00FFFFFF&
         Caption         =   "写入到文件头部"
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   43
         Top             =   540
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.OptionButton OptTxtRecInf 
         BackColor       =   &H00FFFFFF&
         Caption         =   "写入到文件末尾"
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   42
         Top             =   840
         Width           =   1695
      End
   End
   Begin VB.Frame FrmTxtFilter 
      BackColor       =   &H00FFFFFF&
      Caption         =   "过滤器选项"
      Height          =   1215
      Left            =   2880
      TabIndex        =   35
      Top             =   3540
      Width           =   3675
      Begin VB.TextBox TxtTextFilterMin 
         Height          =   375
         Left            =   1140
         TabIndex        =   37
         Text            =   "0"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox TxtTextFilterMax 
         Height          =   375
         Left            =   1140
         TabIndex        =   36
         Text            =   "0"
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label LblTxtFilterMin 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "最小字节数"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   300
         Width           =   900
      End
      Begin VB.Label LblTxtMaxValue 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "最大字节数"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   780
         Width           =   900
      End
      Begin VB.Label LblTextFilterTip 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "（0为不限制）"
         Height          =   255
         Left            =   2280
         TabIndex        =   38
         Top             =   780
         Width           =   1185
      End
   End
   Begin VB.Frame FrmDebugSetting 
      BackColor       =   &H00FFFFFF&
      Caption         =   "调试功能"
      Height          =   1695
      Left            =   6780
      TabIndex        =   28
      Top             =   3240
      Width           =   6795
      Begin VB.CheckBox ChkShowLog 
         BackColor       =   &H00FFFFFF&
         Caption         =   "在主界面显示日志"
         Height          =   375
         Left            =   180
         TabIndex        =   34
         Top             =   1020
         Width           =   1755
      End
      Begin VB.TextBox TxtLogPath 
         Height          =   375
         Left            =   1320
         TabIndex        =   31
         Text            =   "QClipLog_%DATE%%TIME%.log"
         Top             =   600
         Width           =   4155
      End
      Begin VB.CheckBox ChkRecLog 
         BackColor       =   &H00FFFFFF&
         Caption         =   "记录日志"
         Height          =   315
         Left            =   180
         TabIndex        =   29
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "日志文件位置"
         Height          =   255
         Left            =   180
         TabIndex        =   30
         Top             =   660
         Width           =   1080
      End
   End
   Begin VB.Frame FrmProcessFiles 
      BackColor       =   &H00FFFFFF&
      Caption         =   "文件"
      Height          =   1515
      Left            =   6780
      TabIndex        =   19
      Top             =   1680
      Width           =   6795
      Begin VB.TextBox TxtFileFolderPath 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         TabIndex        =   26
         Text            =   "Files_%DATE%TIME%"
         Top             =   1080
         Width           =   4095
      End
      Begin VB.TextBox TxtFileSaveNameRule 
         Height          =   375
         Left            =   2520
         TabIndex        =   24
         Text            =   "QClip_%DATE%%TIME%.log"
         Top             =   660
         Width           =   4095
      End
      Begin VB.OptionButton OptProcessFile 
         BackColor       =   &H00FFFFFF&
         Caption         =   "仅保存文件路径"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   22
         Top             =   720
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton OptProcessFile 
         BackColor       =   &H00FFFFFF&
         Caption         =   "复制文件"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   21
         Top             =   1140
         Width           =   1095
      End
      Begin VB.OptionButton OptProcessFile 
         BackColor       =   &H00FFFFFF&
         Caption         =   "忽略"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   20
         Top             =   300
         Width           =   2235
      End
      Begin VB.Label LblCopyTo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "复制到文件夹"
         Height          =   255
         Left            =   1320
         TabIndex        =   25
         Top             =   1140
         Width           =   1080
      End
      Begin VB.Label LblFileSaveTo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "保存到"
         Height          =   255
         Left            =   1860
         TabIndex        =   23
         Top             =   720
         Width           =   540
      End
   End
   Begin VB.Frame FrmBitmap 
      BackColor       =   &H00FFFFFF&
      Caption         =   "位图类型"
      Height          =   1575
      Left            =   6780
      TabIndex        =   15
      Top             =   60
      Width           =   6795
      Begin VB.TextBox TxtBmpNamingRule 
         Height          =   375
         Left            =   960
         TabIndex        =   17
         Text            =   "QClip_%DATE%%TIME%.bmp"
         Top             =   660
         Width           =   5535
      End
      Begin VB.CheckBox ChkSavebmp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "保存"
         Height          =   315
         Left            =   180
         TabIndex        =   16
         Top             =   300
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.Label LblBmpNamingRule 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "命名规则"
         Height          =   255
         Left            =   180
         TabIndex        =   18
         Top             =   720
         Width           =   720
      End
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   435
      Left            =   12720
      TabIndex        =   14
      Top             =   5100
      Width           =   855
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "保存"
      Height          =   435
      Left            =   11640
      TabIndex        =   13
      Top             =   5100
      Width           =   915
   End
   Begin VB.Frame FrmCommon 
      BackColor       =   &H00FFFFFF&
      Caption         =   "通用设置"
      Height          =   1575
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   6675
      Begin VB.CheckBox ChkClearBeforeClose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "关闭软件时清空剪贴板内容"
         Height          =   315
         Left            =   2520
         TabIndex        =   32
         Top             =   1020
         Width           =   3135
      End
      Begin VB.CheckBox ChkDisableUSBMode 
         BackColor       =   &H00FFFFFF&
         Caption         =   "强制禁用可移动磁盘优化（实验性功能）"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2520
         TabIndex        =   27
         Top             =   720
         Width           =   3555
      End
      Begin VB.CheckBox ChkAutoStart 
         BackColor       =   &H00FFFFFF&
         Caption         =   "启动后自动开始监视"
         Height          =   315
         Left            =   180
         TabIndex        =   10
         Top             =   1020
         Width           =   2055
      End
      Begin VB.CheckBox ChkHideMode 
         BackColor       =   &H00FFFFFF&
         Caption         =   "启动后隐藏主界面"
         Enabled         =   0   'False
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton CmdBrowse 
         Caption         =   "浏览"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5880
         TabIndex        =   5
         Top             =   300
         Width           =   675
      End
      Begin VB.TextBox TxtPlace 
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Text            =   "%APPPATH%\Saved\%DATE%%TIME%"
         Top             =   300
         Width           =   4935
      End
      Begin VB.Label LblSavePath 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "保存位置"
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.Frame FrmText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "文本类型"
      Height          =   3255
      Left            =   60
      TabIndex        =   0
      Top             =   1680
      Width           =   6675
      Begin VB.TextBox TxtMergeSeparator 
         Height          =   375
         Left            =   900
         TabIndex        =   12
         Text            =   "------------------------------------------------"
         Top             =   1440
         Width           =   5655
      End
      Begin VB.CheckBox ChkTextOneFile 
         BackColor       =   &H00FFFFFF&
         Caption         =   "合并到第一个文件中"
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox TxtTextNamingRule 
         Height          =   375
         Left            =   1020
         TabIndex        =   8
         Text            =   "QClip_%DATE%%TIME%.txt"
         Top             =   660
         Width           =   5535
      End
      Begin VB.CheckBox ChkSaveText 
         BackColor       =   &H00FFFFFF&
         Caption         =   "保存"
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   300
         Value           =   1  'Checked
         Width           =   675
      End
      Begin VB.Label LblTextSeperator 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "分隔符"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   540
      End
      Begin VB.Label LblTextSuffix 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "命名规则"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   720
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "注意：您所做的某些修改会在QuickClip下次启动时生效。"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6900
      TabIndex        =   33
      Top             =   5160
      Width           =   4605
   End
End
Attribute VB_Name = "FrmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'设置
Option Explicit

Private Sub CmdCancel_Click()
Unload Me
End Sub

Private Sub CmdSave_Click()
'If isCapturing = True Then
'log "保存设置，重新启动监视……"
'StopCapture
'SaveSettings
'SaveToINI
'StartCapture
'Else
SaveSettings
SaveToINI
'End If
Unload Me
End Sub

Private Sub Form_Load()
LoadSettings
End Sub

Private Sub LoadSettings()
TxtPlace.Text = Common_SavePath
ChkHideMode.Value = Common_HideAfterStart
ChkAutoStart.Value = Common_AutoCapture
ChkDisableUSBMode.Value = Common_DisableUSBCheck
ChkSaveText.Value = Text_Save
TxtTextNamingRule.Text = Text_Name
ChkTextOneFile.Value = Text_MergeFile
TxtMergeSeparator.Text = Text_MergeSeparator
OptTxtRecInf(Text_RecInformation).Value = True
TxtTextFilterMin.Text = Text_FilterMinBytes
TxtTextFilterMax.Text = Text_FilterMaxBytes
ChkSavebmp.Value = Bitmap_Save
TxtBmpNamingRule.Text = Bitmap_Name
OptProcessFile(File_Save).Value = True
TxtFileSaveNameRule.Text = File_LogPath
TxtFileFolderPath.Text = File_SaveFolder
ChkRecLog.Value = Log_Save
TxtLogPath.Text = Log_Filename
ChkShowLog.Value = Log_ShowLogAtFrmMain
End Sub

Private Sub SaveSettings()
Common_SavePath = TxtPlace.Text
Common_HideAfterStart = ChkHideMode.Value
Common_AutoCapture = ChkAutoStart.Value
Common_DisableUSBCheck = ChkDisableUSBMode.Value
Text_Save = ChkSaveText.Value
Text_Name = TxtTextNamingRule.Text
Text_MergeFile = ChkTextOneFile.Value
Text_MergeSeparator = TxtMergeSeparator.Text
Text_RecInformation = GetOptionGroupValue(OptTxtRecInf, 0, 2)
Text_FilterMinBytes = TxtTextFilterMin.Text
Text_FilterMaxBytes = TxtTextFilterMax.Text
Bitmap_Save = ChkSavebmp.Value
Bitmap_Name = TxtBmpNamingRule.Text
File_Save = GetOptionGroupValue(OptProcessFile, 0, 2)
File_LogPath = TxtFileSaveNameRule.Text
File_SaveFolder = TxtFileFolderPath.Text
Log_Save = ChkRecLog.Value
Log_Filename = TxtLogPath.Text
Log_ShowLogAtFrmMain = ChkShowLog.Value
End Sub


Private Function GetOptionGroupValue(ByVal Options, ByVal LB As Byte, ByVal UB As Byte) As Long
On Error Resume Next
Dim i As Integer
For i = LB To UB
If Options(i).Value = True Then
    GetOptionGroupValue = i
    Exit Function
    End If
Next
GetOptionGroupValue = 0
Exit Function
End Function

Private Sub TxtTextFilterMax_Change()
Static last As Long
If Trim(TxtTextFilterMax.Text) <> "" Then
On Error GoTo e
TxtTextFilterMax.Text = CLng(TxtTextFilterMax.Text)
last = CLng(TxtTextFilterMax.Text)
End If
Exit Sub
e:
TxtTextFilterMax.Text = last
End Sub

Private Sub TxtTextFilterMin_Change()
Static last As Long
If Trim(TxtTextFilterMin.Text) <> "" Then
On Error GoTo e
TxtTextFilterMin.Text = CLng(TxtTextFilterMin.Text)
last = CLng(TxtTextFilterMin.Text)
End If
Exit Sub
e:
TxtTextFilterMin.Text = last
End Sub
