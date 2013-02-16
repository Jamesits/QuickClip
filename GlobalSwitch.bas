Attribute VB_Name = "GlobalSwitch"
'模块：全局监视开关和剪贴板监视子程序的调用
Option Explicit
Public isCapturing As Boolean

Public Sub StartCapture()
isCapturing = True
frmMain.CmdStartMonitor.Caption = "停止监视"
FrmCatchMsg.objSC.AddWindowMsgs FrmCatchMsg.hwnd
Log "监视已开始"
End Sub

Public Sub StopCapture()
isCapturing = False
frmMain.CmdStartMonitor.Caption = "开始监视"
FrmCatchMsg.objSC.DeleteWindowMsg FrmCatchMsg.hwnd
Log "监视已停止"
End Sub

Public Sub ProcessChange() '处理剪贴板改变
If Clipboard.GetFormat(vbCFText) Then
    Log "检测到剪贴板更新 数据类型：文本"
    processText
    ElseIf Clipboard.GetFormat(vbCFBitmap) Then Log "检测到剪贴板更新 数据类型：位图（bmp）文件": processBitmap
    ElseIf Clipboard.GetFormat(vbCFMetafile) Then Log "检测到剪贴板更新 数据类型：图元（wmf）文件": processwmf
    ElseIf Clipboard.GetFormat(vbCFDIB) Then Log "检测到剪贴板更新 数据类型：设备无关位图（DIB）文件": processDIB
    ElseIf Clipboard.GetFormat(vbCFPalette) Then Log "检测到剪贴板更新 数据类型：调色板数据": processPalette
    ElseIf Clipboard.GetFormat(vbCFLink) Then Log "检测到剪贴板更新 数据类型：DDE对话信息": processDDE
    ElseIf Clipboard.GetFormat(vbCFFiles) Then Log "检测到剪贴板更新 数据类型：文件列表": processFileList
    ElseIf Clipboard.GetFormat(vbCFRTF) Then Log "检测到剪贴板更新 数据类型：富文本（RTF）文件": processRTF
    ElseIf Clipboard.GetFormat(vbCFEMetafile) Then Log "检测到剪贴板更新 数据类型：增强型图元文件（EMF）": processEMF
    Else: Log "检测到剪贴板更新 数据类型：未知": processUnknownValue
    End If
End Sub


