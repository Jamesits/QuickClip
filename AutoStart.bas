Attribute VB_Name = "AutoStart"
Option Explicit

'提权执行 CreateObject("Shell.Application").ShellExecute "文件名", "参数", "", "RunAs"
'wscript.shell 修改注册表方法：
'CreateObject("Wscript.Shell").RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\QuickClip")
'CreateObject("wscript.shell").regwrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\QuickClip", APPPath("QuickClip.exe")
'CreateObject("wscript.shell").regdelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\QuickClip"



Public Sub SetAutorun()
'CreateObject("wscript.shell").regwrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\QuickClip", APPPath("QuickClip.exe")
End Sub

Public Sub CancelAutorun()
'CreateObject("wscript.shell").regdelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\QuickClip"
End Sub

Public Function isAutorun() As Boolean
'If CreateObject("Wscript.Shell").RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\QuickClip") <> "" Then isAutorun = True Else isAutorun = False
End Function

Public Sub Copyme() '对USB存储器，复制自身到系统目录去

End Sub
