Attribute VB_Name = "AutoStart"
Option Explicit

'��Ȩִ�� CreateObject("Shell.Application").ShellExecute "�ļ���", "����", "", "RunAs"
'wscript.shell �޸�ע�������
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

Public Sub Copyme() '��USB�洢������������ϵͳĿ¼ȥ

End Sub
