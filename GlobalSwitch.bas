Attribute VB_Name = "GlobalSwitch"
'ģ�飺ȫ�ּ��ӿ��غͼ���������ӳ���ĵ���
Option Explicit
Public isCapturing As Boolean

Public Sub StartCapture()
isCapturing = True
frmMain.CmdStartMonitor.Caption = "ֹͣ����"
FrmCatchMsg.objSC.AddWindowMsgs FrmCatchMsg.hwnd
Log "�����ѿ�ʼ"
End Sub

Public Sub StopCapture()
isCapturing = False
frmMain.CmdStartMonitor.Caption = "��ʼ����"
FrmCatchMsg.objSC.DeleteWindowMsg FrmCatchMsg.hwnd
Log "������ֹͣ"
End Sub

Public Sub ProcessChange() '���������ı�
If Clipboard.GetFormat(vbCFText) Then
    Log "��⵽��������� �������ͣ��ı�"
    processText
    ElseIf Clipboard.GetFormat(vbCFBitmap) Then Log "��⵽��������� �������ͣ�λͼ��bmp���ļ�": processBitmap
    ElseIf Clipboard.GetFormat(vbCFMetafile) Then Log "��⵽��������� �������ͣ�ͼԪ��wmf���ļ�": processwmf
    ElseIf Clipboard.GetFormat(vbCFDIB) Then Log "��⵽��������� �������ͣ��豸�޹�λͼ��DIB���ļ�": processDIB
    ElseIf Clipboard.GetFormat(vbCFPalette) Then Log "��⵽��������� �������ͣ���ɫ������": processPalette
    ElseIf Clipboard.GetFormat(vbCFLink) Then Log "��⵽��������� �������ͣ�DDE�Ի���Ϣ": processDDE
    ElseIf Clipboard.GetFormat(vbCFFiles) Then Log "��⵽��������� �������ͣ��ļ��б�": processFileList
    ElseIf Clipboard.GetFormat(vbCFRTF) Then Log "��⵽��������� �������ͣ����ı���RTF���ļ�": processRTF
    ElseIf Clipboard.GetFormat(vbCFEMetafile) Then Log "��⵽��������� �������ͣ���ǿ��ͼԪ�ļ���EMF��": processEMF
    Else: Log "��⵽��������� �������ͣ�δ֪": processUnknownValue
    End If
End Sub


