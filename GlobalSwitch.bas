Attribute VB_Name = "GlobalSwitch"
Option Explicit
Public isCapturing As Boolean

Public Sub StartCapture()
isCapturing = True
frmMain.CmdStartMonitor.Caption = "ֹͣ����"
FrmCatchMsg.objSC.AddWindowMsgs FrmCatchMsg.hwnd
log "�����ѿ�ʼ"
End Sub

Public Sub StopCapture()
isCapturing = False
frmMain.CmdStartMonitor.Caption = "��ʼ����"
FrmCatchMsg.objSC.DeleteWindowMsg FrmCatchMsg.hwnd
log "������ֹͣ"
End Sub

Public Sub ProcessChange() '���������ı�
If Clipboard.GetFormat(vbCFText) Then
    log "��⵽��������� �������ͣ��ı�"
    processText
    ElseIf Clipboard.GetFormat(vbCFBitmap) Then log "��⵽��������� �������ͣ�λͼ��bmp���ļ�": processBitmap
    ElseIf Clipboard.GetFormat(vbCFMetafile) Then log "��⵽��������� �������ͣ�ͼԪ��wmf���ļ�": processwmf
    ElseIf Clipboard.GetFormat(vbCFDIB) Then log "��⵽��������� �������ͣ��豸�޹�λͼ��DIB���ļ�": processDIB
    ElseIf Clipboard.GetFormat(vbCFPalette) Then log "��⵽��������� �������ͣ���ɫ������": processPalette
    ElseIf Clipboard.GetFormat(vbCFLink) Then log "��⵽��������� �������ͣ�DDE�Ի���Ϣ": processDDE
    ElseIf Clipboard.GetFormat(vbCFFiles) Then log "��⵽��������� �������ͣ��ļ��б�": processFileList
    ElseIf Clipboard.GetFormat(vbCFRTF) Then log "��⵽��������� �������ͣ����ı���RTF���ļ�": processRTF
    ElseIf Clipboard.GetFormat(vbCFEMetafile) Then log "��⵽��������� �������ͣ���ǿ��ͼԪ�ļ���EMF��": processEMF
    Else: log "��⵽��������� �������ͣ�δ֪": processUnknownValue
    End If
End Sub


