VERSION 5.00
Begin VB.Form FrmCatchMsg 
   BorderStyle     =   0  'None
   Caption         =   "QuickClip��̨֧�ִ���"
   ClientHeight    =   210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   210
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   210
   ScaleWidth      =   210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Visible         =   0   'False
End
Attribute VB_Name = "FrmCatchMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents objSC As cSubclass
Attribute objSC.VB_VarHelpID = -1

Private Sub Form_Load()
    Call SetClipboardViewer(Me.hwnd)                    '��ӱ������������鿴���б�
    Set objSC = New cSubclass
End Sub

Private Sub objSC_MsgCome(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, lng_hWnd As Long, uMsg As Long, wParam As Long, lParam As Long)
    If bBefore Then
        Select Case uMsg
            Case WM_DRAWCLIPBOARD                               '�����屻�ı�
                    ProcessChange
        End Select
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    StopCapture
    Set objSC = Nothing
End Sub
