VERSION 5.00
Begin VB.Form FrmSavePicture 
   BorderStyle     =   0  'None
   Caption         =   "QuickClip_SavingPicture"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4335
      Left            =   0
      ScaleHeight     =   4275
      ScaleWidth      =   7215
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7275
   End
End
Attribute VB_Name = "FrmSavePicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Picture1.Picture = Clipboard.GetData
SavePicture Picture1.Picture, processString(Bitmap_Name)
Debug.Print "file saved!"
End Sub
