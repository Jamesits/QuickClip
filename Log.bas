Attribute VB_Name = "LogRecorder"
Option Explicit

Dim Log_isOpen As Boolean
Dim Log_FileNum As Integer

Public Sub Log_Write(ByVal str As String)
If Log_Save Then
If Not (Log_isOpen) Then Log_Open
Print #1, str
Log_Close
End If
End Sub

Public Sub Log_Open()
If Log_Save Then
Log_FileNum = FreeFile()
Open LogFileName For Append Access Write Lock Write As #Log_FileNum
Log_isOpen = True
End If
End Sub

Public Sub Log_Close()
If Log_Save Or Log_isOpen Then
Close #Log_FileNum
Log_isOpen = False
End If
End Sub

Public Sub Log(ByVal str As String, Optional showtime As Boolean = True, Optional appendReturn As Boolean = True)
Dim Temps As String
If showtime Then
Temps = "[" & Time & "] " & str
Else: Temps = str & vbCrLf
End If
If appendReturn Then frmMain.Txtlog.Text = frmMain.Txtlog.Text & Temps & vbCrLf Else frmMain.Txtlog.Text = frmMain.Txtlog.Text & Temps
If Log_Save Then
Log_Write Temps
End If
End Sub

