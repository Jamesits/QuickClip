Attribute VB_Name = "modParseClipboard"
Option Explicit
Public Const CF_HDROP = 15
Public Type POINT
  X As Long
  Y As Long
End Type
Public Type DROPFILES
  pFiles As Long
  pt As POINT
  fNC As Long
  fWide As Long
End Type
Public Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function SetClipboardViewer Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Const WM_DRAWCLIPBOARD As Long = &H308

Sub processUnknownValue()
commonUnsupported False, Clipboard.GetData
End Sub

Sub processRTF()
commonUnsupported False, Clipboard.GetData
End Sub

Sub processEMF()
commonUnsupported False, Clipboard.GetData
End Sub

Sub processFileList()
  Dim lHandle As Long
  Dim lpResults As Long
  Dim lRet As Long
  Dim df As DROPFILES
  Dim strDest As String
  Dim lBufferSize As Long
  Dim arBuffer() As Byte
  Dim vNames As Variant
  Dim i As Long
  If OpenClipboard(0) Then
   lHandle = GetClipboardData(CF_HDROP)
   ' If you don't find a CF_HDROP, you don't want to process anything
   If lHandle > 0 Then
     lpResults = GlobalLock(lHandle)
    
     lBufferSize = GlobalSize(lpResults)
     ReDim arBuffer(0 To lBufferSize)
    
     CopyMemory df, ByVal lpResults, Len(df)
     Call CopyMemory(arBuffer(0), ByVal lpResults + df.pFiles, _
             (lBufferSize - Len(df)))
     If df.fWide = 1 Then
      ' it is wide chars--unicode
      strDest = arBuffer
     Else
      strDest = StrConv(arBuffer, vbUnicode)
     End If
     GlobalUnlock lHandle
     vNames = Split(strDest, vbNullChar)
     i = 0
     While Len(vNames(i)) > 0
      vNames(i) = Replace(vNames(i), Chr(10), "")
      vNames(i) = Replace(vNames(i), Chr(13), "")
      processFilename vNames(i)
      i = i + 1
     Wend
     Log "文件总数：" & i, False, False
   End If
  End If
  CloseClipboard
End Sub

Sub processFilename(ByVal name As String)
Log "文件：" & name, False, False
End Sub

Sub processBitmap()
If Bitmap_Save = 1 Then
frmMain.Picture1.Picture = Clipboard.GetData
SavePicture frmMain.Picture1.Picture, processString(Bitmap_Name)
End If
End Sub

Sub processText()
Dim Temps As String
Dim Log As String
Static Filename As String
Temps = Clipboard.GetText
'Log "文本："
'Log Temps, False
If Text_MergeFile = False Or Filename = "" Then Filename = processString(Text_Name)
If Text_RecInformation <> 0 Then Log = "QuickClip文本记录" & vbCrLf & "时间：" & Format(Now(), "yyyy年mm月dd日hh时mm分ss秒") & vbCrLf & "内容：" & vbCrLf
If Text_MergeFile = False Then
    processText_NewFile Temps, Filename, Log
    Else
    processText_Merge Temps, Filename, Log
End If
End Sub

Private Sub processText_NewFile(content As String, Filename As String, Log As String)
Dim File As Integer
File = FreeFile()
Open Filename For Append Access Write Lock Write As #File
If Text_RecInformation = 1 Then Print #File, Log
Print #File, content
If Text_RecInformation = 2 Then Print #File, Log
Close #File
End Sub

Private Sub processText_Merge(content As String, Filename As String, Log As String)
Static isFirstTime As Boolean '注意意思和值是反的……
Dim File As Integer
File = FreeFile()
Open Filename For Append Access Write Lock Write As #File
If isFirstTime = False Then
Print #File, Text_MergeSeparator
isFirstTime = True
End If
If Text_RecInformation = 1 Then Print #File, Log
Print #File, content
If Text_RecInformation = 2 Then Print #File, Log
Close #File
End Sub

Sub processwmf()
commonUnsupported False, Clipboard.GetData
End Sub

Sub processDIB()
commonUnsupported False, Clipboard.GetData
End Sub

Sub processPalette()
commonUnsupported False, Clipboard.GetData
End Sub

Sub processDDE()
commonUnsupported False, Clipboard.GetData
End Sub

Private Sub commonUnsupported(Optional ByVal savedata As Boolean = False, Optional ByVal data)
Log "暂时无法处理此类数据。"
End Sub

