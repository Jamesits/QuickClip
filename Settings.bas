Attribute VB_Name = "Settings"
Option Explicit
'格式常量
Const DateFormat As String = "yyyymmdd"
Const TimeFormat As String = "hhmmss"
'运行时检测项
Public IsRunningOnRemovableDevice As Boolean
'INI设置项
Public Common_SavePath As String
Public Common_HideAfterStart As Long
Public Common_AutoCapture As Long
Public Common_DisableUSBCheck As Long
Public Common_ClearBeforeClose As Long
Public Text_Save As Long
Public Text_Name As String
Public Text_MergeFile As Long
Public Text_MergeSeparator As String
Public Text_RecInformation As Long
Public Text_FilterMinBytes As Long
Public Text_FilterMaxBytes As Long
Public Bitmap_Save As Long
Public Bitmap_Name As String
Public File_Save As Long
Public File_LogPath As String
Public File_SaveFolder As String
Public Log_Save As Long
Public Log_Filename As String
'启动时处理项
Public SavePath As String '保存位置
Public LogFileName As String '日志文件位置

Public Sub processSettingsOnStartup()
SavePath = processString(Common_SavePath)
LogFileName = processString(Log_Filename)

End Sub

Public Sub SaveToINI()
SetIniS "Common", "SavePath", Common_SavePath
SetIniN "Common", "HideAfterStart", Common_HideAfterStart
SetIniN "Common", "AutoCapture", Common_AutoCapture
SetIniN "Common", "DisableUSBCheck", Common_DisableUSBCheck
SetIniN "common", "ClearBeforeClose", Common_ClearBeforeClose
SetIniN "Text", "Save", Text_Save
SetIniS "Text", "Name", Text_Name
SetIniN "Text", "MergeFile", Text_MergeFile
SetIniS "Text", "MergeSeparator", Text_MergeSeparator
SetIniN "Text", "RecInformation", Text_RecInformation 'GetOptionGroupValue(OptTxtRecInf)
SetIniN "Text", "FilterMinBytes", Text_FilterMinBytes
SetIniN "Text", "FilterMaxBytes", Text_FilterMaxBytes
SetIniN "Bitmap", "Save", Bitmap_Save
SetIniS "Bitmap", "Name", Bitmap_Name
SetIniN "File", "Save", File_Save 'GetOptionGroupValue(OptProcessFile)
SetIniS "File", "LogPath", File_LogPath
SetIniS "File", "SaveFolder", File_SaveFolder
SetIniN "Log", "Save", Log_Save
SetIniS "Log", "Name", Log_Filename
End Sub

Public Sub LoadFromINI()

GetIniS "Common", "SavePath", Common_SavePath, "%APPPATH%\Saved\%DATE%%TIME%"
GetIniN "Common", "HideAfterStart", Common_HideAfterStart, 0
GetIniN "Common", "AutoCapture", Common_AutoCapture, 0
GetIniN "Common", "DisableUSBCheck", Common_DisableUSBCheck, 0
GetIniN "common", "ClearBeforeClose", Common_ClearBeforeClose, 0
GetIniN "Text", "Save", Text_Save, 1
GetIniS "Text", "Name", Text_Name, "QClip_%DATE%%TIME%.txt"
GetIniN "Text", "MergeFile", Text_MergeFile, 0
GetIniS "Text", "MergeSeparator", Text_MergeSeparator, "------------------------------------------------"
GetIniN "Text", "RecInformation", Text_RecInformation, 0 'GetOptionGroupValue(OptTxtRecInf)
GetIniN "Text", "FilterMinBytes", Text_FilterMinBytes, 0
GetIniN "Text", "FilterMaxBytes", Text_FilterMaxBytes, 0
GetIniN "Bitmap", "Save", Bitmap_Save, 1
GetIniS "Bitmap", "Name", Bitmap_Name, "QClip_%DATE%%TIME%.bmp"
GetIniN "File", "Save", File_Save, 1 'GetOptionGroupValue(OptProcessFile)
GetIniS "File", "LogPath", File_LogPath, "QClip_%DATE%%TIME%.log"
GetIniS "File", "SaveFolder", File_SaveFolder, "Files_%DATE%TIME%"
GetIniN "Log", "Save", Log_Save, 0
GetIniS "Log", "Name", Log_Filename, "QClipLog_%DATE%%TIME%.log"
End Sub

Public Function processString(str As String)
processString = str
'日期时间
processString = Replace(processString, "%DATE%", Format(Date, DateFormat))
processString = Replace(processString, "%TIME%", Format(Time, TimeFormat))
'本应用程序信息
processString = Replace(processString, "%APPPATH%", APPPath())
processString = Replace(processString, "%APPDRIVE%", Left(APPPath(), 1))
processString = Replace(processString, "%VER%", GetFileVerInfo(APPPath("QuickClip.exe"))(0))
'系统信息
processString = Replace(processString, "%Sys_ComputerName%", Sys_ComputerName)
processString = Replace(processString, "%Sys_ComputerStatus%", Sys_ComputerStatus)
processString = Replace(processString, "%Sys_SystemType%", Sys_SystemType)
processString = Replace(processString, "%Sys_Manufacturer%", Sys_Manufacturer)
processString = Replace(processString, "%Sys_Model%", Sys_Model)
processString = Replace(processString, "%Sys_TotalMemory%", Sys_TotalMemory)
processString = Replace(processString, "%Sys_Domain%", Sys_Domain)
processString = Replace(processString, "%Sys_Workgroup%", Sys_Workgroup)
processString = Replace(processString, "%Sys_Usename%", Sys_Usename)
processString = Replace(processString, "%Sys_BootupState%", Sys_BootupState)
processString = Replace(processString, "%Sys_OwnerName%", Sys_OwnerName)
processString = Replace(processString, "%Sys_CreationClassName%", Sys_CreationClassName)
processString = Replace(processString, "%Sys_Description%", Sys_Description)
'考虑加入环境变量选择器
'processString = Replace(processString, "%%",1 )
Debug.Print processString
End Function

Public Function ConcatPath(ByVal path1 As String, Optional ByVal path2 As String, Optional ByVal VerifyUnavailableChar As Boolean) As String
Dim ts1 As String, ts2 As String, i As Integer
Dim tc As String * 1
For i = 1 To Len(path1)
    tc = Mid(path1, i, 1)
    If Not (isUnavailableCharacter(tc)) Then ts1 = ts1 & tc
Next
If Not (IsMissing(path2)) Then
For i = 1 To Len(path2)
    tc = Mid(path2, i, 1)
    If Not (isUnavailableCharacter(tc)) Then ts1 = ts1 & tc
Next
Else: ts2 = ""
End If
If Mid(ts1, 2, 2) <> ":\" Then ts1 = KillSlashes(APPPath) & "\" & KillSlashes(ts1)
ConcatPath = KillSlashes(ts1) & "\" & KillSlashes(ts2)
End Function

Public Function KillSlashes(ByVal str As String) As String '干掉字符串两侧的\/
Dim ts As String
ts = Trim(str)
While Left(ts, 1) = "/" Or Left(ts, 1) = "\"
ts = Right(ts, Len(ts) - 1)
Wend
While Right(ts, 1) = "/" Or Right(ts, 1) = "\"
ts = Left(ts, Len(ts) - 1)
Wend
End Function

Public Function isUnavailableCharacter(ByVal Char As String) As Boolean '此函数没有实现！
'Public Const UnavailableCharacters = Array("/", "\", ":", "*", " ", "<", ">", "|")
Dim tc As String * 1
'tc = Left(Char, 1)
'if tc="/"then return "\"
If tc <> " " Then isUnavailableCharacter = False Else isUnavailableCharacter = True
End Function
