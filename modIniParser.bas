Attribute VB_Name = "modINIParser"
'ini读写模块
Option Explicit

Public iniFileName As String
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'----------------------------------ini文件读写（别人的模块）---------------------------------
    
    '****************************************获取Ini字符串值(Function)******************************************
    Function GetIniSA(ByVal SectionName As String, ByVal KeyWord As String, ByVal DefString As String) As String
    Dim ResultString As String * 144, Temp As Integer
    Dim s As String, i As Integer
    Temp% = GetPrivateProfileString(SectionName, KeyWord, "", ResultString, 144, AppProFileName(iniFileName))
    '检索关键词的值
    If Temp% > 0 Then '关键词的值不为空
    s = ""
    For i = 1 To 144
    If Asc(Mid$(ResultString, i, 1)) = 0 Then
    Exit For
    Else
    s = s & Mid$(ResultString, i, 1)
    End If
    Next
    Else
    Temp% = WritePrivateProfileString(SectionName, KeyWord, DefString, AppProFileName(iniFileName))
    '将缺省值写入INI文件
    s = DefString
    End If
    GetIniSA = s
    End Function
    
    Sub GetIniS(ByVal SectionName As String, ByVal KeyWord As String, ByRef Aim As String, Optional ByVal DefString As String)
    Aim = GetIniSA(SectionName, KeyWord, DefString)
    'Debug.Print "[" + SectionName + "]" + KeyWord + " = " + Aim
    End Sub

    '**************************************获取Ini数值(Function)***************************************************
    Function GetIniNA(ByVal SectionName As String, ByVal KeyWord As String, ByVal DefValue As Long) As Integer
    Dim d As Long, s As String
    d = DefValue
    GetIniNA = GetPrivateProfileInt(SectionName, KeyWord, DefValue, AppProFileName(iniFileName))
    If d <> DefValue Then
    s = "" & d
    d = WritePrivateProfileString(SectionName, KeyWord, s, AppProFileName(iniFileName))
    End If
    End Function
    
    Sub GetIniN(ByVal SectionName As String, ByVal KeyWord As String, ByRef Aim As Long, Optional ByVal DefValue As String)
    Aim = GetIniSA(SectionName, KeyWord, DefValue)
    'Debug.Print "[" + SectionName + "]" + KeyWord + " = " & Aim
    End Sub

    '***************************************写入字符串值(Sub)**************************************************
    Sub SetIniS(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValStr As String)
    Dim res%
    res% = WritePrivateProfileString(SectionName, KeyWord, ValStr, AppProFileName(iniFileName))
    'Debug.Print "Save - [" + SectionName + "]" + KeyWord + " = " + ValStr
    End Sub
    '****************************************写入数值(Sub)******************************************************
    Sub SetIniN(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValInt As Long)
    Dim res%, s$
    s$ = str$(ValInt)
    res% = WritePrivateProfileString(SectionName, KeyWord, s$, AppProFileName(iniFileName))
    'Debug.Print "Save - [" + SectionName + "]" + KeyWord + " = " & ValInt
    End Sub
    
    '这是我自已不知道怎样清除一个键(keyword) 时
    '写的一个清除字符串值的过程，是有write函数写入一个空的值实现的，'Sub DelIniS(ByVal SectionName As String, ByVal KeyWord As String)
    'Dim retval As Integer
    'retval = WritePrivateProfileString(SectionName, KeyWord, "", AppProFileName(iniFileName))
    'End Sub
    '其实0&表示前面的一个被清除，我多写了一个“”，如果是清除section就少写一个Key多一个“”。

    '***************************************清除KeyWord"键"(Sub)*************************************************
    Sub DelIniKey(ByVal SectionName As String, ByVal KeyWord As String)
    Dim RetVal As Integer
    RetVal = WritePrivateProfileString(SectionName, KeyWord, 0&, AppProFileName(iniFileName))
    End Sub

    '如果是清除section就少写一个Key多一个“”。
    '**************************************清除 Section"段"(Sub)***********************************************
    Sub DelIniSec(ByVal SectionName As String) '清除section
    Dim RetVal As Integer
    RetVal = WritePrivateProfileString(SectionName, 0&, "", AppProFileName(iniFileName))
    End Sub

    '*************************************定义Ini文件名(Function)***************************************************
    '定义ini文件名
    Function AppProFileName(iniFileName)
    'AppProFileName = Trim(App.Path & "\" & iniFileName)
    AppProFileName = Trim(iniFileName)
    End Function


    '用法: 首先 定义iniFileName="文件名" 不需要 加ini后缀
    '这就是说，你可以赋值给iniFileName就可以写入记录，而且你可以随时写入不同的ini文件(不管这个文件是否已存在），通过修改这个公用变量。

    '然后　 DelInikey（ByVal SectionName As String, ByVal KeyWord As String） 清除键
              'DelIniSec(ByVal SectionName As String)) 清除部
              'SetIniN(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValInt As Long) 写入数
              'GetIniN(ByVal SectionName As String, ByVal KeyWord As String, ByVal DefValue As Long)读取数
              'SetIniS (ByVal SectionName As String, ByVal KeyWord As String, ByVal ValStr As String) 写入字符
              'GetIniS(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValStr As String) 读取字符

