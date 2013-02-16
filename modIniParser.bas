Attribute VB_Name = "modINIParser"
'ini��дģ��
Option Explicit

Public iniFileName As String
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'----------------------------------ini�ļ���д�����˵�ģ�飩---------------------------------
    
    '****************************************��ȡIni�ַ���ֵ(Function)******************************************
    Function GetIniSA(ByVal SectionName As String, ByVal KeyWord As String, ByVal DefString As String) As String
    Dim ResultString As String * 144, Temp As Integer
    Dim s As String, i As Integer
    Temp% = GetPrivateProfileString(SectionName, KeyWord, "", ResultString, 144, AppProFileName(iniFileName))
    '�����ؼ��ʵ�ֵ
    If Temp% > 0 Then '�ؼ��ʵ�ֵ��Ϊ��
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
    '��ȱʡֵд��INI�ļ�
    s = DefString
    End If
    GetIniSA = s
    End Function
    
    Sub GetIniS(ByVal SectionName As String, ByVal KeyWord As String, ByRef Aim As String, Optional ByVal DefString As String)
    Aim = GetIniSA(SectionName, KeyWord, DefString)
    'Debug.Print "[" + SectionName + "]" + KeyWord + " = " + Aim
    End Sub

    '**************************************��ȡIni��ֵ(Function)***************************************************
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

    '***************************************д���ַ���ֵ(Sub)**************************************************
    Sub SetIniS(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValStr As String)
    Dim res%
    res% = WritePrivateProfileString(SectionName, KeyWord, ValStr, AppProFileName(iniFileName))
    'Debug.Print "Save - [" + SectionName + "]" + KeyWord + " = " + ValStr
    End Sub
    '****************************************д����ֵ(Sub)******************************************************
    Sub SetIniN(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValInt As Long)
    Dim res%, s$
    s$ = str$(ValInt)
    res% = WritePrivateProfileString(SectionName, KeyWord, s$, AppProFileName(iniFileName))
    'Debug.Print "Save - [" + SectionName + "]" + KeyWord + " = " & ValInt
    End Sub
    
    '���������Ѳ�֪���������һ����(keyword) ʱ
    'д��һ������ַ���ֵ�Ĺ��̣�����write����д��һ���յ�ֵʵ�ֵģ�'Sub DelIniS(ByVal SectionName As String, ByVal KeyWord As String)
    'Dim retval As Integer
    'retval = WritePrivateProfileString(SectionName, KeyWord, "", AppProFileName(iniFileName))
    'End Sub
    '��ʵ0&��ʾǰ���һ����������Ҷ�д��һ����������������section����дһ��Key��һ��������

    '***************************************���KeyWord"��"(Sub)*************************************************
    Sub DelIniKey(ByVal SectionName As String, ByVal KeyWord As String)
    Dim RetVal As Integer
    RetVal = WritePrivateProfileString(SectionName, KeyWord, 0&, AppProFileName(iniFileName))
    End Sub

    '��������section����дһ��Key��һ��������
    '**************************************��� Section"��"(Sub)***********************************************
    Sub DelIniSec(ByVal SectionName As String) '���section
    Dim RetVal As Integer
    RetVal = WritePrivateProfileString(SectionName, 0&, "", AppProFileName(iniFileName))
    End Sub

    '*************************************����Ini�ļ���(Function)***************************************************
    '����ini�ļ���
    Function AppProFileName(iniFileName)
    'AppProFileName = Trim(App.Path & "\" & iniFileName)
    AppProFileName = Trim(iniFileName)
    End Function


    '�÷�: ���� ����iniFileName="�ļ���" ����Ҫ ��ini��׺
    '�����˵������Ը�ֵ��iniFileName�Ϳ���д���¼�������������ʱд�벻ͬ��ini�ļ�(��������ļ��Ƿ��Ѵ��ڣ���ͨ���޸�������ñ�����

    'Ȼ�� DelInikey��ByVal SectionName As String, ByVal KeyWord As String�� �����
              'DelIniSec(ByVal SectionName As String)) �����
              'SetIniN(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValInt As Long) д����
              'GetIniN(ByVal SectionName As String, ByVal KeyWord As String, ByVal DefValue As Long)��ȡ��
              'SetIniS (ByVal SectionName As String, ByVal KeyWord As String, ByVal ValStr As String) д���ַ�
              'GetIniS(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValStr As String) ��ȡ�ַ�

